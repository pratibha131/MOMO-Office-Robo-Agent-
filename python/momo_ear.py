"""Momo's ears: microphone / system-audio capture + offline speech-to-text.

Command mode (default, --command-engine whisper):
    Hybrid listener. Vosk streams LIVE partials (so the bubble shows text as you
    speak) and, together with an adaptive energy detector, decides when the
    utterance ended. The final command text comes from faster-whisper — far more
    accurate, quiet-voice tolerant, and multilingual (English/Hindi/Hinglish,
    names like Kavya/Pratibha). Auto-gain boosts soft speech up to 8x.
    --command-engine vosk falls back to pure vosk streaming.

Meeting mode (--meeting, --engine whisper):
    Chunked multilingual transcription: mic = [you] (auto-gain), system loopback
    = [them]; chunks transcribed in a worker and appended to --transcript with
    elapsed timestamps. --engine vosk = legacy English-only streaming.

Emits JSON lines on stdout:
  {"event":"ready","device":"..."}      capture running
  {"event":"partial","text":"..."}      live partial (command mode)
  {"event":"final","text":"..."}        finalized utterance (command mode)
  {"event":"status","message":"..."}    progress info
  {"event":"error","message":"..."}

Graceful stop: send the line "stop" on stdin (or close stdin).
"""
import argparse
import json
import os
import queue
import sys
import threading
import time

# Windows pipes default to the ANSI codepage; force UTF-8 so Hindi/accented text
# can't crash the process mid-listen.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

_emit_lock = threading.Lock()


def emit(obj):
    with _emit_lock:
        sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
        sys.stdout.flush()


def fmt_elapsed(sec):
    sec = int(sec)
    return f"{sec // 60:02d}:{sec % 60:02d}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", required=True, help="path to vosk model dir")
    ap.add_argument("--silence-ms", type=int, default=1500)
    ap.add_argument("--meeting", action="store_true")
    ap.add_argument("--transcript", default=None)
    ap.add_argument("--engine", choices=["whisper", "vosk"], default="whisper",
                    help="meeting-mode STT engine")
    ap.add_argument("--command-engine", choices=["whisper", "vosk"], default="whisper",
                    help="command-mode final-text engine")
    ap.add_argument("--whisper-model", default="small")
    ap.add_argument("--chunk-sec", type=int, default=45)
    ap.add_argument("--input-device", default=None,
                    help="substring of the input device name (default: system default)")
    args = ap.parse_args()

    DBG = os.environ.get("MOMO_DEBUG") == "1"
    _boot0 = time.monotonic()

    def bootdbg(msg):
        if DBG:
            print(f"boot {time.monotonic()-_boot0:5.2f}s {msg}", file=sys.stderr, flush=True)

    bootdbg("args parsed")
    if args.meeting and not args.transcript:
        emit({"event": "error", "message": "--meeting requires --transcript"})
        sys.exit(2)

    try:
        import numpy as np
        import sounddevice as sd
    except Exception as e:  # noqa: BLE001
        emit({"event": "error", "message": f"missing python deps: {e}"})
        sys.exit(2)
    bootdbg("numpy + sounddevice imported")

    sample_rate = 16000
    use_whisper_meeting = args.meeting and args.engine == "whisper"
    use_whisper_command = (not args.meeting) and args.command_engine == "whisper"

    # ---- input device ----
    device = None
    device_name = "default"
    try:
        if args.input_device:
            for i, d in enumerate(sd.query_devices()):
                if d.get("max_input_channels", 0) > 0 and \
                        args.input_device.lower() in d.get("name", "").lower():
                    device = i
                    device_name = d["name"]
                    break
            if device is None:
                emit({"event": "status",
                      "message": f"input device '{args.input_device}' not found, using default"})
        if device is None:
            di = sd.query_devices(kind="input")
            device_name = di.get("name", "default")
    except Exception:
        pass

    # ---- vosk (partials/VAD in command mode; legacy meeting engine) ----
    vosk = None
    model = None
    if not use_whisper_meeting:
        try:
            import vosk as _vosk
            vosk = _vosk
        except Exception as e:  # noqa: BLE001
            emit({"event": "error", "message": f"missing python deps: {e}"})
            sys.exit(2)
        if not os.path.isdir(args.model):
            emit({"event": "error", "message": f"vosk model not found: {args.model}"})
            sys.exit(2)
        vosk.SetLogLevel(-1)
        try:
            model = vosk.Model(args.model)
        except Exception as e:  # noqa: BLE001
            emit({"event": "error", "message": f"failed to load speech model: {e}"})
            sys.exit(2)
        bootdbg("vosk model loaded")

    def load_whisper():
        """Import + load in the MAIN thread (thread-side imports deadlock on
        Windows piped stdio), preferring the local cache (network check is slow)."""
        from faster_whisper import WhisperModel
        try:
            wm = WhisperModel(args.whisper_model, device="cpu",
                              compute_type="int8", cpu_threads=4, local_files_only=True)
        except Exception:
            wm = WhisperModel(args.whisper_model, device="cpu",
                              compute_type="int8", cpu_threads=4)
        return wm

    stop_event = threading.Event()

    def stdin_watcher():
        # RAW os.read: iterating sys.stdin from a thread holds the TextIOWrapper
        # lock while blocked, deadlocking main-thread imports on piped stdio.
        try:
            while True:
                data = os.read(0, 1024)
                if not data or b"stop" in data.lower():
                    break
        except Exception:
            pass
        stop_event.set()

    def start_stdin_watcher():
        # started LAST, after all imports/model loads
        threading.Thread(target=stdin_watcher, daemon=True).start()
        bootdbg("stdin watcher started")

    # ---- transcript sink (thread-safe) ----
    transcript_fh = None
    transcript_lock = threading.Lock()
    if args.transcript:
        d = os.path.dirname(args.transcript)
        if d:
            os.makedirs(d, exist_ok=True)
        transcript_fh = open(args.transcript, "a", encoding="utf-8")
        with transcript_lock:
            transcript_fh.write(
                f"\n--- capture started {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
            transcript_fh.flush()

    def tlog(tag, txt):
        if transcript_fh and txt:
            with transcript_lock:
                if transcript_fh.closed:
                    return  # straggling thread after shutdown: drop, don't crash
                try:
                    transcript_fh.write(f"[{tag}] {txt}\n")
                    transcript_fh.flush()
                except ValueError:
                    pass

    bootdbg("transcript opened")
    t_start = time.monotonic()

    def dbg(msg):
        if DBG:
            emit({"event": "status", "message": f"dbg {time.monotonic()-t_start:.1f}s {msg}"})

    # ---- auto-gain: boost quiet voices (up to 8x) so soft speech is heard ----
    def make_agc():
        return {"ema_peak": 4000.0}

    def boost(pcm_bytes, agc):
        a = np.frombuffer(pcm_bytes, dtype=np.int16).astype(np.float32)
        if a.size == 0:
            return pcm_bytes, a
        peak = float(np.abs(a).max())
        if peak > agc["ema_peak"]:
            agc["ema_peak"] = peak                      # attack: follow loud speech fast
        else:
            agc["ema_peak"] = agc["ema_peak"] * 0.998 + peak * 0.002   # slow decay
        gain = min(8.0, 26000.0 / max(agc["ema_peak"], 400.0))
        if gain > 1.0:
            a = np.clip(a * gain, -32767, 32767)
        return a.astype(np.int16).tobytes(), a

    # ================= WHISPER MEETING MODE =================
    if use_whisper_meeting:
        try:
            wm = load_whisper()
            emit({"event": "status", "message": "whisper model ready"})
        except Exception as e:  # noqa: BLE001
            emit({"event": "error", "message": f"whisper unavailable: {e}"})
            sys.exit(2)
        bootdbg("whisper loaded")

        jobs: "queue.Queue" = queue.Queue()
        bufs = {"you": bytearray(), "them": bytearray()}
        buf_start = {"you": 0.0, "them": 0.0}
        buf_lock = threading.Lock()
        you_agc = make_agc()

        def push_audio(tag, pcm_bytes):
            with buf_lock:
                if not bufs[tag]:
                    buf_start[tag] = time.monotonic() - t_start
                bufs[tag].extend(pcm_bytes)
                if len(bufs[tag]) >= args.chunk_sec * sample_rate * 2:
                    jobs.put((tag, buf_start[tag], bytes(bufs[tag])))
                    bufs[tag] = bytearray()

        def flush_bufs():
            with buf_lock:
                for tag in ("you", "them"):
                    if bufs[tag]:
                        jobs.put((tag, buf_start[tag], bytes(bufs[tag])))
                        bufs[tag] = bytearray()

        def whisper_worker():
            dbg("worker thread started")
            while True:
                job = jobs.get()
                if job is None:
                    break
                tag, started, pcm = job
                try:
                    audio = np.frombuffer(pcm, dtype=np.int16).astype(np.float32) / 32768.0
                    dbg(f"job {tag} {audio.size/sample_rate:.1f}s peak={float(np.abs(audio).max()):.4f}")
                    if audio.size < sample_rate // 2 or np.abs(audio).max() < 0.004:
                        continue  # sub-half-second or silent chunk: skip
                    segments, info = wm.transcribe(
                        audio, language=None, vad_filter=True, beam_size=1)
                    text = " ".join(s.text.strip() for s in segments).strip()
                    dbg(f"transcribed {tag}: {len(text)} chars")
                    if text:
                        tlog(f"{tag} +{fmt_elapsed(started)}", text)
                except Exception as e:  # noqa: BLE001
                    emit({"event": "error", "message": f"transcribe chunk failed: {e}"})

        worker = threading.Thread(target=whisper_worker, daemon=True)
        worker.start()
        bootdbg("worker started")

        def mic_cb(indata, frames, t, status):  # noqa: ARG001
            b, _ = boost(bytes(indata), you_agc)
            push_audio("you", b)

        try:
            stream = sd.RawInputStream(samplerate=sample_rate, blocksize=4000,
                                       dtype="int16", channels=1,
                                       callback=mic_cb, device=device)
            stream.start()
        except Exception as e:  # noqa: BLE001
            emit({"event": "error", "message": f"microphone unavailable: {e}"})
            sys.exit(2)
        bootdbg("mic stream started")

        def run_loopback():
            try:
                import soundcard as sc
            except Exception as e:  # noqa: BLE001
                emit({"event": "error",
                      "message": f"loopback unavailable, meeting notes use mic only: {e}"})
                return
            last_err = 0.0
            while not stop_event.is_set():
                try:
                    spk = sc.default_speaker()
                    loop_mic = sc.get_microphone(id=str(spk.name), include_loopback=True)
                    with loop_mic.recorder(samplerate=sample_rate, channels=1,
                                           blocksize=4000) as lr:
                        last_dev_check = time.monotonic()
                        while not stop_event.is_set():
                            data = lr.record(numframes=None)  # non-blocking-ish
                            if data is not None and len(data):
                                dd = data if data.ndim == 1 else data[:, 0]
                                pcm = (np.clip(dd, -1, 1) * 32767).astype(np.int16).tobytes()
                                push_audio("them", pcm)
                            else:
                                stop_event.wait(0.05)
                            if time.monotonic() - last_dev_check > 5:
                                last_dev_check = time.monotonic()
                                try:
                                    if sc.default_speaker().name != spk.name:
                                        break  # follow default-device switches
                                except Exception:
                                    break
                except Exception as e:  # noqa: BLE001
                    if time.monotonic() - last_err > 30:
                        last_err = time.monotonic()
                        emit({"event": "error", "message": f"loopback hiccup, retrying: {e}"})
                    stop_event.wait(1.0)

        loop_thread = threading.Thread(target=run_loopback, daemon=True)
        loop_thread.start()
        bootdbg("loopback thread started")
        start_stdin_watcher()

        emit({"event": "ready", "device": device_name})
        while not stop_event.is_set():
            stop_event.wait(0.5)

        try:
            stream.stop()
        except Exception:
            pass
        loop_thread.join(timeout=2.0)
        flush_bufs()
        jobs.put(None)
        emit({"event": "status", "message": "finishing transcription…"})
        worker.join(timeout=240.0)
        if transcript_fh:
            with transcript_lock:
                transcript_fh.write(
                    f"--- capture ended {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
                transcript_fh.close()
        return

    # ================= COMMAND MODE (hybrid) & VOSK MEETING =================
    wm = None
    if use_whisper_command:
        try:
            wm = load_whisper()
            emit({"event": "status", "message": "whisper model ready"})
        except Exception as e:  # noqa: BLE001
            emit({"event": "status",
                  "message": f"whisper unavailable, falling back to vosk: {e}"})
            use_whisper_command = False
        bootdbg("whisper loaded (command)")

    rec = vosk.KaldiRecognizer(model, sample_rate)
    rec.SetWords(False)

    audio_q: "queue.Queue[bytes]" = queue.Queue()

    def mic_cb(indata, frames, t, status):  # noqa: ARG001
        audio_q.put(bytes(indata))

    try:
        stream = sd.RawInputStream(samplerate=sample_rate, blocksize=4000, dtype="int16",
                                   channels=1, callback=mic_cb, device=device)
        stream.start()
    except Exception as e:  # noqa: BLE001
        emit({"event": "error", "message": f"microphone unavailable: {e}"})
        sys.exit(2)
    bootdbg("mic stream started")

    loop_thread = None
    if args.meeting:  # legacy vosk meeting engine
        def run_loopback():
            try:
                import soundcard as sc
            except Exception as e:  # noqa: BLE001
                emit({"event": "error",
                      "message": f"loopback unavailable, meeting notes use mic only: {e}"})
                return
            rec2 = vosk.KaldiRecognizer(model, sample_rate)
            last_err = 0.0
            them_heard = False
            them_last = time.monotonic()
            while not stop_event.is_set():
                try:
                    spk = sc.default_speaker()
                    loop_mic = sc.get_microphone(id=str(spk.name), include_loopback=True)
                    with loop_mic.recorder(samplerate=sample_rate, channels=1,
                                           blocksize=4000) as lr:
                        last_dev_check = time.monotonic()
                        while not stop_event.is_set():
                            data = lr.record(numframes=None)
                            now = time.monotonic()
                            if data is not None and len(data):
                                dd = data if data.ndim == 1 else data[:, 0]
                                pcm = (np.clip(dd, -1, 1) * 32767).astype(np.int16).tobytes()
                                if rec2.AcceptWaveform(pcm):
                                    tlog("them", json.loads(rec2.Result()).get("text", ""))
                                    them_heard = False
                                elif json.loads(rec2.PartialResult()).get("partial", ""):
                                    them_heard = True
                                    them_last = now
                            else:
                                stop_event.wait(0.05)
                            if them_heard and (now - them_last) * 1000 >= args.silence_ms:
                                tlog("them", json.loads(rec2.FinalResult()).get("text", ""))
                                rec2 = vosk.KaldiRecognizer(model, sample_rate)
                                them_heard = False
                            if now - last_dev_check > 5:
                                last_dev_check = now
                                try:
                                    if sc.default_speaker().name != spk.name:
                                        break
                                except Exception:
                                    break
                except Exception as e:  # noqa: BLE001
                    if time.monotonic() - last_err > 30:
                        last_err = time.monotonic()
                        emit({"event": "error", "message": f"loopback hiccup, retrying: {e}"})
                    stop_event.wait(1.0)
            try:
                tlog("them", json.loads(rec2.FinalResult()).get("text", ""))
            except Exception:
                pass

        loop_thread = threading.Thread(target=run_loopback, daemon=True)
        loop_thread.start()

    start_stdin_watcher()
    emit({"event": "ready", "device": device_name})

    agc = make_agc()
    utter = bytearray()             # boosted PCM of the current utterance
    vosk_texts = []                 # vosk segment texts (fallback if whisper fails)
    heard_any = False
    last_voice = time.monotonic()
    noise_floor = 300.0             # adaptive RMS noise floor
    PRE_ROLL = sample_rate * 2 * 3  # keep 3s of audio from before speech starts
    MAX_UTTER = sample_rate * 2 * 90

    def finalize_utterance():
        nonlocal rec, utter, vosk_texts, heard_any
        tail = json.loads(rec.FinalResult()).get("text", "")
        if tail:
            vosk_texts.append(tail)
        text = ""
        if use_whisper_command and len(utter) > sample_rate:  # >0.5s
            try:
                emit({"event": "status", "message": "transcribing"})
                audio = np.frombuffer(bytes(utter), dtype=np.int16).astype(np.float32) / 32768.0
                segments, info = wm.transcribe(audio, language=None,
                                               vad_filter=True, beam_size=1)
                text = " ".join(s.text.strip() for s in segments).strip()
                dbg(f"whisper final: {text!r}")
            except Exception as e:  # noqa: BLE001
                emit({"event": "status", "message": f"whisper failed, using vosk: {e}"})
        if not text:
            text = " ".join(t for t in vosk_texts if t).strip()
        if args.meeting:
            tlog("you", text)
        elif text:
            emit({"event": "final", "text": text})
        rec = vosk.KaldiRecognizer(model, sample_rate)
        utter = bytearray()
        vosk_texts = []
        heard_any = False

    while not stop_event.is_set():
        try:
            data = audio_q.get(timeout=0.25)
        except queue.Empty:
            continue
        b, arr = boost(data, agc)
        now = time.monotonic()
        utter.extend(b)
        if not heard_any and len(utter) > PRE_ROLL:
            utter = utter[-PRE_ROLL:]           # rolling pre-roll while waiting for speech

        # adaptive energy detector: catches soft speech vosk misses
        rms = float(np.sqrt(np.mean(arr * arr))) if arr.size else 0.0
        if rms < noise_floor * 2:
            noise_floor = noise_floor * 0.98 + rms * 0.02
        if rms > max(600.0, noise_floor * 3):
            heard_any = True
            last_voice = now

        if rec.AcceptWaveform(b):
            txt = json.loads(rec.Result()).get("text", "")
            if txt:
                vosk_texts.append(txt)
                heard_any = True
                last_voice = now
                if not args.meeting:
                    emit({"event": "partial", "text": " ".join(vosk_texts)})
        else:
            partial = json.loads(rec.PartialResult()).get("partial", "")
            if partial:
                heard_any = True
                last_voice = now
                if not args.meeting:
                    emit({"event": "partial",
                          "text": " ".join(vosk_texts + [partial]).strip()})

        if heard_any and ((now - last_voice) * 1000 >= args.silence_ms
                          or len(utter) >= MAX_UTTER):
            finalize_utterance()

    # graceful shutdown: flush whatever was in flight
    try:
        stream.stop()
    except Exception:
        pass
    try:
        if heard_any or len(utter) > sample_rate:
            finalize_utterance()
    except Exception:
        pass
    if loop_thread:
        loop_thread.join(timeout=2.0)
    if transcript_fh:
        with transcript_lock:
            transcript_fh.write(
                f"--- capture ended {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
            transcript_fh.close()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
