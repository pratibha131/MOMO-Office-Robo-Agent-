"""Momo's ears — a persistent listening daemon.

Started once when Momo starts, it loads the speech models ONCE and then waits for
commands on stdin, so tapping the mic responds instantly.

Why hybrid: Vosk streams live partial text (the bubble updates as you speak) and,
with an adaptive energy detector, decides when you stopped. The FINAL text comes
from faster-whisper — far more accurate on quiet speech, Indian names (Kavya,
Pratibha), and Hindi/Hinglish. Auto-gain boosts soft voices up to 8x, so you can
speak normally instead of loudly.

stdin commands (one per line):
  listen                 open the mic and transcribe one utterance
  cancel                 stop listening, discard (also suppresses an in-flight final)
  meeting <path>         start meeting capture (mic = [you], speakers = [them])
  meeting-stop           finish meeting capture, flush transcript (async, non-blocking)
  quit                   shut down

stdout events (one JSON per line):
  {"event":"daemon-ready","device":"..."}     models loaded, ready for commands
  {"event":"listening","device":"..."}        mic open for a command
  {"event":"partial","text":"..."}            live text while you speak
  {"event":"transcribing"}                    turning the utterance into final text
  {"event":"final","text":"..."}              the recognized command
  {"event":"cancelled"}                       listening aborted, nothing recognized
  {"event":"meeting-started","transcript":"..."}
  {"event":"meeting-stopped"}
  {"event":"status","message":"..."} / {"event":"error","message":"..."}
"""
import argparse
import json
import os
import queue
import sys
import threading
import time

# Windows pipes default to the ANSI codepage; force UTF-8 so Hindi/accented text
# can't crash the daemon mid-listen.
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

SAMPLE_RATE = 16000
_emit_lock = threading.Lock()
DBG = os.environ.get("MOMO_DEBUG") == "1"
_t0 = time.monotonic()


def emit(obj):
    with _emit_lock:
        sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
        sys.stdout.flush()


def dbg(msg):
    if DBG:
        print(f"dbg {time.monotonic()-_t0:6.2f}s {msg}", file=sys.stderr, flush=True)


def fmt_elapsed(sec):
    sec = int(sec)
    return f"{sec // 60:02d}:{sec % 60:02d}"


class AutoGain:
    """Track a decaying peak so quiet speech gets amplified, loud speech doesn't clip."""

    def __init__(self):
        self.ema_peak = 4000.0

    def apply(self, np, pcm_bytes):
        """Return (boosted_bytes, raw_float_array). Detection MUST use the raw
        array: boosting lifts room noise over any fixed threshold."""
        a = np.frombuffer(pcm_bytes, dtype=np.int16).astype(np.float32)
        if a.size == 0:
            return pcm_bytes, a
        raw = a
        peak = float(np.abs(a).max())
        if peak > self.ema_peak:
            self.ema_peak = peak                                  # fast attack
        else:
            self.ema_peak = self.ema_peak * 0.998 + peak * 0.002   # slow decay
        gain = min(8.0, 26000.0 / max(self.ema_peak, 400.0))
        if gain > 1.0:
            a = np.clip(a * gain, -32767, 32767)
        return a.astype(np.int16).tobytes(), raw


class Ear:
    def __init__(self, args):
        self.args = args
        self.listening = False
        self.listen_gen = 0            # bumped by cancel: suppresses in-flight finals
        self.meeting = False
        self.meeting_gen = 0           # each meeting gets its own transcript handle
        self.gen_fh = {}               # meeting_gen -> open transcript file
        self.meeting_t0 = 0.0
        self.stream = None
        self.stream_lock = threading.RLock()
        self.cmd_q = queue.Queue()
        self.jobs = queue.Queue()
        self.bufs = {"you": bytearray(), "them": bytearray()}
        self.buf_start = {"you": 0.0, "them": 0.0}
        self.buf_lock = threading.Lock()
        self.transcript_lock = threading.Lock()
        self.whisper_lock = threading.Lock()
        self.pending = 0               # jobs enqueued but not yet fully written
        self.pending_lock = threading.Lock()
        self.cmd_priority = threading.Event()   # a command outranks meeting chunks
        self.stop_all = threading.Event()
        self.loop_stop = threading.Event()
        self.loop_thread = None
        self.device = None
        self.device_name = "default"
        self.noise_floor = 150.0       # learned across utterances, not reset each time
        self.glossary = ""
        self.you_agc = AutoGain()
        self.cmd_agc = AutoGain()

    # ---------- setup ----------
    def load(self):
        import numpy as np
        import sounddevice as sd
        self.np = np
        self.sd = sd
        dbg("numpy + sounddevice imported")

        try:
            if self.args.input_device:
                for i, d in enumerate(sd.query_devices()):
                    if d.get("max_input_channels", 0) > 0 and \
                            self.args.input_device.lower() in d.get("name", "").lower():
                        self.device = i
                        self.device_name = d["name"]
                        break
                if self.device is None:
                    emit({"event": "status",
                          "message": f"input device '{self.args.input_device}' not found, using default"})
            if self.device is None:
                self.device_name = sd.query_devices(kind="input").get("name", "default")
        except Exception:
            pass

        import vosk
        vosk.SetLogLevel(-1)
        self.vosk = vosk
        if not os.path.isdir(self.args.model):
            emit({"event": "error", "message": f"vosk model not found: {self.args.model}"})
            sys.exit(2)
        self.vmodel = vosk.Model(self.args.model)
        dbg("vosk model loaded")

        self.wm = None
        if self.args.engine == "whisper":
            try:
                from faster_whisper import WhisperModel
                try:
                    # cached model: skip the slow HuggingFace network check
                    self.wm = WhisperModel(self.args.whisper_model, device="cpu",
                                           compute_type="int8", cpu_threads=4,
                                           local_files_only=True)
                except Exception:
                    self.wm = WhisperModel(self.args.whisper_model, device="cpu",
                                           compute_type="int8", cpu_threads=4)
                dbg("whisper model loaded")
            except Exception as e:  # noqa: BLE001
                emit({"event": "status",
                      "message": f"whisper unavailable, using vosk only: {e}"})

    # ---------- transcript (per meeting generation) ----------
    def tlog(self, gen, tag, txt):
        fh = self.gen_fh.get(gen)
        if fh and txt:
            with self.transcript_lock:
                if fh.closed:
                    return
                try:
                    fh.write(f"[{tag}] {txt}\n")
                    fh.flush()
                except ValueError:
                    pass

    # ---------- mic stream (shared by command + meeting) ----------
    def _mic_cb(self, indata, frames, t, status):  # noqa: ARG001
        raw = bytes(indata)
        if self.meeting:
            b, _ = self.you_agc.apply(self.np, raw)
            self.push_audio("you", b)
        if self.listening:
            self.cmd_q.put(raw)

    def _feed_file(self):
        """Diagnostics: stream a WAV through the same path the microphone uses."""
        import wave
        np = self.np
        try:
            with wave.open(self.args.test_feed) as w:
                sr = w.getframerate()
                pcm = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16)
            if sr != SAMPLE_RATE:                       # cheap resample to 16k
                idx = (np.arange(int(len(pcm) * SAMPLE_RATE / sr)) * sr / SAMPLE_RATE)
                pcm = pcm[np.clip(idx.astype(np.int64), 0, len(pcm) - 1)]
            block = 4000
            for i in range(0, len(pcm), block):
                if self.stop_all.is_set() or self.stream != "feed":
                    return
                chunk = pcm[i:i + block]
                if len(chunk) < block:
                    chunk = np.pad(chunk, (0, block - len(chunk)))
                self._mic_cb(chunk.tobytes(), block, None, None)
                time.sleep(block / SAMPLE_RATE)
            silence = np.zeros(block, dtype=np.int16)
            for _ in range(40):
                if self.stop_all.is_set() or self.stream != "feed":
                    return
                self._mic_cb(silence.tobytes(), block, None, None)
                time.sleep(block / SAMPLE_RATE)
        finally:
            # allow the NEXT listen/meeting to start a fresh feed pass
            with self.stream_lock:
                if self.stream == "feed":
                    self.stream = None

    def ensure_stream(self):
        with self.stream_lock:
            if self.stream is not None:
                return True
            if self.args.test_feed:
                self.stream = "feed"
                threading.Thread(target=self._feed_file, daemon=True).start()
                dbg("test feed started")
                return True
            try:
                self.stream = self.sd.RawInputStream(
                    samplerate=SAMPLE_RATE, blocksize=4000, dtype="int16", channels=1,
                    callback=self._mic_cb, device=self.device)
                self.stream.start()
                dbg("mic stream opened")
                return True
            except Exception as e:  # noqa: BLE001
                self.stream = None
                emit({"event": "error", "message": f"microphone unavailable: {e}"})
                return False

    def maybe_close_stream(self):
        with self.stream_lock:
            if self.stream is not None and not self.listening and not self.meeting:
                if self.stream != "feed":
                    try:
                        self.stream.stop()
                        self.stream.close()
                    except Exception:
                        pass
                self.stream = None
                dbg("mic stream closed")

    # ---------- whisper helper ----------
    def load_glossary(self):
        """Names + domain jargon prime the model — the difference between
        'काविय बारदवाज' and 'Kavya Bhardwaj' in a Hinglish meeting."""
        try:
            if self.args.glossary_file and os.path.exists(self.args.glossary_file):
                with open(self.args.glossary_file, encoding="utf-8") as f:
                    self.glossary = " ".join(f.read().split())[:900]
        except Exception:
            pass

    def transcribe(self, audio):
        if self.wm is None:
            return ""
        segments, _info = self.wm.transcribe(audio, language=None,
                                             vad_filter=True, beam_size=1,
                                             initial_prompt=self.glossary or None)
        return " ".join(s.text.strip() for s in segments).strip()

    # ---------- command listening ----------
    def start_listening(self):
        if self.listening:
            return
        while not self.cmd_q.empty():
            try:
                self.cmd_q.get_nowait()
            except queue.Empty:
                break
        self.cmd_agc = AutoGain()
        self.listening = True
        if not self.ensure_stream():
            self.listening = False
            emit({"event": "cancelled"})
            return
        emit({"event": "listening", "device": self.device_name})

    def stop_listening(self, cancelled):
        was = self.listening
        self.listening = False
        self.listen_gen += 1           # suppresses a final still being transcribed
        self.maybe_close_stream()
        if was and cancelled:
            emit({"event": "cancelled"})

    def command_worker(self):
        np = self.np
        rec = self.vosk.KaldiRecognizer(self.vmodel, SAMPLE_RATE)
        rec.SetWords(False)
        utter = bytearray()
        texts = []
        heard = False
        last_voice = time.monotonic()
        PRE_ROLL = SAMPLE_RATE * 2 * 2
        MAX_UTTER = SAMPLE_RATE * 2 * 30      # 30s cap for a spoken command
        was_listening = False

        def reset():
            nonlocal rec, utter, texts, heard
            rec = self.vosk.KaldiRecognizer(self.vmodel, SAMPLE_RATE)
            rec.SetWords(False)
            utter = bytearray()
            texts = []
            heard = False

        while not self.stop_all.is_set():
            if not self.listening:
                if was_listening:
                    reset()
                    was_listening = False
                time.sleep(0.05)
                continue
            was_listening = True
            try:
                raw = self.cmd_q.get(timeout=0.2)
            except queue.Empty:
                if heard and (time.monotonic() - last_voice) * 1000 >= self.args.silence_ms:
                    self._finalize(rec, utter, texts)
                    reset()
                continue

            b, arr = self.cmd_agc.apply(np, raw)
            now = time.monotonic()
            utter.extend(b)
            if not heard and len(utter) > PRE_ROLL:
                utter = bytearray(utter[-PRE_ROLL:])   # rolling pre-speech buffer

            # adaptive energy detector on RAW audio. The floor is learned across
            # utterances; it only adapts DOWNWARD-or-slowly so sustained soft speech
            # can't drag it up over itself.
            rms = float(np.sqrt(np.mean(arr * arr))) if arr.size else 0.0
            if rms < self.noise_floor * 1.5:
                self.noise_floor = self.noise_floor * 0.97 + rms * 0.03
            elif not heard:
                self.noise_floor = self.noise_floor * 0.995 + rms * 0.005
            if rms > max(120.0, self.noise_floor * 3.5):
                heard = True
                last_voice = now

            if rec.AcceptWaveform(b):
                txt = json.loads(rec.Result()).get("text", "")
                if txt:
                    texts.append(txt)
                    heard = True
                    last_voice = now
                    emit({"event": "partial", "text": " ".join(texts)})
            else:
                partial = json.loads(rec.PartialResult()).get("partial", "")
                if partial:
                    heard = True
                    last_voice = now
                    emit({"event": "partial", "text": " ".join(texts + [partial]).strip()})

            if heard and ((now - last_voice) * 1000 >= self.args.silence_ms
                          or len(utter) >= MAX_UTTER):
                self._finalize(rec, utter, texts)
                reset()

    def _finalize(self, rec, utter, texts):
        """Produce the final command text (whisper preferred, vosk as fallback)."""
        np = self.np
        gen = self.listen_gen          # a cancel during transcription bumps this
        tail = json.loads(rec.FinalResult()).get("text", "")
        if tail:
            texts = texts + [tail]
        text = ""
        if self.wm is not None and len(utter) > SAMPLE_RATE:      # >0.5s of audio
            emit({"event": "transcribing"})
            self.cmd_priority.set()                                # outrank meeting chunks
            try:
                with self.whisper_lock:
                    audio = np.frombuffer(bytes(utter), dtype=np.int16).astype(np.float32) / 32768.0
                    text = self.transcribe(audio)
                dbg(f"whisper command: {text!r}")
            except Exception as e:  # noqa: BLE001
                emit({"event": "status", "message": f"whisper failed, using vosk: {e}"})
            finally:
                self.cmd_priority.clear()
        if self.listen_gen != gen:
            dbg("final suppressed: cancelled during transcription")
            return
        if not text:
            text = " ".join(t for t in texts if t).strip()
        self.listening = False
        self.maybe_close_stream()
        if text:
            emit({"event": "final", "text": text})
        else:
            emit({"event": "cancelled"})

    # ---------- meeting capture ----------
    def push_audio(self, tag, pcm_bytes):
        with self.buf_lock:
            if not self.bufs[tag]:
                self.buf_start[tag] = time.monotonic() - self.meeting_t0
            self.bufs[tag].extend(pcm_bytes)
            if len(self.bufs[tag]) >= self.args.chunk_sec * SAMPLE_RATE * 2:
                self._enqueue(tag, self.buf_start[tag], bytes(self.bufs[tag]))
                self.bufs[tag] = bytearray()

    def flush_bufs(self):
        with self.buf_lock:
            for tag in ("you", "them"):
                if self.bufs[tag]:
                    self._enqueue(tag, self.buf_start[tag], bytes(self.bufs[tag]))
                    self.bufs[tag] = bytearray()

    def _enqueue(self, tag, started, pcm):
        # pending is incremented BEFORE the job is visible to the worker, so
        # "queue empty and pending==0" really means everything is written out
        with self.pending_lock:
            self.pending += 1
        self.jobs.put((self.meeting_gen, tag, started, pcm))

    def meeting_worker(self):
        np = self.np
        while not self.stop_all.is_set():
            try:
                job = self.jobs.get(timeout=0.3)
            except queue.Empty:
                continue
            if job is None:
                continue
            gen, tag, started, pcm = job
            try:
                if gen not in self.gen_fh:
                    continue           # straggler from a closed meeting: drop it
                while self.cmd_priority.is_set():   # let a spoken command jump ahead
                    time.sleep(0.05)
                audio = np.frombuffer(pcm, dtype=np.int16).astype(np.float32) / 32768.0
                dbg(f"job {tag} {audio.size/SAMPLE_RATE:.1f}s peak={float(np.abs(audio).max()):.4f}")
                if audio.size < SAMPLE_RATE // 2 or np.abs(audio).max() < 0.004:
                    continue
                with self.whisper_lock:
                    text = self.transcribe(audio)
                if text:
                    self.tlog(gen, f"{tag} +{fmt_elapsed(started)}", text)
            except Exception as e:  # noqa: BLE001
                emit({"event": "error", "message": f"transcribe chunk failed: {e}"})
            finally:
                with self.pending_lock:
                    self.pending -= 1

    def loopback_worker(self):
        np = self.np
        try:
            import soundcard as sc
        except Exception as e:  # noqa: BLE001
            emit({"event": "error",
                  "message": f"loopback unavailable, meeting notes use mic only: {e}"})
            return
        last_err = 0.0
        while not self.loop_stop.is_set() and not self.stop_all.is_set():
            try:
                spk = sc.default_speaker()
                loop_mic = sc.get_microphone(id=str(spk.name), include_loopback=True)
                with loop_mic.recorder(samplerate=SAMPLE_RATE, channels=1, blocksize=4000) as lr:
                    last_dev_check = time.monotonic()
                    while not self.loop_stop.is_set() and not self.stop_all.is_set():
                        data = lr.record(numframes=None)
                        if data is not None and len(data):
                            d = data if data.ndim == 1 else data[:, 0]
                            pcm = (np.clip(d, -1, 1) * 32767).astype(np.int16).tobytes()
                            self.push_audio("them", pcm)
                        else:
                            self.loop_stop.wait(0.05)
                        if time.monotonic() - last_dev_check > 5:
                            last_dev_check = time.monotonic()
                            try:
                                if sc.default_speaker().name != spk.name:
                                    break     # follow headset/speaker switches
                            except Exception:
                                break
            except Exception as e:  # noqa: BLE001
                if time.monotonic() - last_err > 30:
                    last_err = time.monotonic()
                    emit({"event": "error", "message": f"loopback hiccup, retrying: {e}"})
                self.loop_stop.wait(1.0)

    def start_meeting(self, path):
        if self.meeting:
            return
        self.load_glossary()           # fresh names/terms for this meeting
        try:
            d = os.path.dirname(path)
            if d:
                os.makedirs(d, exist_ok=True)
            fh = open(path, "a", encoding="utf-8")
        except Exception as e:  # noqa: BLE001
            emit({"event": "error", "message": f"cannot open transcript: {e}"})
            return
        self.meeting_gen += 1
        self.gen_fh[self.meeting_gen] = fh
        with self.transcript_lock:
            fh.write(f"\n--- capture started {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
            fh.flush()
        self.meeting_t0 = time.monotonic()
        self.you_agc = AutoGain()
        self.meeting = True
        if not self.ensure_stream():
            self.meeting = False
            self.gen_fh.pop(self.meeting_gen, None)
            try:
                fh.close()
            except Exception:
                pass
            emit({"event": "error", "message": "meeting capture could not start (no microphone)"})
            return
        self.loop_stop.clear()
        self.loop_thread = threading.Thread(target=self.loopback_worker, daemon=True)
        self.loop_thread.start()
        emit({"event": "meeting-started", "transcript": path})

    def stop_meeting(self):
        if not self.meeting:
            return
        gen = self.meeting_gen
        self.meeting = False
        self.loop_stop.set()
        if self.loop_thread:
            self.loop_thread.join(timeout=3.0)
            self.loop_thread = None
        self.maybe_close_stream()
        self.flush_bufs()
        emit({"event": "status", "message": "finishing transcription…"})
        # finish asynchronously so stdin commands (listen/quit) stay responsive
        threading.Thread(target=self._finish_stop, args=(gen,), daemon=True).start()

    def _finish_stop(self, gen):
        deadline = time.monotonic() + 300
        while time.monotonic() < deadline and not self.stop_all.is_set():
            with self.pending_lock:
                busy = self.pending
            if self.jobs.empty() and busy == 0:
                break
            time.sleep(0.2)
        fh = self.gen_fh.pop(gen, None)   # after pop, stragglers are dropped by gen check
        if fh:
            with self.transcript_lock:
                try:
                    fh.write(f"--- capture ended {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
                    fh.close()
                except Exception:
                    pass
        emit({"event": "meeting-stopped"})

    # ---------- main loop ----------
    def run(self):
        self.load()
        self.load_glossary()
        threading.Thread(target=self.command_worker, daemon=True).start()
        threading.Thread(target=self.meeting_worker, daemon=True).start()
        emit({"event": "daemon-ready", "device": self.device_name})
        dbg("daemon ready")

        buf = b""
        while True:
            try:
                data = os.read(0, 4096)     # raw read: no TextIOWrapper lock
            except Exception:
                break
            if not data:
                break
            buf += data
            while b"\n" in buf:
                line, buf = buf.split(b"\n", 1)
                cmd = line.decode("utf-8", "replace").strip().lower()
                if not cmd:
                    continue
                dbg(f"cmd: {cmd}")
                if cmd == "listen":
                    self.start_listening()
                elif cmd == "cancel":
                    self.stop_listening(cancelled=True)
                elif cmd.startswith("meeting-stop"):
                    self.stop_meeting()
                elif cmd.startswith("meeting "):
                    self.start_meeting(line.decode("utf-8", "replace").strip()[8:].strip())
                elif cmd == "quit":
                    self.shutdown()
                    return
        self.shutdown()

    def shutdown(self):
        self.listening = False
        if self.meeting:
            self.stop_meeting()
        # give in-flight transcript flushes a moment (a hard kill may follow anyway;
        # committed lines are already flushed line-by-line)
        deadline = time.monotonic() + 8
        while self.gen_fh and time.monotonic() < deadline:
            time.sleep(0.2)
        self.stop_all.set()
        self.maybe_close_stream()


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--model", required=True, help="path to vosk model dir")
    ap.add_argument("--silence-ms", type=int, default=1800)
    ap.add_argument("--engine", choices=["whisper", "vosk"], default="whisper")
    ap.add_argument("--whisper-model", default="small")
    ap.add_argument("--chunk-sec", type=int, default=20)
    ap.add_argument("--input-device", default=None)
    ap.add_argument("--glossary-file", default=None,
                    help="text file of names/jargon used to prime transcription")
    ap.add_argument("--test-feed", default=None,
                    help="diagnostics: play a 16-bit mono WAV into the pipeline "
                         "instead of opening the microphone")
    args = ap.parse_args()
    try:
        Ear(args).run()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
