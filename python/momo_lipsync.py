"""Real lip-sync: portrait photo + narration audio -> talking-head video (Wav2Lip).

Runs fully locally (CPU torch) using the checkpoints in %LOCALAPPDATA%/Momo/wav2lip.

Usage:
  python momo_lipsync.py --face portrait.jpg --audio narration.wav --out head.mp4
"""
import argparse
import os
import subprocess
import sys
import tempfile

import numpy as np

W2L = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Momo", "wav2lip")
sys.path.insert(0, W2L)

# ---- shim: Wav2Lip's audio.py was written for old librosa (positional args) ----
import librosa
import librosa.filters

_orig_mel = librosa.filters.mel

def _mel_shim(*args, **kwargs):
    if args:
        kwargs.setdefault("sr", args[0])
        if len(args) > 1:
            kwargs.setdefault("n_fft", args[1])
    return _orig_mel(**kwargs)

librosa.filters.mel = _mel_shim

import audio as w2l_audio          # noqa: E402  (from the wav2lip repo)
import torch                        # noqa: E402
import cv2                          # noqa: E402
from models.wav2lip import Wav2Lip  # noqa: E402

MEL_STEP = 16
FPS = 25


def load_model(path):
    ckpt = torch.load(path, map_location="cpu", weights_only=True)
    sd = ckpt["state_dict"] if "state_dict" in ckpt else ckpt
    sd = {k.replace("module.", ""): v for k, v in sd.items()}
    model = Wav2Lip()
    model.load_state_dict(sd)
    return model.eval()


def detect_face(img):
    import face_detection
    det = face_detection.FaceAlignment(face_detection.LandmarksType._2D,
                                       flip_input=False, device="cpu")
    preds = det.get_detections_for_batch(np.array([img]))
    if preds[0] is None:
        raise SystemExit("no face found in the portrait")
    x1, y1, x2, y2 = preds[0]
    # pads (top bias so the chin is included)
    y2 = min(img.shape[0], y2 + 10)
    return x1, y1, x2, y2


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--face", required=True)
    ap.add_argument("--audio", required=True, help="16k mono wav")
    ap.add_argument("--out", required=True)
    ap.add_argument("--batch", type=int, default=64)
    args = ap.parse_args()

    import imageio_ffmpeg
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()

    img = cv2.imread(args.face)
    if img is None:
        raise SystemExit("cannot read portrait: " + args.face)
    # a small portrait upscales badly through the pipeline — normalize height
    if img.shape[0] < 480:
        s = 480 / img.shape[0]
        img = cv2.resize(img, (int(img.shape[1] * s), 480))

    print("detecting face...", flush=True)
    x1, y1, x2, y2 = detect_face(img)
    face_crop = cv2.resize(img[y1:y2, x1:x2], (96, 96))

    print("computing mel spectrogram...", flush=True)
    from scipy.io import wavfile
    sr, wav = wavfile.read(args.audio)
    if wav.dtype != np.float32:
        wav = wav.astype(np.float32) / 32768.0
    if wav.ndim > 1:
        wav = wav.mean(axis=1)
    mel = w2l_audio.melspectrogram(wav)

    mel_chunks = []
    mel_idx_mult = 80.0 / FPS
    i = 0
    while True:
        start = int(i * mel_idx_mult)
        if start + MEL_STEP > mel.shape[1]:
            mel_chunks.append(mel[:, -MEL_STEP:])
            break
        mel_chunks.append(mel[:, start:start + MEL_STEP])
        i += 1
    n = len(mel_chunks)
    print(f"{n} frames ({n / FPS:.1f}s) to synthesize...", flush=True)

    model = load_model(os.path.join(W2L, "checkpoints", "wav2lip_gan.pth"))

    work = tempfile.mkdtemp(prefix="momo_lip_")
    frames_dir = os.path.join(work, "frames")
    os.makedirs(frames_dir)

    masked = face_crop.copy()
    masked[96 // 2:] = 0
    fi = 0
    with torch.no_grad():
        for b0 in range(0, n, args.batch):
            chunk = mel_chunks[b0:b0 + args.batch]
            img_batch = np.repeat(
                np.concatenate((masked, face_crop), axis=2)[None], len(chunk), axis=0)
            mel_batch = np.array(chunk)[..., None]
            img_t = torch.FloatTensor(img_batch.transpose(0, 3, 1, 2)) / 255.0
            mel_t = torch.FloatTensor(mel_batch.transpose(0, 3, 1, 2))
            pred = model(mel_t, img_t).numpy().transpose(0, 2, 3, 1) * 255.0
            for p in pred:
                frame = img.copy()
                frame[y1:y2, x1:x2] = cv2.resize(
                    p.astype(np.uint8), (x2 - x1, y2 - y1))
                cv2.imwrite(os.path.join(frames_dir, f"{fi:05d}.jpg"), frame,
                            [cv2.IMWRITE_JPEG_QUALITY, 92])
                fi += 1
            print(f"  {min(b0 + args.batch, n)}/{n}", flush=True)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    subprocess.run([ffmpeg, "-y", "-framerate", str(FPS),
                    "-i", os.path.join(frames_dir, "%05d.jpg"),
                    "-i", args.audio, "-c:v", "libx264", "-pix_fmt", "yuv420p",
                    "-c:a", "aac", "-shortest", args.out],
                   check=True, capture_output=True)
    import shutil
    shutil.rmtree(work, ignore_errors=True)
    print("TALKING HEAD READY:", args.out)


if __name__ == "__main__":
    main()
