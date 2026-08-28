"""Momo's voice: synthesize speech with Microsoft neural voices (edge-tts).

Natural, soft, human — and bilingual: hi-IN-SwaraNeural speaks Hindi, English and
Hinglish in one warm voice. Synthesis only; playback is handled by the caller.

Usage:
  python momo_voice.py --voice hi-IN-SwaraNeural --rate +0% --out x.mp3 --text-b64 <base64 utf-8>

Exit codes: 0 = mp3 written, 2 = failure (caller falls back to offline SAPI).
"""
import argparse
import asyncio
import base64
import sys


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--voice", default="hi-IN-SwaraNeural")
    ap.add_argument("--rate", default="+0%")
    ap.add_argument("--out", required=True)
    ap.add_argument("--text-b64", required=True)
    args = ap.parse_args()

    try:
        text = base64.b64decode(args.text_b64).decode("utf-8")
        if not text.strip():
            sys.exit(2)
        import edge_tts

        async def run():
            tts = edge_tts.Communicate(text, voice=args.voice, rate=args.rate)
            await tts.save(args.out)

        asyncio.run(asyncio.wait_for(run(), timeout=20))
    except Exception as e:  # noqa: BLE001
        print(f"edge-tts failed: {e}", file=sys.stderr)
        sys.exit(2)


if __name__ == "__main__":
    main()
