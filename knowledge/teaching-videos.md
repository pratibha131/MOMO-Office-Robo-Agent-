# Teaching videos (narrated MP4, professional)  — python/momo_video.py

The user asks: "make a video that teaches me <topic> / this deck". Two modes:

## MODE A — user gave a PPTX (USE THEIR EXACT SLIDES — never re-create generic scenes)
1. **Export the deck's slides** via PowerPoint COM:
   `$pres.Slides.Item($i).Export("$dir\sN.png", "PNG", 1280, 720)` and extract each
   slide's text (python-pptx) to write narration.
2. **Storyboard with `"type": "slide"` scenes** — one per slide, in deck order:
   ```json
   {"type": "slide", "image": "ai_slides/s2.png", "zoom": "in|out",
    "narration": "teacher-style walkthrough of THIS slide's content",
    "popup": {"at": 9.0, "title": "Try it yourself", "w": 590,
              "lines": ["> user prompt", "AI: response", "..."]}}
   ```
   - The engine adds **Ken Burns motion** (slow zoom, alternate in/out per slide)
     and **crossfade cuts** between scenes automatically — professional editing.
   - **BE SMART about popups**: analyze each slide; when content implies steps, a
     demo, or "how to use" (e.g. prompt→response workflows), add a `popup` — a
     terminal-style window animates in and TYPES a live demo (lines starting with
     "> " render as user prompts in green, others as responses). Time `at` so the
     narration is discussing that part. Don't popup every slide — only where a
     demo genuinely teaches.
3. Narration: warm teacher voice, 3-5 sentences per slide, reference what's ON the
   slide ("watch the little demo on screen"), never read the slide verbatim.

## MODE B — no deck given: design scenes from scratch
Same as before: title/bullets/chart/image/outro scenes + a **topic-specific
`"theme"`** (base/glow1/glow2/text/sub/accents colors chosen for the subject —
never default Momo pink/blue for unrelated topics). Numbers in the narration →
`chart` scene (animated growth).

## Lip-synced human presenter (REAL Wav2Lip — installed and working)
3-step flow:
```
1) python python/momo_video.py --storyboard b.json --out main.mp4 --export-audio narr.wav
2) python python/momo_lipsync.py --face data/videos/presenter.jpg --audio narr.wav --out head.mp4
   (CPU: ~4-5 min per minute of audio; prints progress N/total)
3) ffmpeg -i main.mp4 -i head.mp4 -filter_complex
     "[1:v]scale=250:-2[ph];[0:v][ph]overlay=W-w-42:H-h-42:eof_action=pass[v]"
     -map "[v]" -map 0:a -c:v libx264 -pix_fmt yuv420p -c:a copy final.mp4
```
(Or set `"presenter_video": "head.mp4"` in the storyboard and rerun momo_video —
it composites automatically.)
- **Portrait**: data/videos/presenter.jpg (the user's chosen face). Any clear
  front-facing photo works; the user can replace it. Weights live in
  %LOCALAPPDATA%/Momo/wav2lip. The proxy blocks AI-face websites — new portraits
  must come from the user.
- Alternative non-human presenter: static `"presenter": {"image": ..., "name":...}`
  (circular PiP with pulse ring, no lip-sync).

## Voices (edge-tts, natural neural)
- "en-US-JennyNeural" / "en-US-AriaNeural" — warm American women (default)
- "en-US-GuyNeural" man; "en-IN-NeerjaNeural" Indian English woman
- Spanish/Latin-American: "es-MX-DaliaNeural", "es-US-PalomaNeural"

## QA (mandatory before telling the user it's done)
Extract 3-4 frames (`ffmpeg -ss N -i out.mp4 -frames:v 1 f.png`), LOOK at them
(popup placement, text legibility, presenter not covering content), then
Start-Process the mp4.

## Extending
Renderers in momo_video.py are plain PIL/matplotlib — add scene types (comparison,
process flow, quote) by copying one. Keep motion language: ease() reveals,
staggered timing, crossfades.
