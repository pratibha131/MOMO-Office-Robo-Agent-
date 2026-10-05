# Instagram Reels (vertical 9:16 promo videos) — python/momo_reel.py

Trigger phrases: "make an Instagram reel about yourself", "make a reel promoting Momo",
"make a 30-second vertical video about X", "reel for Instagram / Shorts / TikTok".
Teaching videos (16:9, narrated slides) are a different tool — knowledge/teaching-videos.md.

## One command
```powershell
$env:PYTHONUTF8 = "1"
python python\momo_reel.py --storyboard data\videos\momo_reel_board.json --out data\videos\Momo_Reel.mp4
```
Renders 1080x1920 at 30 fps: animated glow background, Momo's character art floating in,
big pop-in captions synced to the narration, feature cards sliding in, a call-to-action
close, crossfade cuts, neural narration (Neerja = Momo's voice). Takes about 4-6 minutes
for a 45-50 s reel — tell the user "give me five minutes" and continue in the background.
Optional `--music path.mp3` mixes a LOCAL track under the voice (the proxy blocks music
sites; the user must supply the file). The post caption + hashtags from the storyboard are
saved next to the mp4 as `<name>_caption.txt`.

## Ready-made promo: the Momo reel
`data/videos/momo_reel_board.json` is the approved self-promo storyline. For "make a reel
about yourself / promoting Momo" just render it (edit the wording only if the user asks).
Storyline (35-45 s): hook (0-4 s) -> relatable problem (4-9 s) -> 4 features (9-22 s) ->
proof moment, the Excel-to-PowerPoint story (22-31 s) -> privacy reassurance (31-37 s) ->
call-to-action (37-44 s).

## Writing a NEW reel storyboard (any topic)
Scene types: `hook` (heading + sub, character rises), `caption` (label pill + spoken-word
captions, character on a side: alternate left/right), `feature` (heading + 2-4 short
cards, character small), `cta` (heading, sub, handle pill, character large with pulse).
Rules that make reels work:
1. First 3 seconds decide everything: the hook heading is 2-4 words, the narration is one
   punchy sentence. Never start with "Hi, in this video...".
2. 35-45 seconds total, 5-7 scenes, one idea per scene, narration 1-2 sentences each,
   spoken naturally (contractions, no jargon). Captions are cut automatically from the
   narration, so write narration you'd be happy to read on screen.
3. Feature cards: max 4, each under 8 words, start with a verb.
4. End with a clear ask (follow / try / comment) and a warm last line.
5. Keep facts true to what Momo really does (see README) — no invented capabilities.
6. Theme: keep Momo's midnight/pink/blue palette for Momo promos; pick topic-specific
   colours for anything else (`theme` in the storyboard).
Save new storyboards under data/videos/<name>_board.json.

## QA (mandatory before saying it's done)
Extract 4 frames and LOOK at them:
```powershell
$ff = python -c "import imageio_ffmpeg;print(imageio_ffmpeg.get_ffmpeg_exe())"
foreach ($s in 1,8,20,40) { & $ff -y -ss $s -i data\videos\Momo_Reel.mp4 -frames:v 1 data\videos\qa_reel_$s.png }
```
Check: text inside the safe zone (not under Instagram's top bar or bottom buttons),
captions legible, character not covering text, no cut-off words. Then
`Start-Process data\videos\Momo_Reel.mp4` and tell the user the length and where the
caption text is. Suggest posting from the phone: AirDrop/OneDrive the mp4, paste the
caption file.
