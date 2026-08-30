# Building amazing presentations (Canva-quality, in PowerPoint)

The user asks for decks ("make me a presentation on X"). Deliver **rich, designed,
animated** decks — never plain bullets on a flat background. Canva itself has no
API access here; you build Canva quality natively. A finished deck example lives at
data/presentations/Momo_Marketing_Deck.pptx — study it before your first deck.

## Workflow (proven end-to-end on this machine)
1. **Design system first — UNIQUE PER TOPIC**: pick a palette, motif and layout
   language that belong to THIS subject (a cardiology deck, a finance deck and a
   Momo deck must not look like siblings). One dominant color, 1-2 support tones,
   one accent; Calibri/Segoe typography (36-44pt titles, 14-16pt body); ONE
   repeating motif (icon circles, chip pills, rounded cards — or something the
   topic itself suggests). 16:9 = 13.333 x 7.5 inches. The Momo deck is an example
   of the METHOD, not a template to clone.
2. **Backgrounds with PIL**: render 1920x1080 PNGs — dark base + 2-3 blurred radial
   color glows (GaussianBlur of ellipses) + faint dot grid = premium gradient feel.
   Make a dark variant (title/section/closing) and a light variant (content).
3. **Build with python-pptx** (installed): full-bleed background picture first on
   every slide, then content. Helpers worth recreating: text() with per-run styling,
   rounded-rect cards with soft shadows, icon circles with EMOJI text (Segoe UI
   Emoji renders in PowerPoint — instant colorful icons), chip pills.
4. **Images**: reuse assets from reference decks the user provides (unzip the pptx,
   ppt/media/), the character art in app/renderer/ (momo.png, toto.png,
   Philips-uniformed, transparent), and rounded-corner crops of photos (PIL mask).
   The corporate proxy blocks most stock-photo sites — prefer local/harvested art.
5. **Animations via PowerPoint COM** (the magic step):
   ```powershell
   $pp = New-Object -ComObject PowerPoint.Application
   $pres = $pp.Presentations.Open($path, $false, $false, $false)
   foreach ($slide in $pres.Slides) {
     $i = 0
     foreach ($shape in $slide.Shapes) {
       if ($shape.ZOrderPosition -eq 1) { continue }        # skip background
       $eff = $slide.TimeLine.MainSequence.AddEffect($shape, 10, 0, 2)  # Fade, WithPrevious
       $eff.Timing.Duration = 0.5
       $eff.Timing.TriggerDelayTime = 0.25 + ($i * 0.10)    # staggered cascade
       $i++
     } }
   $pres.SaveAs($out); $pres.Close(); $pp.Quit()
   ```
6. **Slide transitions via XML injection** (python, zipfile): remove existing
   `<p:transition>`, insert e.g. `<p:transition spd="med"><p:fade/></p:transition>`
   (or `<p:push dir="u"/>`, `<p:wipe dir="l"/>`) before `<p:timing>` in each
   ppt/slides/slideN.xml. Vary: fade for content, push for section changes.
7. **QA (mandatory)**: export every slide via COM
   `$pres.Slides.Item($i).Export($png, "PNG", 1280, 720)` and LOOK at them —
   check overflow past slide edges, overlaps, contrast. Fix and re-run.

## Hard-won gotchas
- `shape.shadow.inherit = False` already writes an empty `<a:effectLst/>` —
  REUSE it for custom shadow XML; appending a second one makes a file PowerPoint
  REFUSES TO OPEN (python-pptx will still open it — always COM-validate).
- PowerPoint COM cannot open files under Temp (Protected View) — build into the
  project folder.
- Strikethrough: `run.font._rPr.set('strike', 'sngStrike')` — safe.
- Character images: compute width from aspect ratio; check right edge < 13.33".
- Never: accent bars under titles, color stripes, text-only slides, same layout
  twice in a row, centered body text.

## Mimicking a reference deck the user likes
Unzip it, export its slides to PNG via COM, LOOK at them, extract ppt/media/
assets, and adopt its strongest ideas (half-bleed imagery, header brand row)
into your own layout system — don't clone it slide-for-slide.
