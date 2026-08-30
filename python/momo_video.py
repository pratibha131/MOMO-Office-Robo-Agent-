"""Momo's video studio: storyboard JSON -> narrated teaching video (MP4).

Scenes are rendered as brand-styled animated frames (PIL + matplotlib), narrated
with a natural American female neural voice (edge-tts), and assembled with the
bundled ffmpeg. Optional presenter portrait appears picture-in-picture.

Usage:
  python momo_video.py --storyboard board.json --out lesson.mp4

Storyboard format:
{
  "voice": "en-US-JennyNeural",              // natural US female (AriaNeural also good)
  "presenter": {"image": "path.jpg", "name": "Ava"},   // optional
  "scenes": [
    {"type": "title",   "heading": "...", "sub": "...",        "narration": "..."},
    {"type": "bullets", "heading": "...", "bullets": ["..."],  "narration": "..."},
    {"type": "chart",   "heading": "...", "narration": "...",
       "chart": {"kind": "bar"|"line", "labels": [...], "values": [...], "unit": "min"}},
    {"type": "image",   "heading": "...", "image": "path",     "narration": "..."},
    {"type": "outro",   "heading": "...", "sub": "...", "image": "path", "narration": "..."}
  ]
}
"""
import argparse
import asyncio
import json
import math
import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFilter, ImageFont

W, H = 1280, 720
FPS = 24

# Default theme (Momo brand). A storyboard can override ALL of it with
#   "theme": {"base": [r,g,b], "glow1": [...], "glow2": [...],
#             "text": [...], "sub": [...], "accents": [[...], ...]}
# so every video carries its TOPIC's identity, not a fixed template.
MIDNIGHT = (23, 20, 48)
PINK = (232, 98, 140)
BLUE = (59, 130, 217)
GOLD = (242, 184, 75)
WHITE = (247, 245, 250)
MUTED = (183, 176, 216)
ACCENTS = [PINK, BLUE, GOLD, (156, 107, 217), (58, 164, 107)]

def apply_theme(theme):
    global MIDNIGHT, PINK, BLUE, GOLD, WHITE, MUTED, ACCENTS
    if not theme:
        return
    def t(key, cur):
        v = theme.get(key)
        return tuple(v) if v else cur
    MIDNIGHT = t("base", MIDNIGHT)
    PINK = t("glow1", PINK)
    BLUE = t("glow2", BLUE)
    WHITE = t("text", WHITE)
    MUTED = t("sub", MUTED)
    if theme.get("accents"):
        ACCENTS = [tuple(a) for a in theme["accents"]]
        PINK = ACCENTS[0]
        if len(ACCENTS) > 1:
            BLUE = ACCENTS[1]
        if len(ACCENTS) > 2:
            GOLD = ACCENTS[2]

FONT_DIR = r"C:\Windows\Fonts"

def font(size, bold=False):
    name = "segoeuib.ttf" if bold else "segoeui.ttf"
    return ImageFont.truetype(os.path.join(FONT_DIR, name), size)

def ease(t):
    return 0.5 - 0.5 * math.cos(min(max(t, 0.0), 1.0) * math.pi)

def base_frame():
    im = Image.new("RGBA", (W, H), MIDNIGHT + (255,))
    for cx, cy, r, col, a in ((1150, 80, 330, PINK, 80), (100, 660, 360, BLUE, 70)):
        g = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(g)
        d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=col + (a,))
        im.alpha_composite(g.filter(ImageFilter.GaussianBlur(r * 0.5)))
    return im

def draw_presenter(im, presenter, t):
    if not presenter or not presenter.get("_img"):
        return
    d = ImageDraw.Draw(im)
    size = 150
    cx, cy = W - 115, H - 115
    pulse = 6 + 4 * math.sin(t * 4.0)
    d.ellipse([cx - size // 2 - pulse, cy - size // 2 - pulse,
               cx + size // 2 + pulse, cy + size // 2 + pulse],
              outline=PINK + (200,), width=4)
    im.alpha_composite(presenter["_img"], (cx - size // 2, cy - size // 2))
    name = presenter.get("name", "")
    if name:
        f = font(17, bold=True)
        tw = d.textlength(name, font=f)
        d.rounded_rectangle([cx - tw / 2 - 10, cy + size // 2 + 8,
                             cx + tw / 2 + 10, cy + size // 2 + 34], 13, fill=(0, 0, 0, 140))
        d.text((cx - tw / 2, cy + size // 2 + 11), name, font=f, fill=WHITE)

def wrap(d, txt, f, maxw):
    words = txt.split()
    lines, cur = [], ""
    for w_ in words:
        t = (cur + " " + w_).strip()
        if d.textlength(t, font=f) <= maxw:
            cur = t
        else:
            lines.append(cur)
            cur = w_
    if cur:
        lines.append(cur)
    return lines

def render_title(sc, t, dur, presenter):
    im = base_frame()
    d = ImageDraw.Draw(im)
    a1 = ease(t / 0.9)
    y = 250 + (1 - a1) * 40
    f = font(64, bold=True)
    for i, line in enumerate(wrap(d, sc["heading"], f, 1000)):
        d.text((90, y + i * 78), line, font=f, fill=WHITE + (int(255 * a1),))
    if sc.get("sub"):
        a2 = ease((t - 0.7) / 0.9)
        d.text((92, y + 170), sc["sub"], font=font(30), fill=(*PINK, int(255 * a2)))
    draw_presenter(im, presenter, t)
    return im

def render_bullets(sc, t, dur, presenter):
    im = base_frame()
    d = ImageDraw.Draw(im)
    d.text((90, 80), sc["heading"], font=font(44, bold=True), fill=WHITE)
    bullets = sc["bullets"]
    reveal_span = max(dur - 2.0, 1.0)
    per = reveal_span / max(len(bullets), 1)
    for i, b in enumerate(bullets):
        a = ease((t - 0.8 - i * per) / 0.6)
        if a <= 0:
            continue
        y = 200 + i * 92
        x = 110 + (1 - a) * 40
        col = ACCENTS[i % len(ACCENTS)]
        d.ellipse([x - 20, y + 6, x + 8, y + 34], fill=(*col, int(255 * a)))
        f = font(28)
        for j, line in enumerate(wrap(d, b, f, 950)):
            d.text((x + 30, y + j * 36), line, font=f, fill=(*WHITE, int(255 * a)))
    draw_presenter(im, presenter, t)
    return im

def render_chart(sc, t, dur, presenter):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    ch = sc["chart"]
    grow = ease(t / (dur * 0.55))
    fig, ax = plt.subplots(figsize=(8.2, 4.6), dpi=100)
    fig.patch.set_alpha(0.0)
    ax.set_facecolor((0, 0, 0, 0))
    labels, values = ch["labels"], [v * grow for v in ch["values"]]
    cols = ["#%02x%02x%02x" % a for a in ACCENTS]
    if ch.get("kind") == "line":
        n = max(2, int(len(values) * grow) or 2)
        ax.plot(labels[:n], ch["values"][:n], color=cols[0], linewidth=4, marker="o", markersize=9)
    else:
        bars = ax.bar(labels, values, color=[cols[i % len(cols)] for i in range(len(labels))])
        for b, v0 in zip(bars, ch["values"]):
            if grow > 0.97:
                ax.text(b.get_x() + b.get_width() / 2, b.get_height() + max(ch["values"]) * 0.03,
                        f"{v0:g}", ha="center", color="white", fontsize=13, fontweight="bold")
    ax.tick_params(colors="white", labelsize=12)
    for sp in ax.spines.values():
        sp.set_color("#5A527E")
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.set_ylim(0, max(ch["values"]) * 1.18)
    if ch.get("unit"):
        ax.set_ylabel(ch["unit"], color="white", fontsize=12)
    fig.tight_layout()
    tmp = os.path.join(tempfile.gettempdir(), "momo_chart.png")
    fig.savefig(tmp, transparent=True)
    plt.close(fig)

    im = base_frame()
    d = ImageDraw.Draw(im)
    d.text((90, 70), sc["heading"], font=font(42, bold=True), fill=WHITE)
    chart_im = Image.open(tmp).convert("RGBA")
    im.alpha_composite(chart_im, (110, 170))
    draw_presenter(im, presenter, t)
    return im

def render_image(sc, t, dur, presenter):
    im = base_frame()
    d = ImageDraw.Draw(im)
    d.text((90, 70), sc["heading"], font=font(42, bold=True), fill=WHITE)
    pic = Image.open(sc["image"]).convert("RGBA")
    a = ease(t / 0.9)
    hgt = 470
    wdt = int(hgt * pic.width / pic.height)
    pic = pic.resize((wdt, hgt))
    x = int(W / 2 - wdt / 2 + (1 - a) * 60)
    faded = pic.copy()
    faded.putalpha(faded.getchannel("A").point(lambda v: int(v * a)))
    im.alpha_composite(faded, (x, 180))
    draw_presenter(im, presenter, t)
    return im

def render_outro(sc, t, dur, presenter):
    im = base_frame()
    d = ImageDraw.Draw(im)
    a = ease(t / 0.9)
    f = font(56, bold=True)
    lines = wrap(d, sc["heading"], f, 700)
    for i, line in enumerate(lines):
        d.text((90, 250 + i * 70), line, font=f, fill=(*WHITE, int(255 * a)))
    if sc.get("sub"):
        d.text((92, 250 + len(lines) * 70 + 25), sc["sub"], font=font(28),
               fill=(*PINK, int(255 * ease((t - 0.6) / 0.9))))
    if sc.get("image") and os.path.exists(sc["image"]):
        pic = Image.open(sc["image"]).convert("RGBA")
        hgt = 560
        wdt = int(hgt * pic.width / pic.height)
        pic = pic.resize((wdt, hgt))
        im.alpha_composite(pic, (W - wdt - 80, H - hgt - 40))
    draw_presenter(im, presenter, t)
    return im


def render_slide(sc, t, dur, presenter):
    """EXACT deck slide as the scene: Ken Burns zoom + optional popup demo window."""
    slide = Image.open(sc["image"]).convert("RGB")
    z0, z1 = (1.0, 1.07) if sc.get("zoom") != "out" else (1.07, 1.0)
    z = z0 + (z1 - z0) * (t / max(dur, 0.1))
    zw, zh = int(W * z), int(H * z)
    slide = slide.resize((zw, zh))
    im = slide.crop(((zw - W) // 2, (zh - H) // 2,
                     (zw - W) // 2 + W, (zh - H) // 2 + H)).convert("RGBA")

    pop = sc.get("popup")
    if pop:
        a = ease((t - pop.get("at", 2.0)) / 0.6)
        if a > 0:
            lines = pop.get("lines", [])
            pw, ph = pop.get("w", 560), 60 + 44 * len(lines)
            px = pop.get("x", W - pw - 70)
            py = int(pop.get("y", H - ph - 90) + (1 - a) * 50)
            win = Image.new("RGBA", (pw, ph), (0, 0, 0, 0))
            wd = ImageDraw.Draw(win)
            wd.rounded_rectangle([0, 0, pw - 1, ph - 1], 16, fill=(24, 22, 34, int(242 * a)),
                                 outline=(255, 255, 255, int(70 * a)), width=1)
            wd.rounded_rectangle([0, 0, pw - 1, 38], 16, fill=(42, 39, 58, int(250 * a)))
            wd.rectangle([0, 20, pw - 1, 38], fill=(42, 39, 58, int(250 * a)))
            for ci, cc in enumerate([(255, 95, 86), (255, 189, 46), (39, 201, 63)]):
                wd.ellipse([14 + ci * 22, 13, 26 + ci * 22, 25], fill=cc + (int(255 * a),))
            wd.text((84, 9), pop.get("title", "demo"), font=font(16, bold=True),
                    fill=(255, 255, 255, int(235 * a)))
            t_pop = t - pop.get("at", 2.0) - 0.4
            per = max((dur - pop.get("at", 2.0) - 1.2) / max(len(lines), 1), 0.8)
            for li, line in enumerate(lines):
                la = ease((t_pop - li * per) / 0.4)
                if la <= 0:
                    continue
                frac = min(max((t_pop - li * per) / (per * 0.7), 0.0), 1.0)
                shown = line[:max(1, int(len(line) * frac))]
                col = (170, 255, 200) if line.strip().startswith(">") else (238, 235, 248)
                wd.text((22, 50 + li * 44), shown, font=font(18), fill=col + (int(255 * la),))
            im.alpha_composite(win, (px, py))
    draw_presenter(im, presenter, t)
    return im

RENDERERS = {"title": render_title, "bullets": render_bullets, "chart": render_chart,
             "image": render_image, "outro": render_outro, "slide": render_slide}

def synth(text, voice, out):
    import edge_tts

    async def run():
        await edge_tts.Communicate(text, voice=voice, rate="-4%").save(out)

    asyncio.run(asyncio.wait_for(run(), timeout=60))

def audio_len(path):
    from mutagen.mp3 import MP3
    return MP3(path).info.length

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--storyboard", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--export-audio", default=None,
                    help="also write the full narration as one 16k mono wav (for lip-sync)")
    args = ap.parse_args()

    import imageio_ffmpeg
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()

    board = json.load(open(args.storyboard, encoding="utf-8"))
    apply_theme(board.get("theme"))
    voice = board.get("voice", "en-US-JennyNeural")
    presenter = board.get("presenter")
    if presenter and presenter.get("image") and os.path.exists(presenter["image"]):
        p = Image.open(presenter["image"]).convert("RGBA").resize((150, 150))
        mask = Image.new("L", (150, 150), 0)
        ImageDraw.Draw(mask).ellipse([0, 0, 150, 150], fill=255)
        p.putalpha(mask)
        presenter["_img"] = p
    elif presenter:
        presenter["_img"] = None

    work = tempfile.mkdtemp(prefix="momo_video_")
    scene_files = []
    mp3s = []
    prev_last = None            # last frame of previous scene -> crossfade cuts
    XFADE = int(0.45 * FPS)
    try:
        for si, sc in enumerate(board["scenes"]):
            mp3 = os.path.join(work, f"n{si}.mp3")
            synth(sc["narration"], voice, mp3)
            mp3s.append(mp3)
            dur = audio_len(mp3) + 0.7
            frames_dir = os.path.join(work, f"f{si}")
            os.makedirs(frames_dir)
            n_frames = int(dur * FPS)
            render = RENDERERS[sc["type"]]
            print(f"scene {si + 1}/{len(board['scenes'])} [{sc['type']}] {dur:.1f}s "
                  f"({n_frames} frames)...", flush=True)
            for fi in range(n_frames):
                im = render(sc, fi / FPS, dur, presenter)
                if prev_last is not None and fi < XFADE:
                    im = Image.blend(prev_last, im.convert("RGB").convert("RGBA"),
                                     ease(fi / XFADE))
                im.convert("RGB").save(os.path.join(frames_dir, f"{fi:05d}.jpg"), quality=88)
            prev_last = render(sc, max(n_frames - 1, 0) / FPS, dur,
                               presenter).convert("RGB").convert("RGBA")
            scene_mp4 = os.path.join(work, f"s{si}.mp4")
            subprocess.run([ffmpeg, "-y", "-framerate", str(FPS),
                            "-i", os.path.join(frames_dir, "%05d.jpg"),
                            "-i", mp3, "-c:v", "libx264", "-pix_fmt", "yuv420p",
                            "-c:a", "aac", "-shortest", scene_mp4],
                           check=True, capture_output=True)
            scene_files.append(scene_mp4)
            shutil.rmtree(frames_dir)

        concat = os.path.join(work, "concat.txt")
        with open(concat, "w") as f:
            for s in scene_files:
                f.write(f"file '{s}'\n")
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        head = board.get("presenter_video")
        main_mp4 = os.path.join(work, "main.mp4") if head else args.out
        subprocess.run([ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", concat,
                        "-c", "copy", main_mp4], check=True, capture_output=True)

        if args.export_audio:
            alist = os.path.join(work, "alist.txt")
            with open(alist, "w") as f:
                for m in mp3s:
                    f.write(f"file '{m}'\n")
            subprocess.run([ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", alist,
                            "-ar", "16000", "-ac", "1", args.export_audio],
                           check=True, capture_output=True)
            print("narration audio:", args.export_audio)

        if head:
            # lip-synced talking head (momo_lipsync.py output) bottom-right
            subprocess.run([ffmpeg, "-y", "-i", main_mp4, "-i", head,
                            "-filter_complex",
                            "[1:v]scale=250:-2[ph];[0:v][ph]overlay=W-w-42:H-h-42:eof_action=pass[v]",
                            "-map", "[v]", "-map", "0:a", "-c:v", "libx264",
                            "-pix_fmt", "yuv420p", "-c:a", "copy", args.out],
                           check=True, capture_output=True)
        print("VIDEO READY:", args.out)
    finally:
        shutil.rmtree(work, ignore_errors=True)

if __name__ == "__main__":
    main()
