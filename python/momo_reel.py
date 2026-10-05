# -*- coding: utf-8 -*-
"""
momo_reel.py — vertical 9:16 Instagram Reel renderer (1080x1920, 30 fps) for Momo.

    python python/momo_reel.py --storyboard board.json --out data/videos/momo_reel.mp4 [--music track.mp3]

Storyboard (JSON):
{
  "voice": "en-IN-NeerjaNeural",
  "character": "app/renderer/momo.png",          # transparent PNG of the character (optional)
  "theme": {"base":[r,g,b], "glow1":[..], "glow2":[..], "text":[..], "sub":[..], "accent":[..]},
  "caption": "post text with #hashtags (saved next to the mp4)",
  "scenes": [
    {"type": "hook",    "heading": "Meet Momo.", "sub": "...", "narration": "..."},
    {"type": "caption", "label": "THE PROBLEM", "side": "left|right", "narration": "..."},
    {"type": "feature", "heading": "...", "cards": ["...", "..."], "narration": "..."},
    {"type": "cta",     "heading": "...", "sub": "...", "handle": "@...", "narration": "..."}
  ]
}
Narration is spoken with edge-tts; on-screen captions are cut from the narration in short
chunks and pop in as the words are spoken (reel style). Text stays inside Instagram's safe
zone (away from the top and bottom UI overlays).
"""
import argparse
import asyncio
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw, ImageFilter, ImageFont

W, H = 1080, 1920
FPS = 24
SAFE_TOP, SAFE_BOTTOM = 240, 1560
FONT_DIR = r"C:\Windows\Fonts"
HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT = os.path.dirname(HERE)

THEME = {
    "base": (23, 20, 48), "glow1": (232, 98, 140), "glow2": (59, 130, 217),
    "text": (247, 245, 250), "sub": (200, 194, 228), "accent": (242, 184, 75),
}


# ----------------------------------------------------------------------------- helpers
def apply_theme(theme):
    if theme:
        for k, v in theme.items():
            if k in THEME and v:
                THEME[k] = tuple(v)


def font(size, bold=False):
    return ImageFont.truetype(os.path.join(FONT_DIR, "segoeuib.ttf" if bold else "segoeui.ttf"), size)


def ease(t):
    t = min(max(t, 0.0), 1.0)
    return 0.5 - 0.5 * math.cos(t * math.pi)


def ease_out_back(t, s=1.4):
    t = min(max(t, 0.0), 1.0) - 1
    return 1 + t * t * ((s + 1) * t + s)


_GLOWS = {}
_GRID = None


def _glow(col, r, a):
    key = (col, r, a)
    if key not in _GLOWS:
        size = int(r * 2.6)
        g = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        ImageDraw.Draw(g).ellipse([size / 2 - r, size / 2 - r, size / 2 + r, size / 2 + r], fill=col + (a,))
        _GLOWS[key] = g.filter(ImageFilter.GaussianBlur(r * 0.55))
    return _GLOWS[key]


def base_frame(t):
    """Background: dark base, three slowly drifting colour glows (pre-blurred once), faint dot grid."""
    global _GRID
    base, g1, g2 = THEME["base"], THEME["glow1"], THEME["glow2"]
    im = Image.new("RGBA", (W, H), base + (255,))
    drift = math.sin(t * 0.6) * 60
    for cx, cy, r, col, a in ((880 + drift, 260, 520, g1, 95), (160 - drift, 1500, 560, g2, 85), (560, 1000 + drift * 0.5, 380, g1, 35)):
        g = _glow(col, r, a)
        im.alpha_composite(g, (int(cx - g.width / 2), int(cy - g.height / 2)))
    if _GRID is None:
        _GRID = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(_GRID)
        for y in range(80, H, 120):
            for x in range(60, W, 120):
                d.ellipse([x, y, x + 3, y + 3], fill=THEME["text"] + (18,))
    im.alpha_composite(_GRID)
    return im


def wrap(d, txt, f, maxw):
    lines, cur = [], ""
    for w_ in txt.split():
        cand = (cur + " " + w_).strip()
        if d.textlength(cand, font=f) <= maxw:
            cur = cand
        else:
            if cur:
                lines.append(cur)
            cur = w_
    if cur:
        lines.append(cur)
    return lines


def text_center(d, txt, y, f, fill, maxw=920, gap=1.15, shadow=True):
    lines = wrap(d, txt, f, maxw)
    lh = int(f.size * gap)
    for i, line in enumerate(lines):
        tw = d.textlength(line, font=f)
        x = (W - tw) / 2
        if shadow:
            d.text((x + 3, y + i * lh + 4), line, font=f, fill=(0, 0, 0, int(fill[3] * 0.55) if len(fill) > 3 else 140))
        d.text((x, y + i * lh), line, font=f, fill=fill)
    return y + len(lines) * lh


def load_character(path, height):
    if not path or not os.path.exists(path):
        return None
    im = Image.open(path).convert("RGBA")
    w = int(height * im.width / im.height)
    return im.resize((w, height), Image.LANCZOS)


def draw_character(im, char, x, y, t, bob=16, period=2.6, scale=1.0, alpha=255):
    if char is None:
        return
    c = char
    if scale != 1.0:
        c = c.resize((max(1, int(c.width * scale)), max(1, int(c.height * scale))), Image.LANCZOS)
    dy = int(bob * math.sin(2 * math.pi * t / period))
    # soft shadow under the character (blurred once per size, cached)
    key = ("shadow", c.width)
    if key not in _GLOWS:
        sw, shh = int(c.width * 0.64) + 80, 130
        sh = Image.new("RGBA", (sw, shh), (0, 0, 0, 0))
        ImageDraw.Draw(sh).ellipse([40, 40, sw - 40, shh - 40], fill=(0, 0, 0, 110))
        _GLOWS[key] = sh.filter(ImageFilter.GaussianBlur(18))
    sh = _GLOWS[key]
    im.alpha_composite(sh, (int(x + c.width * 0.18 - 40), int(y + c.height - 70 + dy)))
    if alpha < 255:
        c = c.copy()
        c.putalpha(c.getchannel("A").point(lambda a: a * alpha // 255))
    im.alpha_composite(c, (int(x), int(y + dy)))


def layer():
    """Transparent drawing layer: PIL's ImageDraw writes RGBA pixels without blending, so
    translucent boxes and fading text must be drawn on a layer and alpha-composited."""
    lay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    return lay, ImageDraw.Draw(lay)


def caption_chunks(narration, size=5):
    """Split narration into spoken-word caption chunks: break at punctuation, max `size`
    words, and never leave a dangling 1-2 word chunk."""
    words = re.sub(r"\s+", " ", narration).strip().split(" ")
    chunks, cur = [], []
    for w_ in words:
        cur.append(w_)
        if len(cur) >= size or w_.endswith((".", "!", "?", ",", ":", ";")):
            chunks.append(cur)
            cur = []
    if cur:
        chunks.append(cur)
    merged = []
    for c in chunks:
        if merged and len(c) <= 2 and len(merged[-1]) + len(c) <= size + 2:
            merged[-1] = merged[-1] + c
        else:
            merged.append(c)
    return [" ".join(c) for c in merged]


def draw_caption(im, d, narration, t, dur, y, f, pad_time=0.6):
    """Reel-style caption: chunk of the narration that is being spoken right now, popping in."""
    chunks = caption_chunks(narration)
    speak = max(dur - pad_time, 0.1)
    # time per chunk proportional to its length
    weights = [len(c) for c in chunks]
    total = sum(weights) or 1
    acc, idx, start = 0.0, len(chunks) - 1, 0.0
    for i, wgt in enumerate(weights):
        seg = speak * wgt / total
        if t < acc + seg:
            idx, start = i, acc
            break
        acc += seg
    txt = chunks[idx].strip()
    pop = ease_out_back((t - start) / 0.22)
    size = max(int(f.size * (0.85 + 0.15 * pop)), 8)
    fpop = font(size, bold=True)
    lines = wrap(d, txt, fpop, 880)
    lh = int(size * 1.12)
    box_h = len(lines) * lh + 44
    box_w = max(d.textlength(l, font=fpop) for l in lines) + 72
    x0, y0 = (W - box_w) / 2, y - 22
    d.rounded_rectangle([x0, y0, x0 + box_w, y0 + box_h], 34, fill=(0, 0, 0, 150))
    for i, line in enumerate(lines):
        tw = d.textlength(line, font=fpop)
        # highlight colour for the chunk being spoken
        d.text(((W - tw) / 2, y + i * lh), line, font=fpop, fill=THEME["accent"] + (255,))
    return y0 + box_h


def label_pill(d, text, y):
    f = font(34, bold=True)
    tw = d.textlength(text, font=f)
    x0 = (W - tw) / 2 - 28
    d.rounded_rectangle([x0, y, x0 + tw + 56, y + 62], 31, fill=THEME["glow1"] + (235,))
    d.text((x0 + 28, y + 11), text, font=f, fill=THEME["text"] + (255,))


# ----------------------------------------------------------------------------- scene renderers
def render_hook(sc, t, dur, ctx):
    im = base_frame(t)
    lay, d = layer()
    a = ease_out_back(t / 0.7)
    size = max(int(100 * (0.7 + 0.3 * min(a, 1.2))), 8)
    fh = font(size, bold=True)
    y = text_center(d, sc["heading"], 560, fh, THEME["text"] + (int(255 * ease(t / 0.4)),), maxw=940)
    if sc.get("sub"):
        text_center(d, sc["sub"], y + 30, font(46), THEME["sub"] + (int(255 * ease((t - 0.5) / 0.7)),), maxw=880)
    char = ctx["char_big"]
    if char is not None:
        rise = ease((t - 0.2) / 0.9)
        y_char = H - 90 - char.height + (1 - rise) * (char.height + 120)
        draw_character(im, char, (W - char.width) / 2, y_char, t)
    im.alpha_composite(lay)
    return im


def render_caption(sc, t, dur, ctx):
    im = base_frame(t)
    lay, d = layer()
    if sc.get("label"):
        label_pill(d, sc["label"].upper(), SAFE_TOP + 20)
    char = ctx["char_mid"]
    if char is not None:
        side = sc.get("side", "right")
        x = 30 if side == "left" else W - 30 - char.width
        slide = ease(t / 0.6)
        x_anim = x + ((-1 if side == "left" else 1) * (1 - slide) * (char.width + 60))
        draw_character(im, char, x_anim, 980, t)
    draw_caption(im, d, sc["narration"], t, dur, 560, font(74, bold=True))
    im.alpha_composite(lay)
    return im


def render_feature(sc, t, dur, ctx):
    im = base_frame(t)
    lay, d = layer()
    text_center(d, sc["heading"], SAFE_TOP + 30, font(78, bold=True), THEME["text"] + (int(255 * ease(t / 0.5)),), maxw=940)
    cards = sc.get("cards", [])[:4]
    y = 560
    fc = font(44)
    for i, card in enumerate(cards):
        a = ease((t - 0.35 - i * 0.32) / 0.55)
        if a <= 0:
            continue
        x_off = (1 - a) * 420
        x0, x1 = 80 + x_off, 1000 + x_off
        lines = wrap(d, card, fc, 700)
        h = max(150, 60 + len(lines) * 52)
        d.rounded_rectangle([x0, y, x1, y + h], 36, fill=THEME["text"] + (int(30 * a),), outline=THEME["text"] + (int(70 * a),), width=2)
        cx, cy = x0 + 80, y + h / 2
        d.ellipse([cx - 34, cy - 34, cx + 34, cy + 34], fill=THEME["accent"] + (int(255 * a),))
        fn = font(36, bold=True)
        d.text((cx - d.textlength(str(i + 1), font=fn) / 2, cy - 24), str(i + 1), font=fn, fill=THEME["base"] + (255,))
        for li, line in enumerate(lines):
            d.text((x0 + 150, y + (h - len(lines) * 52) / 2 + li * 52), line, font=fc, fill=THEME["text"] + (int(255 * a),))
        y += h + 26
    char = ctx["char_small"]
    if char is not None:
        draw_character(im, char, W - 60 - char.width, H - 60 - char.height, t, bob=10)
    im.alpha_composite(lay)
    return im


def render_cta(sc, t, dur, ctx):
    im = base_frame(t)
    lay, d = layer()
    y = text_center(d, sc["heading"], 380, font(86, bold=True), THEME["text"] + (int(255 * ease(t / 0.5)),), maxw=940)
    if sc.get("sub"):
        y = text_center(d, sc["sub"], y + 24, font(46), THEME["sub"] + (int(255 * ease((t - 0.4) / 0.6)),), maxw=880)
    if sc.get("handle"):
        f = font(40, bold=True)
        tw = d.textlength(sc["handle"], font=f)
        x0 = (W - tw) / 2 - 30
        a = ease((t - 0.8) / 0.6)
        d.rounded_rectangle([x0, y + 40, x0 + tw + 60, y + 112], 36, fill=THEME["accent"] + (int(240 * a),))
        d.text((x0 + 30, y + 52), sc["handle"], font=f, fill=THEME["base"] + (int(255 * a),))
    char = ctx["char_big"]
    if char is not None:
        pulse = 18 + 10 * math.sin(t * 3.2)
        cx, cy = W / 2, H - 90 - char.height / 2 + 40
        r = char.height * 0.52 + pulse
        ring, rd = layer()
        for wd, al in ((22, 40), (14, 90), (8, 170)):   # soft ring without a per-frame blur
            rd.ellipse([cx - r, cy - r, cx + r, cy + r], outline=THEME["glow1"] + (al,), width=wd)
        im.alpha_composite(ring)
        draw_character(im, char, (W - char.width) / 2, H - 90 - char.height, t)
    im.alpha_composite(lay)
    return im


RENDERERS = {"hook": render_hook, "caption": render_caption, "feature": render_feature, "cta": render_cta}


# ----------------------------------------------------------------------------- audio
def synth(text, voice, out):
    import edge_tts

    async def run():
        await edge_tts.Communicate(text, voice=voice, rate="+2%").save(out)

    asyncio.run(asyncio.wait_for(run(), timeout=60))


def audio_len(path):
    from mutagen.mp3 import MP3
    return MP3(path).info.length


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="Render a 9:16 Instagram reel from a storyboard.")
    ap.add_argument("--storyboard", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--music", default=None, help="optional local music file, mixed under the narration")
    args = ap.parse_args()

    import imageio_ffmpeg
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()

    board = json.load(open(args.storyboard, encoding="utf-8"))
    apply_theme(board.get("theme"))
    voice = board.get("voice", "en-IN-NeerjaNeural")
    char_path = board.get("character") or os.path.join(PROJECT, "app", "renderer", "momo.png")
    if not os.path.isabs(char_path):
        char_path = os.path.join(PROJECT, char_path)
    ctx = {"char_big": load_character(char_path, 760), "char_mid": load_character(char_path, 620), "char_small": load_character(char_path, 400)}
    music = args.music or board.get("music")
    if music and not os.path.isabs(music):
        music = os.path.join(PROJECT, music)

    work = tempfile.mkdtemp(prefix="momo_reel_")
    scene_files, prev_last = [], None
    XFADE = int(0.35 * FPS)
    try:
        for si, sc in enumerate(board["scenes"]):
            mp3 = os.path.join(work, f"n{si}.mp3")
            synth(sc["narration"], voice, mp3)
            dur = audio_len(mp3) + (0.9 if sc["type"] in ("hook", "cta") else 0.6)
            frames_dir = os.path.join(work, f"f{si}")
            os.makedirs(frames_dir)
            n = int(dur * FPS)
            render = RENDERERS[sc["type"]]
            print(f"scene {si + 1}/{len(board['scenes'])} [{sc['type']}] {dur:.1f}s ({n} frames)...", flush=True)
            for fi in range(n):
                im = render(sc, fi / FPS, dur, ctx)
                if prev_last is not None and fi < XFADE:
                    im = Image.blend(prev_last, im, ease(fi / XFADE))
                im.convert("RGB").save(os.path.join(frames_dir, f"{fi:05d}.jpg"), quality=90)
            prev_last = render(sc, max(n - 1, 0) / FPS, dur, ctx)
            scene_mp4 = os.path.join(work, f"s{si}.mp4")
            subprocess.run([ffmpeg, "-y", "-framerate", str(FPS), "-i", os.path.join(frames_dir, "%05d.jpg"), "-i", mp3,
                            "-c:v", "libx264", "-pix_fmt", "yuv420p", "-c:a", "aac", "-shortest", scene_mp4],
                           check=True, capture_output=True)
            scene_files.append(scene_mp4)
            shutil.rmtree(frames_dir)

        concat = os.path.join(work, "concat.txt")
        with open(concat, "w") as f:
            for s in scene_files:
                f.write(f"file '{s}'\n")
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        main_mp4 = os.path.join(work, "main.mp4") if (music and os.path.exists(music)) else args.out
        subprocess.run([ffmpeg, "-y", "-f", "concat", "-safe", "0", "-i", concat, "-c", "copy", main_mp4], check=True, capture_output=True)
        if music and os.path.exists(music):
            subprocess.run([ffmpeg, "-y", "-i", main_mp4, "-stream_loop", "-1", "-i", music, "-filter_complex",
                            "[1:a]volume=0.16,afade=t=in:d=1.5[m];[0:a][m]amix=inputs=2:duration=first:dropout_transition=2[a]",
                            "-map", "0:v", "-map", "[a]", "-c:v", "copy", "-c:a", "aac", "-shortest", args.out],
                           check=True, capture_output=True)
        if board.get("caption"):
            cap = os.path.splitext(args.out)[0] + "_caption.txt"
            open(cap, "w", encoding="utf-8").write(board["caption"].strip() + "\n")
            print("caption text:", cap)
        print("REEL READY:", args.out)
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == "__main__":
    main()
