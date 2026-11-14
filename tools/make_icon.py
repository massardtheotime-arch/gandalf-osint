"""Build the app icon: Gollum cut out of gollum.png, set in a rounded tile with a gold border.

gollum.png carries a fake transparency checkerboard (opaque 20px white/240-grey squares),
so the background is removed by matching that exact pattern before compositing.

Usage:
    python3 tools/make_icon.py preview OUT.png        # both framings + small sizes, for review
    python3 tools/make_icon.py write [full|closeup]   # icon.icns, icon.iconset/*, (Windows reads the 1024 png)
"""
import os
import subprocess
import sys
import tempfile

import numpy as np
from PIL import Image, ImageChops, ImageDraw, ImageFilter

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
SRC = os.path.join(ROOT, "gollum.png")

CANVAS = 1024
BODY = 824                      # macOS icon grid: 824px tile centred in a 1024px canvas
MARGIN = (CANVAS - BODY) // 2
RADIUS = 185
BORDER = 22
GOLD_TOP, GOLD_BOTTOM = (246, 220, 151), (200, 150, 74)
SS = 2                          # supersampling for smooth edges


def cutout():
    src = Image.open(SRC).convert("RGB")
    im = np.array(src).astype(int)
    h, w, _ = im.shape
    yy, xx = np.mgrid[0:h, 0:w]
    checker = np.where(((xx // 20) + (yy // 20)) % 2 == 0, 255, 240)[..., None]
    is_bg = (np.abs(im - checker).max(axis=2) <= 4)
    alpha = Image.fromarray(np.where(is_bg, 0, 255).astype("uint8"))
    # drop 1px so blended edge pixels (part checker, part Gollum) go away, then soften
    alpha = alpha.filter(ImageFilter.MinFilter(3)).filter(ImageFilter.GaussianBlur(0.7))

    # defringe: repaint the edge band with colours bled outward from the solid interior,
    # so no light checkerboard tint survives on the dark tile
    core = np.array(alpha.filter(ImageFilter.MinFilter(5))).astype(float) / 255
    rgb = np.array(src).astype(float)
    def blur(a):  # a: float array in 0..255, HxWxC
        chans = [np.array(Image.fromarray(a[..., i].clip(0, 255).astype("uint8"), "L")
                          .filter(ImageFilter.GaussianBlur(3))).astype(float) for i in range(a.shape[2])]
        return np.stack(chans, axis=2)
    num = blur(rgb * core[..., None])
    den = blur(core[..., None] * 255) / 255
    est = num / np.maximum(den, 1e-3)
    fringe = ((core < 0.5) & (den[..., 0] > 0.04))[..., None]
    fixed = np.where(fringe, est, rgb).clip(0, 255).astype("uint8")
    out = Image.fromarray(fixed, "RGB").convert("RGBA")
    out.putalpha(alpha)
    return out


def rounded_mask(size, inset, radius):
    m = Image.new("L", (size, size), 0)
    ImageDraw.Draw(m).rounded_rectangle([inset, inset, size - 1 - inset, size - 1 - inset], radius=radius, fill=255)
    return m


def vertical_gradient(size, top, bottom):
    t = np.linspace(0, 1, size)[:, None, None]
    arr = (np.array(top) * (1 - t) + np.array(bottom) * t).astype("uint8")
    arr = np.repeat(arr, size, axis=1)
    return Image.fromarray(arr, "RGB")


def build(variant="full"):
    s = CANVAS * SS
    m, body, rad, bord = MARGIN * SS, BODY * SS, RADIUS * SS, BORDER * SS

    # tile background: warm dark with a soft glow behind the subject
    yy, xx = np.mgrid[0:s, 0:s]
    d = np.sqrt(((xx - s / 2) / (s * 0.5)) ** 2 + ((yy - s * 0.42) / (s * 0.5)) ** 2)
    t = np.clip(d, 0, 1)[..., None]
    bg = (np.array([54, 42, 28]) * (1 - t) + np.array([14, 12, 10]) * t).astype("uint8")
    tile = Image.fromarray(bg, "RGB").convert("RGBA")

    # subject
    g = cutout()
    inner = body - 2 * bord
    if variant == "closeup":
        box = (215, 15, 685, 485)          # head and shoulders
        g = g.crop(box)
        scale = inner / g.width * 1.02
        g = g.resize((round(g.width * scale), round(g.height * scale)), Image.LANCZOS)
        gx, gy = (s - g.width) // 2, m + bord + 4 * SS
    else:
        bb = g.getchannel("A").getbbox()
        g = g.crop(bb)
        scale = inner * 0.9 / g.height
        g = g.resize((round(g.width * scale), round(g.height * scale)), Image.LANCZOS)
        gx, gy = (s - g.width) // 2, m + bord + (inner - g.height) // 2 + 6 * SS
    layer = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    layer.alpha_composite(g, (gx, gy))
    tile = Image.alpha_composite(tile, layer)

    # clip to the rounded tile
    outer = rounded_mask(s, m, rad)
    inner_mask = rounded_mask(s, m + bord, rad - bord)
    tile.putalpha(ImageChops.multiply(tile.getchannel("A"), inner_mask))

    # inner shadow along the border so the photo sits "inside" the frame
    ring_in = ImageChops.subtract(inner_mask, inner_mask.filter(ImageFilter.GaussianBlur(14 * SS)).point(lambda v: v))
    shade = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    shade.putalpha(ring_in.point(lambda v: int(v * 0.55)))
    tile = Image.alpha_composite(tile, shade)
    tile.putalpha(ImageChops.multiply(tile.getchannel("A"), inner_mask))

    # gold border
    ring = ImageChops.subtract(outer, inner_mask)
    gold = vertical_gradient(s, GOLD_TOP, GOLD_BOTTOM).convert("RGBA")
    gold.putalpha(ring)
    tile = Image.alpha_composite(tile, gold)

    # drop shadow + final canvas
    canvas = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    sh = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    sh.putalpha(outer.point(lambda v: int(v * 0.5)))
    sh = ImageChops.offset(sh, 0, 10 * SS).filter(ImageFilter.GaussianBlur(14 * SS))
    canvas = Image.alpha_composite(canvas, sh)
    canvas = Image.alpha_composite(canvas, tile)
    return canvas.resize((CANVAS, CANVAS), Image.LANCZOS)


def preview(out_path):
    variants = [("full", build("full")), ("closeup", build("closeup"))]
    sizes = [128, 64, 32, 16]
    pad = 24
    row_h = 400
    W = 2 * 400 + 3 * pad + 330
    sheet = Image.new("RGB", (W, 2 * row_h + 3 * pad), (30, 30, 34))
    light = Image.new("RGB", (330, 2 * row_h + 3 * pad), (236, 236, 238))
    sheet.paste(light, (W - 330, 0))
    for r, (name, icon) in enumerate(variants):
        y0 = pad + r * (row_h + pad)
        big = icon.resize((400, 400), Image.LANCZOS)
        sheet.paste(big, (pad, y0), big)
        x = 2 * pad + 400
        for bgx in (x, W - 330 + pad):
            xx = bgx
            for sz in sizes:
                small = icon.resize((sz, sz), Image.LANCZOS)
                sheet.paste(small, (xx, y0 + 40), small)
                xx += sz + 12
        sheet_draw = ImageDraw.Draw(sheet)
        sheet_draw.text((pad, y0 - 14), {"full": "A  silhouette entiere", "closeup": "B  gros plan"}[name], fill=(240, 220, 160))
    sheet.save(out_path)


def write(variant):
    icon = build(variant)
    iconset_dir = os.path.join(ROOT, "icon.iconset")
    for sz in (16, 32, 64, 128, 256, 512, 1024):
        icon.resize((sz, sz), Image.LANCZOS).save(os.path.join(iconset_dir, f"icon_{sz}x{sz}.png"))
    names = {16: ["icon_16x16.png"], 32: ["icon_16x16@2x.png", "icon_32x32.png"], 64: ["icon_32x32@2x.png"],
             128: ["icon_128x128.png"], 256: ["icon_128x128@2x.png", "icon_256x256.png"],
             512: ["icon_256x256@2x.png", "icon_512x512.png"], 1024: ["icon_512x512@2x.png"]}
    with tempfile.TemporaryDirectory(suffix=".iconset") as tmp:
        for sz, files in names.items():
            img = icon.resize((sz, sz), Image.LANCZOS)
            for f in files:
                img.save(os.path.join(tmp, f))
        subprocess.run(["iconutil", "-c", "icns", tmp, "-o", os.path.join(ROOT, "icon.icns")], check=True)
    print("wrote icon.icns and icon.iconset/*")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "preview" and len(sys.argv) > 2:
        preview(sys.argv[2])
    elif cmd == "write":
        write(sys.argv[2] if len(sys.argv) > 2 else "full")
    else:
        print(__doc__)
