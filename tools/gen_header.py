"""
Generate the Google Play developer-page header image.

Play spec: 4096 x 2304, 24-bit JPEG/PNG, non-transparent, <= 1 MB.

The header is cropped aggressively at different viewport sizes, so all content
sits inside a centred safe area rather than near the edges.

Arabic is reshaped + bidi-reordered the same way app.py does it (_prepare_rtl_text),
using the Amiri font already vendored at static/fonts/. Rendering Arabic without
that step produces disconnected, backwards glyphs.
"""
import os
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import arabic_reshaper
from bidi.algorithm import get_display

REPO = os.path.join(os.path.dirname(__file__), "convertly-web")
OUT = os.path.join(REPO, "static", "icons", "play-developer-header-4096x2304.jpg")

W, H = 4096, 2304
BG = (13, 13, 16)
ACCENT = (74, 222, 128)

AMIRI = os.path.join(REPO, "static", "fonts", "Amiri-Regular.ttf")
LATIN_BOLD = "C:/Windows/Fonts/segoeuib.ttf"


def rtl(text):
    return get_display(arabic_reshaper.reshape(text))


img = Image.new("RGB", (W, H), BG)

# Soft radial brand glow behind the mark, built at low res then upscaled so the
# blur stays cheap at 4K.
glow = Image.new("L", (W // 4, H // 4), 0)
gd = ImageDraw.Draw(glow)
cx, cy = W // 8, H // 8
gd.ellipse([cx - 300, cy - 300, cx + 300, cy + 300], fill=90)
# Heavy blur at quarter-res, then a second pass after upscaling: a single
# low-res blur leaves visible banding once stretched to 4K.
glow = glow.filter(ImageFilter.GaussianBlur(160)).resize((W, H), Image.LANCZOS)
glow = glow.filter(ImageFilter.GaussianBlur(40))
img = Image.composite(Image.new("RGB", (W, H), (30, 70, 45)), img, glow)

# Brand mark, reusing the same source SVG as the app icons (already rasterised).
icon_png = os.path.join(REPO, "static", "icons", "icon-512.png")
icon = Image.open(icon_png).convert("RGBA").resize((460, 460), Image.LANCZOS)

f_title = ImageFont.truetype(LATIN_BOLD, 300)
f_sub = ImageFont.truetype(AMIRI, 150)

d = ImageDraw.Draw(img)
title = "Convertly"
sub = rtl("حوِّل ملفاتك في ثوانٍ")

tw = d.textbbox((0, 0), title, font=f_title)[2]
sw = d.textbbox((0, 0), sub, font=f_sub)[2]

block_w = icon.width + 60 + tw
start_x = (W - block_w) // 2
row_y = H // 2 - 210

img.paste(icon, (start_x, row_y - 60), icon)
d.text((start_x + icon.width + 60, row_y), title, font=f_title, fill=(238, 238, 245))
d.text(((W - sw) // 2, row_y + 420), sub, font=f_sub, fill=(150, 200, 170))

# Accent rule under the lockup
d.rounded_rectangle([W // 2 - 190, row_y + 800, W // 2 + 190, row_y + 812],
                    radius=6, fill=ACCENT)

q = 92
while q >= 60:
    img.save(OUT, "JPEG", quality=q, optimize=True, subsampling=0)
    size = os.path.getsize(OUT)
    if size <= 1024 * 1024:
        break
    q -= 6

im = Image.open(OUT)
print(f"wrote {os.path.basename(OUT)}")
print(f"  {im.width}x{im.height}  mode={im.mode}  quality={q}  {size/1024:.0f} KB")
print(f"  within 1MB: {size <= 1024*1024}")
