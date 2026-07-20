"""Play store feature graphic: exactly 1024x500, 24-bit, no alpha."""
import os
from PIL import Image, ImageDraw, ImageFont, ImageFilter

REPO = os.path.join(os.path.dirname(__file__), "convertly-web")
OUT = os.path.join(REPO, "static", "icons", "play-feature-graphic-1024x500.png")

W, H = 1024, 500
BG = (13, 13, 16)
LATIN_BOLD = "C:/Windows/Fonts/segoeuib.ttf"
LATIN = "C:/Windows/Fonts/segoeui.ttf"

img = Image.new("RGB", (W, H), BG)

# Brand glow, blurred twice so it does not band after upscaling.
glow = Image.new("L", (W // 4, H // 4), 0)
gd = ImageDraw.Draw(glow)
gd.ellipse([W // 8 - 90, H // 8 - 90, W // 8 + 90, H // 8 + 90], fill=95)
glow = glow.filter(ImageFilter.GaussianBlur(45)).resize((W, H), Image.LANCZOS)
glow = glow.filter(ImageFilter.GaussianBlur(18))
img = Image.composite(Image.new("RGB", (W, H), (30, 72, 46)), img, glow)

icon = Image.open(os.path.join(REPO, "static", "icons", "icon-512.png")).convert("RGBA")
icon = icon.resize((150, 150), Image.LANCZOS)

d = ImageDraw.Draw(img)
f_title = ImageFont.truetype(LATIN_BOLD, 96)
f_sub = ImageFont.truetype(LATIN, 38)

title = "Convertly"
sub = "PDF, Word & Excel - converted in seconds"

tw = d.textbbox((0, 0), title, font=f_title)[2]
block_w = icon.width + 32 + tw
x0 = (W - block_w) // 2
y0 = 150

img.paste(icon, (x0, y0 - 26), icon)
d.text((x0 + icon.width + 32, y0), title, font=f_title, fill=(238, 238, 245))

sw = d.textbbox((0, 0), sub, font=f_sub)[2]
d.text(((W - sw) // 2, y0 + 168), sub, font=f_sub, fill=(150, 200, 170))

img.save(OUT, "PNG", optimize=True)
im = Image.open(OUT)
print(f"{os.path.basename(OUT)}: {im.width}x{im.height} mode={im.mode} "
      f"{os.path.getsize(OUT)/1024:.0f} KB")
assert (im.width, im.height) == (1024, 500) and im.mode == "RGB"
print("meets Play spec: 1024x500, 24-bit, no alpha")
