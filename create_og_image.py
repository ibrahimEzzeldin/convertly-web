#!/usr/bin/env python3
"""Generate og-image.png for social media sharing."""
from PIL import Image, ImageDraw, ImageFont
import os

# Create a 1200x630 image for Open Graph (social media)
width, height = 1200, 630
bg_color = (79, 110, 247)  # Indigo blue (#4f6ef7)
text_color = (255, 255, 255)  # White

image = Image.new("RGB", (width, height), bg_color)
draw = ImageDraw.Draw(image)

# Try to load a nice font, fall back to default
try:
    title_font = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 80)
    subtitle_font = ImageFont.truetype("C:/Windows/Fonts/segoeui.ttf", 40)
except:
    try:
        title_font = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", 80)
        subtitle_font = ImageFont.truetype("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", 40)
    except:
        title_font = ImageFont.load_default()
        subtitle_font = ImageFont.load_default()

# Draw the text in the center
title = "Convertly"
subtitle = "Free PDF Converter — No Sign-up"

# Center coordinates
title_y = (height - 150) // 2
subtitle_y = title_y + 100

try:
    draw.text((width//2, title_y), title, fill=text_color, font=title_font, anchor="mm")
    draw.text((width//2, subtitle_y), subtitle, fill=text_color, font=subtitle_font, anchor="mm")
except TypeError:
    # Fallback for older PIL versions without anchor support
    draw.text((width//2 - 150, title_y - 40), title, fill=text_color, font=title_font)
    draw.text((width//2 - 200, subtitle_y - 20), subtitle, fill=text_color, font=subtitle_font)

# Save the image
output_path = os.path.join("static", "og-image.png")
image.save(output_path)
print(f"✅ og-image.png created at {output_path}")
