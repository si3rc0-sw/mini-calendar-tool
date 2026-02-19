"""Generate the system-tray icon (64×64 PIL Image, in-memory)."""

from datetime import date

from PIL import Image, ImageDraw, ImageFont


def create_icon_image() -> Image.Image:
    """Return a 64×64 RGBA image: black week number on white, filling full height."""
    size = 64
    img = Image.new("RGBA", (size, size), "white")
    draw = ImageDraw.Draw(img)

    today = date.today()
    iso_week = str(today.isocalendar()[1])

    # Find the largest font size that fits the icon
    font_size = 120
    font = None
    while font_size > 10:
        try:
            font = ImageFont.truetype("segoeuib.ttf", font_size)
        except OSError:
            font = ImageFont.load_default()
            break
        bbox = draw.textbbox((0, 0), iso_week, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        if tw <= size and th <= size:
            break
        font_size -= 1

    # Centre the actual visible pixels (compensate for font metric offsets)
    bbox = draw.textbbox((0, 0), iso_week, font=font)
    x = (size - (bbox[2] - bbox[0])) / 2 - bbox[0]
    y = (size - (bbox[3] - bbox[1])) / 2 - bbox[1]
    draw.text((x, y), iso_week, fill="black", font=font)

    return img
