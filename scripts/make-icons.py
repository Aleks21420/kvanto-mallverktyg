# Ritar Mallverktygets ikoner (public/icons/) i de storlekar Excel använder.
# Körs med: npm run ikoner (kräver Pillow).
from PIL import Image, ImageDraw

BLUE = (31, 78, 121, 255)  # byråns färg i förslaget
LIGHT = (205, 226, 251, 255)
DARK = (60, 60, 60, 255)
GREEN = (33, 115, 70, 255)
WHITE = (255, 255, 255, 255)


def make(path, draw, sizes):
    big = Image.new("RGBA", (320, 320), (0, 0, 0, 0))
    draw(ImageDraw.Draw(big))
    for size in sizes:
        big.resize((size, size), Image.LANCZOS).save(path.format(size=size))


def paper(d):
    # Ett arbetspapper: huvud i byråns färg och rader under.
    d.rounded_rectangle((50, 20, 270, 300), 20, fill=WHITE, outline=DARK, width=18)
    d.rectangle((59, 29, 261, 105), fill=BLUE)
    for y in (150, 195, 240):
        d.line((90, y, 230, y), fill=DARK, width=16)


def logo(d):
    d.rounded_rectangle((0, 0, 320, 320), 60, fill=BLUE)
    d.rounded_rectangle((80, 50, 240, 270), 14, fill=WHITE)
    d.rectangle((80, 50, 240, 110), fill=LIGHT)
    for y in (150, 190, 230):
        d.line((105, y, 215, y), fill=BLUE, width=14)


def prepared(d):
    paper(d)
    # En penna snett över papperet.
    d.line((150, 290, 290, 150), fill=GREEN, width=44)
    d.polygon([(122, 318), (138, 268), (172, 302)], fill=GREEN)


def reviewed(d):
    paper(d)
    d.ellipse((150, 150, 310, 310), fill=GREEN)
    d.line([(185, 232), (220, 268), (280, 195)], fill=WHITE, width=24, joint="curve")


def template(d):
    d.rounded_rectangle((100, 10, 300, 250), 18, fill=LIGHT, outline=DARK, width=16)
    d.rounded_rectangle((20, 70, 220, 310), 18, fill=WHITE, outline=DARK, width=18)
    d.rectangle((29, 79, 211, 140), fill=BLUE)
    for y in (185, 230, 270):
        d.line((55, y, 185, y), fill=DARK, width=14)


make("public/icons/icon-{size}.png", logo, (16, 32, 64, 80, 128))
for name, draw in (("anpassa", paper), ("upprattat", prepared), ("granskat", reviewed), ("mall", template)):
    make(f"public/icons/band/{name}-{{size}}.png", draw, (16, 20, 24, 32, 40, 48, 64, 80))
