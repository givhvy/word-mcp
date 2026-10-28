"""Render the D0180 'System context' diagram (Figure 1) as a 1600x800 PNG in the house style.

Usage: python make_context_diagram.py <out.png>
Colours come from the original figure: fill #e4f3f3, border/lines #356c78, text #18343e, labels #50616b.
"""
import sys
from PIL import Image, ImageDraw, ImageFont

S = 2  # supersample
W, H = 1600, 800
FILL, EDGE, TEXT, LABEL = "#e4f3f3", "#356c78", "#18343e", "#50616b"


def font(size, bold=False):
    name = "arialbd.ttf" if bold else "arial.ttf"
    return ImageFont.truetype(name, size * S)


img = Image.new("RGB", (W * S, H * S), "white")
d = ImageDraw.Draw(img)


def box(x1, y1, x2, y2, text, size=26):
    d.rounded_rectangle([x1 * S, y1 * S, x2 * S, y2 * S], radius=16 * S, fill=FILL, outline=EDGE, width=3 * S)
    f = font(size)
    w = d.textlength(text, font=f)
    d.text(((x1 + x2) / 2 * S - w / 2, (y1 + y2) / 2 * S), text, font=f, fill=TEXT, anchor="lm")


def line(points, width=3):
    d.line([(x * S, y * S) for x, y in points], fill=EDGE, width=width * S, joint="curve")


def arrow_head(x, y, direction):
    a = 12
    if direction == "right":
        pts = [(x, y), (x - 2 * a, y - a), (x - 2 * a, y + a)]
    else:
        pts = [(x, y), (x - a, y - 2 * a), (x + a, y - 2 * a)]
    d.polygon([(px * S, py * S) for px, py in pts], fill=EDGE)


def label(cx, y, text, size=20, anchor="mm", x=None):
    f = font(size)
    d.text(((x if x is not None else cx) * S, y * S), text, font=f, fill=LABEL, anchor=anchor)


d.text((60 * S, 50 * S), "How external information reaches Mission 8", font=font(34, True), fill=TEXT, anchor="lm")

sources = ["VeloViewer", "Open-Meteo / Windy", "RainViewer", "Feniks AI"]
ys = [130, 240, 350, 460]
for name, y in zip(sources, ys):
    box(60, y, 340, y + 80, name)
    line([(340, y + 40), (420, y + 40)])
centres = [y + 40 for y in ys]
line([(420, centres[0]), (420, centres[-1])])
mid = (centres[0] + centres[-1]) // 2

box(560, mid - 50, 830, mid + 50, "Mission 8 adapters")
box(930, mid - 50, 1220, mid + 50, "Pulse data and feeds")
box(1330, mid - 50, 1560, mid + 50, "Mission 8 app")
box(1330, 620, 1560, 700, "Google Street View")

line([(420, mid), (560 - 24, mid)])
arrow_head(560, mid, "right")
line([(830, mid), (930 - 24, mid)])
arrow_head(930, mid, "right")
line([(1220, mid), (1330 - 24, mid)])
arrow_head(1330, mid, "right")
line([(1445, mid + 50), (1445, 620 - 24)])
arrow_head(1445, 620, "down")

label(490, mid - 22, "API / input")
label(880, mid - 22, "data")
label(1275, mid - 22, "read")
label(0, (mid + 50 + 620) / 2, "imagery request", anchor="rm", x=1425)
label(0, mid + 105, "Route imports go to Master Data. Weather and AI use Pulse messages and feeds.", anchor="lm", x=560)

img = img.resize((W, H), Image.LANCZOS)
img.save(sys.argv[1] if len(sys.argv) > 1 else "context_diagram.png")
