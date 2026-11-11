"""
Generate assets/images/logo.icns with real macOS app-icon geometry.

A Windows .ico is a full-bleed square and looks right that way. macOS is the
opposite: the system draws no mask for a .icns, so an icon that fills its
canvas sits in the Dock as a hard-edged square among rounded ones, and reads
as larger than its neighbours because they all inset themselves.

Apple's template puts the icon body in the middle 824px of a 1024px canvas -
100px of clear space on every side - with continuous ("squircle") corners of
radius 185.4px and a soft drop shadow. This script rebuilds the .icns to that
spec from the source artwork.

    python tools/make_macos_icon.py                      # rebuild from logo.icns
    python tools/make_macos_icon.py --source art.png     # from other artwork
    python tools/make_macos_icon.py --background "#fff"  # override tile colour

Needs Pillow, and iconutil (ships with macOS).
"""

import argparse
import math
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter

PROJECT_ROOT = Path(__file__).resolve().parent.parent
IMAGES = PROJECT_ROOT / "assets" / "images"

CANVAS = 1024           # Apple's icon canvas
BODY = 824              # the rounded tile inside it
CORNER_RADIUS = 185.4   # at 1024
CORNER_SMOOTHING = 0.6  # Apple's continuous-corner value
ARTWORK_FILL = 0.86     # fraction of the tile the artwork spans
SUPERSAMPLE = 4         # drawn big, then downscaled, for clean antialiasing

# Every size a complete .iconset carries. iconutil rejects unexpected names.
ICONSET = [
    ("icon_16x16.png", 16),
    ("icon_16x16@2x.png", 32),
    ("icon_32x32.png", 32),
    ("icon_32x32@2x.png", 64),
    ("icon_128x128.png", 128),
    ("icon_128x128@2x.png", 256),
    ("icon_256x256.png", 256),
    ("icon_256x256@2x.png", 512),
    ("icon_512x512.png", 512),
    ("icon_512x512@2x.png", 1024),
]


# ----------------------------------------------------------------------
# the squircle
# ----------------------------------------------------------------------
def _arc_points(start, end, radius, steps):
    """Sample the short clockwise arc of `radius` from `start` to `end`."""
    (x0, y0), (x1, y1) = start, end
    mx, my = (x0 + x1) / 2, (y0 + y1) / 2
    dx, dy = x1 - x0, y1 - y0
    chord = math.hypot(dx, dy)
    # Distance from the chord's midpoint out to the centre. Clamped because
    # rounding can push the ratio a hair over the radius.
    h = math.sqrt(max(radius * radius - (chord / 2) ** 2, 0.0))
    # Two centres sit at that distance; the corner needs the one on the inside
    # of the turn, or the arc bulges out past the edges instead of rounding in.
    cx, cy = mx - h * (dy / chord), my + h * (dx / chord)

    a0 = math.atan2(y0 - cy, x0 - cx)
    a1 = math.atan2(y1 - cy, x1 - cx)
    while a1 < a0:
        a1 += 2 * math.pi
    return [
        (cx + radius * math.cos(a0 + (a1 - a0) * i / steps),
         cy + radius * math.sin(a0 + (a1 - a0) * i / steps))
        for i in range(steps + 1)
    ]


def _bezier_points(p0, p1, p2, p3, steps):
    out = []
    for i in range(steps + 1):
        t = i / steps
        u = 1 - t
        out.append((
            u**3 * p0[0] + 3 * u*u*t * p1[0] + 3 * u*t*t * p2[0] + t**3 * p3[0],
            u**3 * p0[1] + 3 * u*u*t * p1[1] + 3 * u*t*t * p2[1] + t**3 * p3[1],
        ))
    return out


def squircle_polygon(side, radius, smoothing=CORNER_SMOOTHING, steps=48):
    """
    Points tracing a rounded square with Apple's continuous corners.

    A plain rounded rectangle joins a straight edge to a circular arc, which
    leaves a visible kink in the curvature. Apple eases into the arc with a
    bezier on each side of it; `smoothing` sets how much of the corner that
    easing takes over. The construction below is the standard decomposition
    of that shape into bezier - arc - bezier.
    """
    radius = min(radius, side / 2)
    p = min((1 + smoothing) * radius, side / 2)

    arc_measure = math.radians(90 * (1 - smoothing))
    arc_len = math.sin(arc_measure / 2) * radius * math.sqrt(2)
    angle_alpha = (math.pi / 2 - arc_measure) / 2
    angle_beta = math.radians(45 * smoothing)

    c = radius * math.tan(angle_alpha / 2) * math.cos(angle_beta)
    d = c * math.tan(angle_beta)
    b = (p - arc_len - c - d) / 3
    a = 2 * b

    # Top-right corner, clockwise: leaves the top edge at (side - p, 0) and
    # rejoins the right edge at (side, p).
    x = side - p
    ease_in = _bezier_points(
        (x, 0), (x + a, 0), (x + a + b, 0), (x + a + b + c, d), steps
    )
    arc_start = ease_in[-1]
    arc_end = (arc_start[0] + arc_len, arc_start[1] + arc_len)
    arc = _arc_points(arc_start, arc_end, radius, steps)
    ease_out = _bezier_points(
        arc_end,
        (arc_end[0] + d, arc_end[1] + c),
        (arc_end[0] + d, arc_end[1] + b + c),
        (arc_end[0] + d, arc_end[1] + a + b + c),
        steps,
    )
    corner = ease_in + arc[1:] + ease_out[1:]

    # The other three are this one rotated about the centre; a quarter turn
    # clockwise maps (x, y) to (side - y, x).
    points = []
    for turn in range(4):
        rotated = corner
        for _ in range(turn):
            rotated = [(side - py, px) for px, py in rotated]
        points.extend(rotated)
    return points


def squircle_mask(size, radius_ratio=CORNER_RADIUS / BODY):
    """A single-channel mask of the squircle, antialiased by supersampling."""
    big = size * SUPERSAMPLE
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).polygon(
        squircle_polygon(big, radius_ratio * big), fill=255
    )
    return mask.resize((size, size), Image.LANCZOS)


# ----------------------------------------------------------------------
# artwork
# ----------------------------------------------------------------------
def trim_to_content(image, tolerance=12):
    """Crop the flat border off `image`, returning the artwork and that colour."""
    rgb = image.convert("RGB")
    background = rgb.getpixel((1, 1))
    flat = Image.new("RGB", rgb.size, background)
    diff = ImageChops.difference(rgb, flat).convert("L")
    bbox = diff.point(lambda v: 255 if v > tolerance else 0).getbbox()
    return (image.crop(bbox) if bbox else image), background


def load_source(path):
    """Read the artwork, going through sips for .icns."""
    if path.suffix.lower() != ".icns":
        return Image.open(path).convert("RGBA")
    with tempfile.TemporaryDirectory() as tmp:
        png = Path(tmp) / "source.png"
        subprocess.run(
            ["sips", "-s", "format", "png", str(path), "--out", str(png)],
            check=True, capture_output=True,
        )
        return Image.open(png).convert("RGBA")


def build_icon(artwork, background):
    """Compose the 1024px master: shadow, tile, artwork, all masked."""
    canvas = Image.new("RGBA", (CANVAS, CANVAS), (0, 0, 0, 0))
    inset = (CANVAS - BODY) // 2
    mask = squircle_mask(BODY)

    # Apple bakes a soft shadow under the tile; the Dock's own shadow sits on
    # top of it. Kept subtle so the two do not compound into a smudge.
    shadow = Image.new("RGBA", (CANVAS, CANVAS), (0, 0, 0, 0))
    shadow.paste((0, 0, 0, 64), (inset, inset + 10), mask)
    canvas.alpha_composite(shadow.filter(ImageFilter.GaussianBlur(10)))

    tile = Image.new("RGBA", (BODY, BODY), background + (255,))

    # Scale the artwork to sit inside the tile with breathing room, keeping
    # its aspect ratio.
    span = int(BODY * ARTWORK_FILL)
    art = artwork.copy()
    art.thumbnail((span, span), Image.LANCZOS)
    tile.alpha_composite(art, ((BODY - art.width) // 2, (BODY - art.height) // 2))

    tile.putalpha(mask)
    canvas.alpha_composite(tile, (inset, inset))
    return canvas


def write_icns(master, destination):
    with tempfile.TemporaryDirectory() as tmp:
        iconset = Path(tmp) / "logo.iconset"
        iconset.mkdir()
        for name, size in ICONSET:
            master.resize((size, size), Image.LANCZOS).save(iconset / name)
        subprocess.run(
            ["iconutil", "-c", "icns", str(iconset), "-o", str(destination)],
            check=True,
        )


def parse_colour(text):
    text = text.lstrip("#")
    if len(text) == 3:
        text = "".join(ch * 2 for ch in text)
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[1])
    parser.add_argument(
        "--source", type=Path, default=IMAGES / "logo.icns",
        help="artwork to build from (default: the existing logo.icns)",
    )
    parser.add_argument(
        "--output", type=Path, default=IMAGES / "logo.icns",
        help="where to write the .icns",
    )
    parser.add_argument(
        "--background", type=parse_colour, default=None,
        help="tile colour, e.g. '#ffffff' (default: sampled from the artwork)",
    )
    parser.add_argument(
        "--preview", type=Path, default=None,
        help="also write the 1024px master here as a PNG",
    )
    args = parser.parse_args()

    if sys.platform != "darwin" or shutil.which("iconutil") is None:
        sys.exit("This script needs macOS (it shells out to iconutil and sips).")

    artwork, sampled = trim_to_content(load_source(args.source))
    master = build_icon(artwork, args.background or sampled)

    if args.preview:
        master.save(args.preview)

    # Build to a temp file first: a failed iconutil must not leave a truncated
    # .icns behind, and source and output are the same file by default.
    with tempfile.TemporaryDirectory() as tmp:
        staged = Path(tmp) / "logo.icns"
        write_icns(master, staged)
        args.output.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(staged, args.output)

    print(f"Wrote {args.output} ({args.output.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()
