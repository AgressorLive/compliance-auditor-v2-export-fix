"""Generate a random handwriting/document-themed icon for DocGenerator.exe"""
import random
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont


def _bezier(draw, p0, p1, p2, p3, color, width=4, steps=50):
    pts = []
    for i in range(steps + 1):
        t = i / steps
        mt = 1 - t
        x = mt**3*p0[0] + 3*mt**2*t*p1[0] + 3*mt*t**2*p2[0] + t**3*p3[0]
        y = mt**3*p0[1] + 3*mt**2*t*p1[1] + 3*mt*t**2*p2[1] + t**3*p3[1]
        pts.append((x, y))
    for i in range(len(pts) - 1):
        draw.line([pts[i], pts[i+1]], fill=color, width=width)


def generate_icon(out_path: Path) -> None:
    random.seed()          # truly random each run
    out_path.parent.mkdir(parents=True, exist_ok=True)

    SIZE = 256

    # ── Random colour palette ─────────────────────────────────────────────────
    palettes = [
        # bg,           card,          accent,        ink
        ((17, 17, 27),  (30, 30, 46),  (137,180,250), (0,0,0)),     # dark blue
        ((22, 33, 22),  (30, 46, 30),  (166,227,161), (0,0,0)),     # dark green
        ((33, 17, 17),  (46, 30, 30),  (243,139,168), (0,0,0)),     # dark red
        ((27, 22, 33),  (40, 30, 50),  (203,166,247), (0,0,0)),     # dark purple
        ((22, 30, 33),  (30, 42, 46),  (137,220,235), (0,0,0)),     # dark teal
    ]
    bg_col, card_col, accent_col, ink_col = random.choice(palettes)
    ink_alpha = (ink_col[0], ink_col[1], ink_col[2], 220)

    img = Image.new("RGBA", (SIZE, SIZE), bg_col + (255,))
    draw = ImageDraw.Draw(img)

    # ── Document card ─────────────────────────────────────────────────────────
    # Folded corner page
    cx0, cy0, cx1, cy1 = 44, 28, 212, 228
    fold = 36
    # Main body (without folded corner)
    card_pts = [
        (cx0, cy0),
        (cx1 - fold, cy0),
        (cx1, cy0 + fold),
        (cx1, cy1),
        (cx0, cy1),
    ]
    draw.polygon(card_pts, fill=card_col + (255,))
    # Folded triangle
    draw.polygon([
        (cx1 - fold, cy0),
        (cx1, cy0 + fold),
        (cx1 - fold, cy0 + fold),
    ], fill=accent_col + (180,))
    # Card border
    draw.polygon(card_pts, outline=accent_col + (200,), width=3)

    # ── Text lines on card ────────────────────────────────────────────────────
    line_x0, line_x1 = cx0 + 18, cx1 - 22
    line_cols = [
        accent_col + (220,),
        accent_col + (140,),
        accent_col + (140,),
        accent_col + (100,),
    ]
    line_ys = [cy0 + 52, cy0 + 76, cy0 + 96, cy0 + 116]
    widths   = [0.80,     0.65,     0.70,      0.50]
    for ly, lw, lc in zip(line_ys, widths, line_cols):
        draw.rounded_rectangle(
            (line_x0, ly - 4, line_x0 + int((line_x1 - line_x0) * lw), ly + 4),
            radius=4, fill=lc
        )

    # ── Handwriting flourish signature ────────────────────────────────────────
    sy = cy0 + 150
    sx0 = cx0 + 20
    sx1 = cx1 - 50

    # Signature stroke (Bézier)
    _bezier(draw,
            (sx0,                  sy),
            (sx0 + (sx1-sx0)*0.3,  sy - 14),
            (sx0 + (sx1-sx0)*0.7,  sy + 8),
            (sx1,                  sy),
            ink_alpha, width=5)

    # End loop
    _bezier(draw,
            (sx1,        sy),
            (sx1 + 18,   sy - 22),
            (sx1 + 28,   sy - 30),
            (sx1 + 14,   sy + 6),
            ink_alpha[:3] + (150,), width=3)

    # Leading hook
    _bezier(draw,
            (sx0 - 12,  sy + 5),
            (sx0 - 6,   sy - 10),
            (sx0 - 2,   sy - 4),
            (sx0,       sy - 2),
            ink_alpha[:3] + (130,), width=3)

    # Underline after sig
    draw.line([(sx0, sy + 10), (sx1, sy + 10)],
              fill=ink_alpha[:3] + (80,), width=2)

    # ── Accent dot / stamp ────────────────────────────────────────────────────
    draw.ellipse((cx0 + 14, cy1 - 42, cx0 + 42, cy1 - 14),
                 fill=accent_col + (230,))

    # ── Save as multi-size ICO ────────────────────────────────────────────────
    img.save(
        out_path, format="ICO",
        sizes=[(256,256),(128,128),(64,64),(48,48),(32,32),(16,16)]
    )


if __name__ == "__main__":
    icon_file = Path(__file__).resolve().parent.parent / "assets" / "docgen.ico"
    generate_icon(icon_file)
    print(f"Icon generated: {icon_file}")
