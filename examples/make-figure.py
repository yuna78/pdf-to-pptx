#!/usr/bin/env python3
"""Regenerate the README figure: original PDF vs LibreOffice alone vs pdf-to-pptx.

Usage:  .venv/bin/python examples/make-figure.py [--page 3] [-o examples/before-after.png]

Needs LibreOffice, poppler and Pillow. Maintenance tool, not part of the conversion.
"""
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]

PANEL_W = 1180
BG = (244, 246, 250)
INK = (26, 26, 36)
MUTED = (112, 120, 136)
GOOD = (26, 138, 90)
BAD = (200, 68, 47)
NEUTRAL = (120, 132, 152)
FONT_CANDIDATES = [
    # (path, face index, is_bold) — sans-serif CJK faces first; the label strip must not
    # fall back to a serif face or the figure looks like a different tool made it.
    ("/System/Library/Fonts/Hiragino Sans GB.ttc", 2, True),
    ("/System/Library/Fonts/Hiragino Sans GB.ttc", 0, False),
    ("/System/Library/Fonts/PingFang.ttc", 4, True),
    ("/System/Library/Fonts/PingFang.ttc", 2, False),
    ("/System/Library/Fonts/STHeiti Medium.ttc", 1, True),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 0, True),
    ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0, False),
    ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", 0, False),
]


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    for path, index, is_bold in FONT_CANDIDATES:
        if bold and not is_bold:
            continue
        if not Path(path).exists():
            continue
        try:
            return ImageFont.truetype(path, size, index=index)
        except Exception:
            continue
    for path, index, _ in FONT_CANDIDATES:            # any face beats the bitmap default
        if Path(path).exists():
            try:
                return ImageFont.truetype(path, size, index=index)
            except Exception:
                continue
    return ImageFont.load_default()


def soffice_binary() -> str:
    found = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not Path(found).exists():
        raise SystemExit("LibreOffice not found.")
    return found


def render_page(document: Path, page: int, work: Path, tag: str) -> Image.Image:
    """Render one page of a PDF/PPTX to an image at PANEL_W."""
    pdf = document
    if document.suffix.lower() == ".pptx":
        subprocess.run(
            [soffice_binary(), "--headless", f"-env:UserInstallation=file://{work}/profile-{tag}",
             "--convert-to", "pdf", "--outdir", str(work), str(document)],
            check=True, capture_output=True,
        )
        pdf = work / f"{document.stem}.pdf"
    prefix = work / f"render-{tag}"
    subprocess.run(
        ["pdftoppm", "-png", "-scale-to-x", str(PANEL_W), "-scale-to-y", "-1",
         "-f", str(page), "-l", str(page), str(pdf), str(prefix)],
        check=True, capture_output=True,
    )
    return Image.open(sorted(work.glob(f"render-{tag}-*.png"))[0]).convert("RGB")


def panel(image: Image.Image, caption: str, note: str, colour: tuple[int, int, int],
          crop: float) -> Image.Image:
    if crop < 1.0:
        image = image.crop((0, 0, image.width, int(image.height * crop)))
    strip = 46
    canvas = Image.new("RGB", (image.width + 4, image.height + strip + 4), BG)
    draw = ImageDraw.Draw(canvas)
    draw.rounded_rectangle((2, strip // 2 - 10, 22, strip // 2 + 10), 5, fill=colour)
    draw.text((32, strip // 2 - 12), caption, font=load_font(20, bold=True), fill=INK)
    offset = 32 + draw.textlength(caption, font=load_font(20, bold=True)) + 14
    draw.text((offset, strip // 2 - 9), note, font=load_font(16), fill=MUTED)
    canvas.paste(image, (2, strip))
    draw.rectangle((1, strip - 1, image.width + 2, strip + image.height), outline=(206, 216, 232))
    return canvas


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", type=Path, default=ROOT / "tests/fixtures/sample-deck.pdf")
    parser.add_argument("--page", type=int, default=3)
    parser.add_argument("--crop", type=float, default=0.74, help="keep this fraction of page height")
    parser.add_argument("-o", "--output", type=Path, default=ROOT / "examples/before-after.png")
    args = parser.parse_args()

    convert = ROOT / "scripts" / "convert.py"
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        fixed, raw = work / "fixed.pptx", work / "raw.pptx"
        subprocess.run([sys.executable, str(convert), str(args.pdf), "-o", str(fixed)],
                       check=True, capture_output=True)
        subprocess.run([sys.executable, str(convert), str(args.pdf), "-o", str(raw), "--no-postfix"],
                       check=True, capture_output=True)

        panels = [
            panel(render_page(args.pdf, args.page, work, "src"),
                  "原始 PDF", "参照", NEUTRAL, args.crop),
            panel(render_page(raw, args.page, work, "raw"),
                  "LibreOffice 直转", "文字撞行、字体被替换", BAD, args.crop),
            panel(render_page(fixed, args.page, work, "fixed"),
                  "pdf-to-pptx", "版式回到原样，文字仍可编辑", GOOD, args.crop),
        ]

        gap = 20
        height = sum(p.height for p in panels) + gap * (len(panels) - 1)
        figure = Image.new("RGB", (panels[0].width, height), BG)
        y = 0
        for p in panels:
            figure.paste(p, (0, y))
            y += p.height + gap
        args.output.parent.mkdir(parents=True, exist_ok=True)
        figure.save(args.output, optimize=True)
        print(f"{args.output}  ({figure.width}x{figure.height})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
