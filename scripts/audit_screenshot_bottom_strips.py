from __future__ import annotations

import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont, ImageStat


SCREENSHOT_DIR = Path(r"D:\DATEN\HOMEPAGES\x-erp.de\de\help\Screenshots")
OUT_DIR = Path(r"C:\Users\micha\Documents\X-ERP-HELP\outputs\help-seo\screenshot-bottom-audit")
STRIP_HEIGHT = 170
THUMB_W = 360
THUMB_H = 86
COLS = 3


def load_font(size: int = 14):
    for path in [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
    ]:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


def score_strip(strip: Image.Image) -> dict:
    rgb = strip.convert("RGB")
    pixels = list(rgb.getdata())
    total = len(pixels) or 1
    redish = sum(1 for r, g, b in pixels if r > 120 and g < 95 and b < 95)
    orangeish = sum(1 for r, g, b in pixels if r > 150 and 75 < g < 150 and b < 90)
    greenish = sum(1 for r, g, b in pixels if g > 115 and r < 100 and b < 115)
    yellowish = sum(1 for r, g, b in pixels if r > 160 and g > 135 and b < 80)
    bright = sum(1 for r, g, b in pixels if r > 175 and g > 175 and b > 175)
    stat = ImageStat.Stat(rgb)
    return {
        "redish_ratio": redish / total,
        "orangeish_ratio": orangeish / total,
        "greenish_ratio": greenish / total,
        "yellowish_ratio": yellowish / total,
        "bright_ratio": bright / total,
        "mean": stat.mean,
        "stdev": stat.stddev,
    }


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(SCREENSHOT_DIR.glob("*.webp"), key=lambda p: p.name.lower())
    font = load_font(14)
    label_font = load_font(13)
    records = []
    thumbs = []

    for path in files:
        try:
            with Image.open(path) as img:
                img = img.convert("RGB")
                w, h = img.size
                crop_h = min(STRIP_HEIGHT, h)
                strip = img.crop((0, h - crop_h, w, h))
                score = score_strip(strip)
                strip.thumbnail((THUMB_W, THUMB_H), Image.Resampling.LANCZOS)
                canvas = Image.new("RGB", (THUMB_W, THUMB_H + 38), "white")
                canvas.paste(strip, (0, 0))
                draw = ImageDraw.Draw(canvas)
                draw.text((4, THUMB_H + 3), path.name[:48], fill=(0, 0, 0), font=label_font)
                draw.text(
                    (4, THUMB_H + 20),
                    (
                        f"r {score['redish_ratio']:.4f} o {score['orangeish_ratio']:.4f} "
                        f"g {score['greenish_ratio']:.4f} y {score['yellowish_ratio']:.4f} "
                        f"b {score['bright_ratio']:.3f}"
                    ),
                    fill=(90, 90, 90),
                    font=label_font,
                )
                thumbs.append((path, canvas, score))
                records.append({"file": path.name, "width": w, "height": h, "bytes": path.stat().st_size, **score})
        except Exception as exc:
            records.append({"file": path.name, "error": str(exc)})

    records_sorted = sorted(
        records,
        key=lambda r: (
            r.get("redish_ratio", 0)
            + r.get("orangeish_ratio", 0)
            + r.get("greenish_ratio", 0)
            + r.get("yellowish_ratio", 0),
            r.get("bright_ratio", 0),
        ),
        reverse=True,
    )
    (OUT_DIR / "bottom-strip-scores.json").write_text(json.dumps(records_sorted, indent=2), encoding="utf-8")

    page_size = COLS * THUMB_W, 8 * (THUMB_H + 38)
    sorted_thumbs = sorted(
        thumbs,
        key=lambda item: (
            item[2]["redish_ratio"]
            + item[2]["orangeish_ratio"]
            + item[2]["greenish_ratio"]
            + item[2]["yellowish_ratio"],
            item[2]["bright_ratio"],
        ),
        reverse=True,
    )
    for page_index in range(0, len(sorted_thumbs), COLS * 8):
        batch = sorted_thumbs[page_index:page_index + COLS * 8]
        sheet = Image.new("RGB", page_size, (245, 247, 250))
        for i, (_path, thumb, _score) in enumerate(batch):
            x = (i % COLS) * THUMB_W
            y = (i // COLS) * (THUMB_H + 38)
            sheet.paste(thumb, (x, y))
        sheet.save(OUT_DIR / f"bottom-strips-{page_index // (COLS * 8) + 1:02d}.png")

    print(json.dumps({
        "screenshots": len(files),
        "out_dir": str(OUT_DIR),
        "contact_sheets": len(list(OUT_DIR.glob("bottom-strips-*.png"))),
        "top_suspicious": records_sorted[:20],
    }, indent=2))


if __name__ == "__main__":
    main()
