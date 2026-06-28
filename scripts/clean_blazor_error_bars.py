from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from PIL import Image


HELP_ROOT = Path(r"D:\DATEN\HOMEPAGES\x-erp.de\de\help")
ARCHIVE_ROOT = Path(r"C:\Users\micha\Documents\X-ERP-HELP\ARCHIV")
BAR_MIN_RATIO = 0.45


def pale_error_ratio(row_pixels) -> float:
    total = len(row_pixels) or 1
    pale = sum(1 for r, g, b in row_pixels if r > 238 and g > 235 and 175 < b < 245)
    return pale / total


def detect_bar_start(img: Image.Image) -> int | None:
    rgb = img.convert("RGB")
    w, h = rgb.size
    candidate_rows = []
    for y in range(max(0, h - 120), h):
        row = [rgb.getpixel((x, y)) for x in range(0, w, 4)]
        if pale_error_ratio(row) >= BAR_MIN_RATIO:
            candidate_rows.append(y)
    if not candidate_rows:
        return None

    start = candidate_rows[0]
    if h - start < 18:
        return None
    return start


def dominant_dark_fill(img: Image.Image, y: int) -> tuple[int, int, int]:
    rgb = img.convert("RGB")
    w, h = rgb.size
    sample_y = max(0, min(h - 1, y - 8))
    samples = []
    for yy in range(max(0, sample_y - 10), sample_y + 1):
        for x in range(0, w, 16):
            r, g, b = rgb.getpixel((x, yy))
            if r < 80 and g < 85 and b < 90:
                samples.append((r, g, b))
    if not samples:
        return (31, 36, 39)
    samples.sort()
    return samples[len(samples) // 2]


def clean_file(path: Path, archive_dir: Path) -> dict | None:
    with Image.open(path) as img:
        img = img.convert("RGB")
        bar_start = detect_bar_start(img)
        if bar_start is None:
            return None

        rel = path.relative_to(HELP_ROOT)
        backup = archive_dir / rel
        backup.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, backup)

        fill = dominant_dark_fill(img, bar_start)
        for y in range(bar_start, img.height):
            for x in range(img.width):
                img.putpixel((x, y), fill)
        img.save(path, "WEBP", quality=82)

        return {
            "file": str(path),
            "archive": str(backup),
            "bar_start": bar_start,
            "filled_rows": img.height - bar_start,
            "fill": fill,
        }


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    archive_dir = ARCHIVE_ROOT / f"error-bar-screenshots-{stamp}"
    cleaned = []
    for path in HELP_ROOT.rglob("*.webp"):
        item = clean_file(path, archive_dir)
        if item:
            cleaned.append(item)

    print(json.dumps({
        "help_root": str(HELP_ROOT),
        "archive_dir": str(archive_dir),
        "cleaned": len(cleaned),
        "files": cleaned,
    }, indent=2))


if __name__ == "__main__":
    main()
