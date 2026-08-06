from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


def page_number(path: Path) -> int:
    numbers = re.findall(r"\d+", path.stem)
    return int(numbers[-1]) if numbers else 0


def main() -> None:
    parser = argparse.ArgumentParser(description="Build numbered PPT render contact sheets.")
    parser.add_argument("--input-dir", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--per-sheet", type=int, default=20)
    parser.add_argument("--columns", type=int, default=4)
    parser.add_argument("--thumb-width", type=int, default=320)
    args = parser.parse_args()

    images = sorted(args.input_dir.glob("*.png"), key=page_number)
    if not images:
        raise SystemExit(f"No PNG renders found in {args.input_dir}")
    args.output_dir.mkdir(parents=True, exist_ok=True)
    font = ImageFont.load_default(size=18)
    label_height = 28
    gap = 10

    for start in range(0, len(images), args.per_sheet):
        batch = images[start : start + args.per_sheet]
        with Image.open(batch[0]) as sample:
            thumb_height = round(args.thumb_width * sample.height / sample.width)
        rows = math.ceil(len(batch) / args.columns)
        sheet_width = args.columns * args.thumb_width + (args.columns + 1) * gap
        sheet_height = rows * (thumb_height + label_height) + (rows + 1) * gap
        sheet = Image.new("RGB", (sheet_width, sheet_height), "white")
        draw = ImageDraw.Draw(sheet)
        for index, path in enumerate(batch):
            row, column = divmod(index, args.columns)
            x = gap + column * (args.thumb_width + gap)
            y = gap + row * (thumb_height + label_height + gap)
            with Image.open(path) as image:
                thumb = image.convert("RGB").resize(
                    (args.thumb_width, thumb_height), Image.Resampling.LANCZOS
                )
            sheet.paste(thumb, (x, y))
            draw.text((x + 6, y + thumb_height + 3), f"P{page_number(path):03d}", fill="black", font=font)
        first = page_number(batch[0])
        last = page_number(batch[-1])
        output = args.output_dir / f"contact-{first:03d}-{last:03d}.png"
        sheet.save(output, optimize=True)
        print(output)


if __name__ == "__main__":
    main()
