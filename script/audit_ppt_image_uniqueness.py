from __future__ import annotations

import argparse
import hashlib
import json
from collections import Counter
from pathlib import Path

from PIL import Image, ImageOps


def resolve(path_value: str, root: Path) -> Path:
    path = Path(path_value)
    return path if path.is_absolute() else root / path


def dhash(path: Path, size: int = 16) -> str:
    with Image.open(path) as source:
        image = ImageOps.exif_transpose(source).convert("L").resize(
            (size + 1, size), Image.Resampling.LANCZOS
        )
        pixels = list(image.getdata())
    value = 0
    for row in range(size):
        offset = row * (size + 1)
        for column in range(size):
            value = (value << 1) | int(
                pixels[offset + column] > pixels[offset + column + 1]
            )
    return f"{value:0{size * size // 4}x}"


def hamming(left: str, right: str) -> int:
    return (int(left, 16) ^ int(right, 16)).bit_count()


def repeated(values: list[str]) -> dict[str, int]:
    return {value: count for value, count in Counter(values).items() if count > 1}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Require one independent original asset per PPT image-source record."
    )
    parser.add_argument("--sources", required=True, type=Path)
    parser.add_argument("--repo-root", type=Path, default=Path.cwd())
    parser.add_argument("--near-threshold", type=int, default=4)
    parser.add_argument("--expected-count", type=int)
    parser.add_argument("--output", type=Path)
    parser.add_argument(
        "--allow-shared-visual-slots",
        action="store_true",
        help=(
            "Collapse records with the same non-empty shared_visual_slot when "
            "they point to the same file, caption, and source page."
        ),
    )
    args = parser.parse_args()

    root = args.repo_root.resolve()
    payload = json.loads(args.sources.read_text(encoding="utf-8"))
    raw_records = payload.get("images") if isinstance(payload, dict) else None
    if not isinstance(raw_records, list) or not raw_records:
        raise SystemExit("Sources JSON must contain a non-empty images list")

    records = []
    shared_slots: dict[str, dict] = {}
    shared_slot_errors: list[str] = []
    collapsed_shared_records = 0
    for record in raw_records:
        slot = str(record.get("shared_visual_slot", "") or "").strip()
        if not args.allow_shared_visual_slots or not slot:
            records.append(record)
            continue
        existing = shared_slots.get(slot)
        if existing is None:
            shared_slots[slot] = record
            records.append(record)
            continue
        identity = (
            str(record.get("file", "")),
            str(record.get("caption", "")),
            str(record.get("source_page", "")),
        )
        existing_identity = (
            str(existing.get("file", "")),
            str(existing.get("caption", "")),
            str(existing.get("source_page", "")),
        )
        if identity != existing_identity:
            shared_slot_errors.append(
                f"shared_visual_slot {slot!r} contains different source identities"
            )
        else:
            collapsed_shared_records += 1

    assets = []
    missing = []
    for index, record in enumerate(records, 1):
        path = resolve(str(record["file"]), root).resolve()
        if not path.is_file():
            missing.append(str(path))
            continue
        assets.append(
            {
                "record": index,
                "slide": int(record.get("slide", 0)),
                "shape": str(record.get("shape", "")),
                "file": path.relative_to(root).as_posix()
                if path.is_relative_to(root)
                else str(path),
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "dhash16": dhash(path),
                "source_identity": " | ".join(
                    [
                        str(record.get("source_page", "")).strip(),
                        str(record.get("caption", "")).strip(),
                    ]
                ),
            }
        )

    near_pairs = []
    for index, left in enumerate(assets):
        for right in assets[index + 1 :]:
            distance = hamming(left["dhash16"], right["dhash16"])
            if distance <= args.near_threshold:
                near_pairs.append(
                    {
                        "distance": distance,
                        "left": left,
                        "right": right,
                    }
                )

    path_reuse = repeated([item["file"] for item in assets])
    sha_reuse = repeated([item["sha256"] for item in assets])
    dhash_reuse = repeated([item["dhash16"] for item in assets])
    source_identity_reuse = repeated([item["source_identity"] for item in assets])
    expected_ok = args.expected_count is None or len(records) == args.expected_count
    result = {
        "raw_records": len(raw_records),
        "records": len(records),
        "collapsed_shared_records": collapsed_shared_records,
        "shared_visual_slots": sorted(shared_slots),
        "shared_visual_slot_errors": shared_slot_errors,
        "existing_assets": len(assets),
        "expected_count": args.expected_count,
        "unique_paths": len({item["file"] for item in assets}),
        "unique_sha256": len({item["sha256"] for item in assets}),
        "unique_dhash16": len({item["dhash16"] for item in assets}),
        "unique_source_identities": len(
            {item["source_identity"] for item in assets}
        ),
        "missing": missing,
        "path_reuse": path_reuse,
        "sha256_reuse": sha_reuse,
        "dhash16_reuse": dhash_reuse,
        "source_identity_reuse": source_identity_reuse,
        "near_threshold": args.near_threshold,
        "near_duplicate_pairs": near_pairs,
        "passed": (
            expected_ok
            and not shared_slot_errors
            and not missing
            and not path_reuse
            and not sha_reuse
            and not dhash_reuse
            and not source_identity_reuse
            and not near_pairs
        ),
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    if not result["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
