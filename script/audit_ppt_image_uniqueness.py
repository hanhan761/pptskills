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


def source_identity(record: dict) -> tuple[str, str]:
    """Return a stable identity for the exact source visual.

    Captions are deliberately excluded. Changing a caption must not make the
    same source image look like a new asset. A manifest may provide an explicit
    source_id; otherwise prefer the direct asset URL/file and fall back to the
    source page. The fallback is conservative when one page contains several
    different visuals, so new manifests should always provide source_id or a
    direct source file URL.
    """

    for key in ("source_id", "source_file", "direct_url", "source_page"):
        value = str(record.get(key, "") or "").strip()
        if value:
            return f"{key}:{value}", key
    return "", "missing"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Require one independent original asset per PPT image-source record."
    )
    parser.add_argument("--sources", required=True, type=Path)
    parser.add_argument("--repo-root", type=Path, default=Path.cwd())
    parser.add_argument("--near-threshold", type=int, default=4)
    parser.add_argument("--expected-count", type=int)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    root = args.repo_root.resolve()
    payload = json.loads(args.sources.read_text(encoding="utf-8"))
    raw_records = payload.get("images") if isinstance(payload, dict) else None
    if not isinstance(raw_records, list) or not raw_records:
        raise SystemExit("Sources JSON must contain a non-empty images list")

    records = raw_records
    schema_errors: list[str] = []
    disallowed_shared_visual_slots: list[dict[str, object]] = []
    missing_source_identities: list[int] = []
    for index, record in enumerate(records, 1):
        if not isinstance(record, dict):
            schema_errors.append(f"record {index} is not an object")
            continue
        for field in ("file", "caption", "credit", "source_page"):
            if not str(record.get(field, "") or "").strip():
                schema_errors.append(f"record {index} is missing {field}")
        slot = str(record.get("shared_visual_slot", "") or "").strip()
        if slot:
            disallowed_shared_visual_slots.append(
                {"record": index, "shared_visual_slot": slot}
            )
        identity, _ = source_identity(record)
        if not identity:
            missing_source_identities.append(index)

    assets = []
    missing = []
    for index, record in enumerate(records, 1):
        if not isinstance(record, dict) or not str(record.get("file", "")).strip():
            continue
        path = resolve(str(record["file"]), root).resolve()
        if not path.is_file():
            missing.append(str(path))
            continue
        identity, identity_kind = source_identity(record)
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
                "source_identity": identity,
                "source_identity_kind": identity_kind,
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
    source_identity_reuse = repeated(
        [item["source_identity"] for item in assets if item["source_identity"]]
    )
    expected_ok = args.expected_count is None or len(records) == args.expected_count
    result = {
        "raw_records": len(raw_records),
        "records": len(records),
        "schema_errors": schema_errors,
        "disallowed_shared_visual_slots": disallowed_shared_visual_slots,
        "missing_source_identities": missing_source_identities,
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
            and not schema_errors
            and not disallowed_shared_visual_slots
            and not missing_source_identities
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
