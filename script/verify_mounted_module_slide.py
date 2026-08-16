#!/usr/bin/env python3
"""Verify one mounted content module and its preserved family shell."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from template_ppt import (
    find_repo_root,
    module_geometry_records,
    mounted_geometry_records,
    open_presentation,
    read_json,
    resolve_repo_path,
)


def selector_order(records: list[dict[str, Any]]) -> list[tuple[str, int]]:
    return [(str(record["name"]), int(record["occurrence"])) for record in records]


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--slide", required=True, type=int)
    parser.add_argument("--module", required=True)
    parser.add_argument("--family", required=True)
    parser.add_argument("--shell-baseline", required=True)
    parser.add_argument("--shell-slide", required=True, type=int)
    args = parser.parse_args()

    repo = find_repo_root()
    input_path = resolve_repo_path(args.input, repo)
    module_path = resolve_repo_path(args.module, repo)
    family_path = resolve_repo_path(args.family, repo)
    shell_path = resolve_repo_path(args.shell_baseline, repo)
    module = read_json(module_path)
    family = read_json(family_path)
    source_path = resolve_repo_path(module["source"]["deck"], repo)

    output_deck = open_presentation(input_path)
    source_deck = open_presentation(source_path)
    shell_deck = open_presentation(shell_path)
    errors: list[str] = []
    if args.slide < 1 or args.slide > len(output_deck.slides):
        errors.append("target slide is outside the output deck")
    if args.shell_slide < 1 or args.shell_slide > len(shell_deck.slides):
        errors.append("shell slide is outside the baseline deck")
    if errors:
        print(json.dumps({"ok": False, "errors": errors}, ensure_ascii=False, indent=2))
        raise SystemExit(2)

    source_slide = source_deck.slides[int(module["source"]["slide"]) - 1]
    target_slide = output_deck.slides[args.slide - 1]
    shell_slide = shell_deck.slides[args.shell_slide - 1]
    shape_refs = module["shape_refs"]
    source_records, _, source_errors = module_geometry_records(source_slide, shape_refs)
    target_records, _, target_errors = module_geometry_records(target_slide, shape_refs)
    errors.extend(source_errors)
    errors.extend(target_errors)

    source_geometry = mounted_geometry_records(source_records)
    target_geometry = mounted_geometry_records(target_records)
    if source_geometry != target_geometry:
        errors.append("mounted module type/geometry differs from the approved source")
    if selector_order(source_records) != selector_order(target_records):
        errors.append("mounted module relative z-order differs from the approved source")

    fixed_names = [
        str(value)
        for value in family.get("shell_contract", {}).get("fixed_slide_shapes", [])
    ]
    fixed_refs = [{"name": name, "occurrence": 1} for name in fixed_names]
    shell_records, _, shell_errors = module_geometry_records(shell_slide, fixed_refs)
    target_shell_records, _, target_shell_errors = module_geometry_records(
        target_slide, fixed_refs
    )
    errors.extend(shell_errors)
    errors.extend(target_shell_errors)
    if mounted_geometry_records(shell_records) != mounted_geometry_records(
        target_shell_records
    ):
        errors.append("fixed family shell geometry differs from the frozen baseline")

    expected_top_level = len(shape_refs) + len(fixed_refs)
    actual_top_level = len(target_slide.shapes)
    if actual_top_level != expected_top_level:
        errors.append(
            f"target has {actual_top_level} top-level shapes; expected {expected_top_level}"
        )

    result = {
        "ok": not errors,
        "deck": str(input_path),
        "slide": args.slide,
        "module": str(module_path),
        "module_shapes": len(shape_refs),
        "fixed_shell_shapes": fixed_names,
        "top_level_shape_count": actual_top_level,
        "geometry_equal": source_geometry == target_geometry,
        "relative_z_order_equal": selector_order(source_records)
        == selector_order(target_records),
        "shell_geometry_equal": mounted_geometry_records(shell_records)
        == mounted_geometry_records(target_shell_records),
        "errors": errors,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(2)


if __name__ == "__main__":
    main()
