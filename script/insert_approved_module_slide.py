#!/usr/bin/env python3
"""Insert one approved content module into an existing family deck.

This is a mechanical revision helper for user-adjusted decks. It duplicates an
existing body shell inside the same presentation, removes only the variable body
content, mounts the exact registered module shapes, and writes one native notes
body. It does not redraw or relayout any slide.
"""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path
from typing import Any

from template_ppt import (
    fail,
    mount_external_module_from_presentation,
    powerpoint_app,
    read_json,
    sha256,
)


PPTX_FORMAT = 24


def resolve(path_text: str) -> Path:
    return Path(path_text).expanduser().resolve()


def set_native_note(slide: Any, text: str) -> None:
    body = None
    for index in range(1, int(slide.NotesPage.Shapes.Count) + 1):
        shape = slide.NotesPage.Shapes(index)
        try:
            if int(shape.Type) == 14 and int(shape.PlaceholderFormat.Type) == 2:
                body = shape
                break
        except Exception:
            continue
    if body is None:
        fail(f"Slide {slide.SlideIndex} has no native notes body placeholder")
    body.TextFrame.TextRange.Text = text


def actual_master_count(presentation: Any) -> int:
    return int(presentation.Designs.Count)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Insert one approved module after a body slide in an existing deck."
    )
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--module", required=True)
    parser.add_argument("--family", required=True)
    parser.add_argument("--after-slide", required=True, type=int)
    parser.add_argument("--note", required=True)
    parser.add_argument("--overwrite", action="store_true")
    args = parser.parse_args()

    input_path = resolve(args.input)
    output_path = resolve(args.output)
    module_path = resolve(args.module)
    family_path = resolve(args.family)
    if not input_path.is_file():
        fail(f"Input deck does not exist: {input_path}")
    if output_path.exists() and not args.overwrite:
        fail(f"Output already exists; pass --overwrite: {output_path}")

    module = read_json(module_path)
    family = read_json(family_path)
    family_id = str(family.get("family_id", ""))
    if family_id not in module.get("compatible_master_families", []):
        fail(f"Module is not compatible with master family '{family_id}'")
    module_source = (module_path.parents[3] / module["source"]["deck"]).resolve()
    if not module_source.is_file():
        # Contracts normally store a repository-relative path. Resolve from cwd
        # as a fallback when the module file is outside the standard tree.
        module_source = resolve(module["source"]["deck"])
    if sha256(module_source).upper() != str(module["source"]["sha256"]).upper():
        fail("Module source hash differs from its approved contract")

    fixed_shell_names = {
        str(name)
        for name in family.get("shell_contract", {}).get("fixed_slide_shapes", [])
    }
    if not fixed_shell_names:
        fail("Master family has no fixed slide-shape contract")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary_output = output_path.with_name(output_path.stem + ".building.pptx")
    if temporary_output.exists():
        temporary_output.unlink()
    shutil.copy2(input_path, temporary_output)

    result: dict[str, Any] = {}
    try:
        with powerpoint_app() as app:
            presentation = app.Presentations.Open(
                str(temporary_output), False, False, False
            )
            source_presentation = None
            try:
                before_slides = int(presentation.Slides.Count)
                if args.after_slide < 1 or args.after_slide > before_slides:
                    fail(
                        f"--after-slide {args.after_slide} is outside 1..{before_slides}"
                    )
                before_masters = actual_master_count(presentation)
                shell_slide = presentation.Slides(args.after_slide)
                shell_layout_name = str(shell_slide.CustomLayout.Name)
                shell_master_name = str(shell_slide.Design.Name)

                duplicated = shell_slide.Duplicate()
                new_slide = duplicated.Item(1)
                expected_index = args.after_slide + 1
                if int(new_slide.SlideIndex) != expected_index:
                    new_slide.MoveTo(expected_index)
                    new_slide = presentation.Slides(expected_index)

                source_presentation = app.Presentations.Open(
                    str(module_source), True, False, False
                )
                mount_external_module_from_presentation(
                    new_slide, module, source_presentation, fixed_shell_names
                )
                set_native_note(new_slide, args.note)

                if int(presentation.Slides.Count) != before_slides + 1:
                    fail("Slide insertion changed the deck by more than one slide")
                if str(new_slide.CustomLayout.Name) != shell_layout_name:
                    fail("Inserted slide layout differs from its duplicated family shell")
                if str(new_slide.Design.Name) != shell_master_name:
                    fail("Inserted slide design differs from its duplicated family shell")

                presentation.Save()
                result = {
                    "ok": True,
                    "input": str(input_path),
                    "output": str(output_path),
                    "slide_count_before": before_slides,
                    "slide_count_after": int(presentation.Slides.Count),
                    "inserted_slide": int(new_slide.SlideIndex),
                    "inserted_slide_id": int(new_slide.SlideID),
                    "master_count_before": before_masters,
                    "master_count_after": actual_master_count(presentation),
                    "layout": shell_layout_name,
                    "module": str(module_path),
                    "module_shape_count": len(module.get("shape_refs", [])),
                    "fixed_shell_shapes": sorted(fixed_shell_names),
                    "notes_written": True,
                }
            finally:
                if source_presentation is not None:
                    try:
                        source_presentation.Close()
                    except Exception:
                        pass
                try:
                    presentation.Close()
                except Exception:
                    pass
        shutil.copy2(temporary_output, output_path)
    finally:
        if temporary_output.exists():
            temporary_output.unlink()

    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
