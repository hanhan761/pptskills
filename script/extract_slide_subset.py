#!/usr/bin/env python3
"""Extract a multi-slide subset by copying a source deck and deleting other slides.

This is deliberately a mechanical operation: no slide, layout, master, theme, or
shape is created or redrawn.  It is the multi-slide counterpart of the project's
single-slide faithful extraction workflow.
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

import pythoncom
import win32com.client


def fail(message: str) -> None:
    raise SystemExit(f"ERROR: {message}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--keep", nargs="+", type=int, required=True)
    args = parser.parse_args()

    source = Path(args.input).resolve()
    output = Path(args.output).resolve()
    if not source.is_file():
        fail(f"Source deck not found: {source}")
    keep = sorted(set(args.keep))
    if len(keep) != len(args.keep):
        fail("Duplicate slide numbers in --keep")

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)

    pythoncom.CoInitialize()
    app = None
    presentation = None
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = True
        app.DisplayAlerts = 0
        presentation = app.Presentations.Open(str(output), False, False, False)
        count = int(presentation.Slides.Count)
        invalid = [number for number in keep if number < 1 or number > count]
        if invalid:
            fail(f"Slide numbers outside 1..{count}: {invalid}")
        keep_set = set(keep)
        for number in range(count, 0, -1):
            if number not in keep_set:
                presentation.Slides.Item(number).Delete()
        presentation.Save()
        if int(presentation.Slides.Count) != len(keep):
            fail(
                f"Extraction produced {presentation.Slides.Count} slides; "
                f"expected {len(keep)}"
            )
    finally:
        if presentation is not None:
            presentation.Close()
        if app is not None:
            app.Quit()
        pythoncom.CoUninitialize()

    print(f"Extracted {len(keep)} source slide(s) to {output}")


if __name__ == "__main__":
    main()
