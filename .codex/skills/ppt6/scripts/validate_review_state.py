#!/usr/bin/env python3
"""Validate the human approval gate before PPT6 enters editable production."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


def fail(message: str) -> int:
    print(json.dumps({"passed": False, "error": message}, ensure_ascii=False, indent=2))
    return 2


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a PPT6 review_state.json")
    parser.add_argument("--state", required=True, type=Path)
    parser.add_argument("--phase", choices=("draft", "build"), default="build")
    args = parser.parse_args()

    try:
        state = json.loads(args.state.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return fail(f"review state not found: {args.state}")
    except json.JSONDecodeError as exc:
        return fail(f"invalid JSON: {exc}")

    if state.get("schema_version") != 1:
        return fail("schema_version must be 1")
    if not str(state.get("run_id", "")).strip():
        return fail("run_id is required")
    pages = state.get("pages")
    if not isinstance(pages, list) or not pages:
        return fail("pages must be a non-empty list")

    if args.phase == "build":
        gate = state.get("model_gate") or {}
        if gate.get("passed") is not True:
            return fail("model_gate.passed must be true before build")
        if not gate.get("model_id"):
            return fail("model_gate.model_id is required before build")

    for index, page in enumerate(pages, start=1):
        if not isinstance(page, dict):
            return fail(f"pages[{index}] must be an object")
        page_id = page.get("page_id")
        if not str(page_id or "").strip():
            return fail(f"pages[{index}].page_id is required")
        if not str(page.get("draft_image", "")).strip():
            return fail(f"{page_id}: draft_image is required")
        if args.phase == "build":
            if page.get("status") != "approved":
                return fail(f"{page_id}: status must be approved before build")
            if not str(page.get("human_confirmation", "")).strip():
                return fail(f"{page_id}: human_confirmation is required")
            if not str(page.get("confirmed_at", "")).strip():
                return fail(f"{page_id}: confirmed_at is required")

    print(
        json.dumps(
            {
                "passed": True,
                "phase": args.phase,
                "run_id": state["run_id"],
                "approved_pages": len(pages),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
