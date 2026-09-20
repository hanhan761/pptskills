#!/usr/bin/env python3
"""Enforce the PPT6 GPT-6-only model gate."""

from __future__ import annotations

import argparse
import json
import re
import sys


GPT6_PATTERN = re.compile(r"^gpt[-_ ]?6(?:$|[-._])", re.IGNORECASE)


def main() -> int:
    parser = argparse.ArgumentParser(description="Check whether an active model may run PPT6")
    parser.add_argument("--model-id", required=True)
    args = parser.parse_args()
    model_id = args.model_id.strip()
    passed = bool(model_id) and bool(GPT6_PATTERN.match(model_id))
    result = {
        "skill": "ppt6",
        "model_id": model_id or None,
        "passed": passed,
        "reason": "gpt6-compatible model" if passed else "PPT6 requires a verifiable GPT-6 model ID",
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if passed else 2


if __name__ == "__main__":
    sys.exit(main())
