#!/usr/bin/env python3
"""Apply an explicit keep/merge/discard review to learned PPT templates.

This tool never creates or redraws a slide. It validates that every source page
was reviewed exactly once, removes derived non-representative one-slide files,
and writes the page-level decision back to the learning ledger.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any


def fail(message: str) -> None:
    raise SystemExit(f"ERROR: {message}")


def read_json(path: Path) -> dict[str, Any]:
    if not path.is_file():
        fail(f"JSON file not found: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        fail(f"Cannot read JSON {path}: {exc}")
    if not isinstance(data, dict):
        fail(f"JSON root must be an object: {path}")
    return data


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )


def repo_relative(value: str) -> str:
    return Path(value).as_posix()


def page_key(source: str, slide: int) -> tuple[str, int]:
    return repo_relative(source), int(slide)


def exact_library_path(repo: Path, value: str) -> Path:
    library = (repo / "模板").resolve()
    path = (repo / value).resolve()
    try:
        path.relative_to(library)
    except ValueError:
        fail(f"Refusing to modify a path outside 模板/: {path}")
    return path


def add_decision(
    decisions: dict[tuple[str, int], dict[str, Any]],
    source: str,
    slide: int,
    decision: dict[str, Any],
) -> None:
    key = page_key(source, slide)
    if key in decisions:
        fail(f"Page reviewed more than once: {key[0]} P{key[1]:02d}")
    decisions[key] = decision


def expand_manifest(manifest: dict[str, Any]) -> dict[tuple[str, int], dict[str, Any]]:
    if manifest.get("schema_version") != 1:
        fail("Curation manifest schema_version must be 1")
    sources = manifest.get("sources")
    if not isinstance(sources, list):
        fail("Curation manifest sources must be a list")

    decisions: dict[tuple[str, int], dict[str, Any]] = {}
    for source_spec in sources:
        if not isinstance(source_spec, dict) or not source_spec.get("source"):
            fail("Every curation source needs a source path")
        source = repo_relative(str(source_spec["source"]))
        keep_reason = str(source_spec.get("keep_reason", "保留为代表模板。"))
        for slide in source_spec.get("keep", []):
            add_decision(
                decisions,
                source,
                int(slide),
                {"decision": "keep", "reason": keep_reason},
            )
        for group in source_spec.get("merge", []):
            if not isinstance(group, dict) or not isinstance(group.get("into"), dict):
                fail(f"Invalid merge group in {source}")
            into = group["into"]
            representative = {
                "source": repo_relative(str(into["source"])),
                "slide": int(into["slide"]),
            }
            reason = str(group.get("reason", "同构页面合并到代表模板。"))
            for slide in group.get("slides", []):
                add_decision(
                    decisions,
                    source,
                    int(slide),
                    {
                        "decision": "merge",
                        "representative": representative,
                        "reason": reason,
                    },
                )
        for group in source_spec.get("discard", []):
            if not isinstance(group, dict):
                fail(f"Invalid discard group in {source}")
            reason = str(group.get("reason", "完成度或复用价值不足。"))
            for slide in group.get("slides", []):
                add_decision(
                    decisions,
                    source,
                    int(slide),
                    {"decision": "discard", "reason": reason},
                )
    return decisions


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Validate and apply keep/merge/discard decisions to learned PPT pages."
    )
    parser.add_argument("--repo-root", default=".")
    parser.add_argument("--manifest", default="模板/策展记录.json")
    parser.add_argument("--apply", action="store_true")
    args = parser.parse_args()

    repo = Path(args.repo_root).resolve()
    manifest_path = (repo / args.manifest).resolve()
    ledger_path = repo / "模板" / "学习记录.json"
    result_path = repo / "模板" / "策展结果.json"
    manifest = read_json(manifest_path)
    ledger = read_json(ledger_path)
    decisions = expand_manifest(manifest)

    ledger_pages: dict[tuple[str, int], dict[str, Any]] = {}
    for source_record in ledger.get("sources", []):
        source = repo_relative(str(source_record["source"]))
        for entry in source_record.get("slides", []):
            key = page_key(source, int(entry["slide"]))
            if key in ledger_pages:
                fail(f"Duplicate page in learning ledger: {key}")
            ledger_pages[key] = entry

    missing = sorted(set(ledger_pages) - set(decisions))
    unknown = sorted(set(decisions) - set(ledger_pages))
    if missing:
        fail("Pages without a curation decision: " + ", ".join(f"{s} P{p:02d}" for s, p in missing))
    if unknown:
        fail("Manifest pages absent from learning ledger: " + ", ".join(f"{s} P{p:02d}" for s, p in unknown))

    for key, decision in decisions.items():
        if decision["decision"] != "merge":
            continue
        representative = decision["representative"]
        rep_key = page_key(representative["source"], representative["slide"])
        if rep_key not in decisions:
            fail(f"Merge representative is unknown: {rep_key}")
        if decisions[rep_key]["decision"] != "keep":
            fail(f"Merge representative is not kept: {rep_key}")
        decision["representative"]["template"] = ledger_pages[rep_key]["template"]

    counts = {"keep": 0, "merge": 0, "discard": 0}
    for decision in decisions.values():
        counts[decision["decision"]] += 1

    result_entries: list[dict[str, Any]] = []
    if args.apply:
        for source_record in ledger.get("sources", []):
            source = repo_relative(str(source_record["source"]))
            for entry in source_record.get("slides", []):
                key = page_key(source, int(entry["slide"]))
                decision = decisions[key]
                entry["curation"] = decision
                if decision["decision"] != "keep":
                    for field in ("template", "constraints"):
                        path = exact_library_path(repo, str(entry[field]))
                        if path.exists():
                            path.unlink()
                result_entries.append(
                    {
                        "source": source,
                        "slide": int(entry["slide"]),
                        "template": entry["template"],
                        **decision,
                    }
                )

        validation = ledger.setdefault("validation", {})
        validation.pop("learned_templates", None)
        validation.update(
            {
                "curation_status": "passed",
                "source_pages_reviewed": len(decisions),
                "extracted_candidates": len(decisions),
                "curated_templates": counts["keep"],
                "merged_candidates": counts["merge"],
                "discarded_candidates": counts["discard"],
            }
        )
        write_json(ledger_path, ledger)
        write_json(
            result_path,
            {
                "schema_version": 1,
                "applied_at": datetime.now().astimezone().isoformat(timespec="seconds"),
                "manifest": repo_relative(str(manifest_path.relative_to(repo))),
                "summary": {"reviewed": len(decisions), **counts},
                "entries": result_entries,
            },
        )

    mode = "applied" if args.apply else "validated"
    print(
        json.dumps(
            {"status": mode, "reviewed": len(decisions), **counts},
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
