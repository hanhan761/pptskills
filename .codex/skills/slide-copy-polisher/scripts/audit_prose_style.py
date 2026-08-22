from __future__ import annotations

import argparse
import json
import re
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any


EM_DASH_RE = re.compile(r"[—―]+")
DEFENSIVE_PATTERNS = (
    ("防御性铺垫", re.compile(r"需要指出的是|值得注意的是|不难发现|毋庸置疑|不可否认")),
    ("空泛限定", re.compile(r"从某种意义上说|在一定程度上|总体而言|可以说|通常情况下|相对而言")),
    ("迂回连接", re.compile(r"通过.{0,24}可以|围绕.{0,24}展开|基于.{0,24}进行|从.{0,24}来看|就.{0,24}而言")),
    ("总结套话", re.compile(r"由此可见|在此基础上|进一步说明|综上所述|总的来说")),
)
SENTENCE_SPLIT_RE = re.compile(r"[。！？!?；;\n]+")
NOISE_RE = re.compile(r"[\s\u3000，。！？!?；;：:、,.、（）()【】\[\]“”‘’'\"—―\-_/]+")


def text_from_value(value: Any) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, list):
        return "".join(text_from_value(item) for item in value)
    if isinstance(value, dict):
        if "text" in value:
            return text_from_value(value["text"])
        if "lines" in value:
            return text_from_value(value["lines"])
        if "rich_lines" in value:
            return text_from_value(value["rich_lines"])
        if "value" in value:
            return text_from_value(value["value"])
    return ""


def extract_records(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        raw = payload
    elif isinstance(payload, dict):
        raw = None
        for key in ("texts", "entries", "replacements", "records"):
            if isinstance(payload.get(key), list):
                raw = payload[key]
                break
        if raw is None and isinstance(payload.get("slides"), list):
            raw = payload["slides"]
        if raw is None:
            raw = [payload]
    else:
        return []

    records: list[dict[str, Any]] = []
    for index, item in enumerate(raw, 1):
        if not isinstance(item, dict):
            continue
        value = text_from_value(item)
        if not value.strip():
            continue
        record = dict(item)
        record["_index"] = index
        record["_text"] = value
        records.append(record)
    return records


def clean_for_similarity(value: str) -> str:
    return NOISE_RE.sub("", value).lower()


def snippet(value: str, limit: int = 70) -> str:
    compact = re.sub(r"\s+", " ", value).strip()
    return compact if len(compact) <= limit else compact[: limit - 1] + "…"


def add_issue(
    issues: list[dict[str, Any]],
    kind: str,
    record: dict[str, Any] | None,
    detail: str,
    text: str = "",
) -> None:
    item: dict[str, Any] = {"kind": kind, "detail": detail}
    if record is not None:
        item["record"] = record["_index"]
        for key in ("slide", "shape", "role"):
            if key in record:
                item[key] = record[key]
    if text:
        item["text"] = snippet(text)
    issues.append(item)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Audit slide copy for indirect, defensive, repetitive, complex, or em-dash-heavy wording."
    )
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--max-sentence-chars", type=int, default=45)
    parser.add_argument("--similarity-threshold", type=float, default=0.82)
    args = parser.parse_args()

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    records = extract_records(payload)
    issues: list[dict[str, Any]] = []

    for record in records:
        text = record["_text"]
        em_matches = EM_DASH_RE.findall(text)
        if em_matches:
            add_issue(
                issues,
                "破折号",
                record,
                "remove em dashes and split the sentence with ordinary punctuation",
                text,
            )
        for label, pattern in DEFENSIVE_PATTERNS:
            match = pattern.search(text)
            if match:
                add_issue(
                    issues,
                    label,
                    record,
                    f"review the phrase {match.group(0)!r}; state the judgment or evidence directly",
                    text,
                )
        sentences = [part.strip() for part in SENTENCE_SPLIT_RE.split(text) if part.strip()]
        for sentence in sentences:
            if len(re.sub(r"\s", "", sentence)) > args.max_sentence_chars:
                add_issue(
                    issues,
                    "复杂长句",
                    record,
                    f"sentence exceeds {args.max_sentence_chars} visible characters; split by action or evidence",
                    sentence,
                )
            if sentence.count("，") >= 3 and sentence.count("、") >= 2:
                add_issue(
                    issues,
                    "多层并列",
                    record,
                    "too many parallel clauses; keep one main action per sentence",
                    sentence,
                )

    signatures: dict[str, list[dict[str, Any]]] = {}
    for record in records:
        signature = clean_for_similarity(record["_text"])
        if len(signature) >= 8:
            signatures.setdefault(signature, []).append(record)
    for signature, matches in signatures.items():
        if len(matches) < 2:
            continue
        for record in matches[1:]:
            add_issue(
                issues,
                "精确重复",
                record,
                f"same normalized copy already appears in record {matches[0]['_index']}",
                record["_text"],
            )

    for left_index, left in enumerate(records):
        left_signature = clean_for_similarity(left["_text"])
        if len(left_signature) < 12:
            continue
        for right in records[left_index + 1 :]:
            right_signature = clean_for_similarity(right["_text"])
            if len(right_signature) < 12:
                continue
            ratio = SequenceMatcher(None, left_signature, right_signature).ratio()
            if ratio >= args.similarity_threshold and left_signature != right_signature:
                add_issue(
                    issues,
                    "高相似重复",
                    right,
                    f"record {right['_index']} overlaps record {left['_index']} at {ratio:.2f}; give the two slots different jobs",
                    right["_text"],
                )

    counts: dict[str, int] = {}
    for issue in issues:
        counts[issue["kind"]] = counts.get(issue["kind"], 0) + 1
    result = {
        "records": len(records),
        "max_sentence_chars": args.max_sentence_chars,
        "similarity_threshold": args.similarity_threshold,
        "issue_counts": counts,
        "issues": issues,
        "passed": not issues,
    }
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    if issues:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
