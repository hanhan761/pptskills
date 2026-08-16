from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
import zipfile

from template_ppt import (
    NS,
    find_repo_root,
    named_shape_container,
    read_json,
    resolve_repo_path,
    slide_rels_name,
    slide_xml_name,
)


def digest(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Verify embedded picture bytes, fallback cleanup, caption and source metadata"
    )
    parser.add_argument("--input", required=True)
    parser.add_argument("--sources", required=True)
    parser.add_argument("--repo-root")
    args = parser.parse_args()

    repo = find_repo_root(Path(args.repo_root) if args.repo_root else None)
    deck = resolve_repo_path(args.input, repo)
    payload = read_json(resolve_repo_path(args.sources, repo))
    records = payload.get("images") if isinstance(payload, dict) else None
    if not isinstance(records, list) or not records:
        raise SystemExit("Sources JSON must contain a non-empty 'images' list")

    errors: list[str] = []
    checked = 0
    with zipfile.ZipFile(deck, "r") as archive:
        names = set(archive.namelist())
        slide_cache: dict[int, ET.Element] = {}
        rel_cache: dict[int, dict[str, str]] = {}
        for record in records:
            slide_number = int(record["slide"])
            shape_name = str(record["shape"])
            occurrence = int(record.get("occurrence", 1))
            context = f"slide {slide_number} {shape_name!r} occurrence {occurrence}"
            try:
                if slide_number not in slide_cache:
                    xml_name = slide_xml_name(slide_number)
                    rels_name = slide_rels_name(slide_number)
                    slide_cache[slide_number] = ET.fromstring(archive.read(xml_name))
                    rels = ET.fromstring(archive.read(rels_name))
                    rel_cache[slide_number] = {
                        str(node.attrib["Id"]): str(node.attrib["Target"])
                        for node in rels.findall("rel:Relationship", NS)
                    }
                picture = named_shape_container(
                    slide_cache[slide_number], shape_name, occurrence
                )
                if picture.tag != f"{{{NS['p']}}}pic":
                    raise ValueError("target is not a picture")
                properties = picture.find("./p:nvPicPr/p:cNvPr", NS)
                blip = picture.find("./p:blipFill/a:blip", NS)
                if properties is None or blip is None:
                    raise ValueError("picture properties or embedded blip missing")
                expected_title = str(record["caption"]).strip()
                expected_descr = (
                    f"{expected_title}；来源：{str(record['credit']).strip()}；"
                    f"原始页面：{str(record['source_page']).strip()}"
                )
                if properties.attrib.get("title") != expected_title:
                    errors.append(f"{context}: caption title mismatch")
                if properties.attrib.get("descr") != expected_descr:
                    errors.append(f"{context}: source description mismatch")
                if f"{{{NS['r']}}}link" in blip.attrib:
                    errors.append(f"{context}: stale external image link remains")
                if blip.find("./a:extLst", NS) is not None:
                    errors.append(f"{context}: stale SVG/fallback extension remains")
                relationship_id = blip.attrib.get(f"{{{NS['r']}}}embed")
                target = rel_cache[slide_number].get(str(relationship_id))
                if not target:
                    raise ValueError("embedded image relationship missing")
                media_name = posixpath.normpath(
                    posixpath.join("ppt/slides", target)
                )
                if media_name not in names:
                    raise ValueError(f"embedded media missing: {media_name}")
                source_file = resolve_repo_path(record["file"], repo)
                if digest(archive.read(media_name)) != digest(source_file.read_bytes()):
                    errors.append(f"{context}: embedded bytes do not match source asset")
                checked += 1
            except Exception as exc:
                errors.append(f"{context}: {exc}")

    result = {
        "ok": not errors,
        "deck": str(deck),
        "records_checked": checked,
        "errors": errors,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
