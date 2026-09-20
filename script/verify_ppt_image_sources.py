from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
from pathlib import Path
from xml.etree import ElementTree as ET
import zipfile

from pptx import Presentation

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


def source_identity(record: dict) -> str:
    """Use a source identity that does not depend on the caption text."""

    for key in ("source_id", "source_file", "direct_url", "source_page"):
        value = str(record.get(key, "") or "").strip()
        if value:
            return f"{key}:{value}"
    return ""


def slide_parts_in_presentation_order(deck: Path) -> list[str]:
    """Resolve logical slide positions to package parts.

    PowerPoint preserves part names when slides are deleted or reordered, so
    ``slide 18`` is not always ``ppt/slides/slide18.xml``. Source manifests use
    presentation order, which is also what users mean by a page number.
    """

    presentation = Presentation(str(deck))
    return [str(slide.part.partname).lstrip("/") for slide in presentation.slides]


def relationship_part_name(slide_part: str) -> str:
    parent, filename = posixpath.split(slide_part)
    return posixpath.join(parent, "_rels", f"{filename}.rels")


def image_properties_and_blip(container: ET.Element) -> tuple[ET.Element, ET.Element]:
    """Return metadata and embedded media for native pictures or filled shapes.

    PowerPoint can implement an existing photo frame as either ``p:pic`` or a
    normal editable ``p:sp`` whose fill is an ``a:blipFill``. Both are valid
    image slots. Treating the latter as non-images caused source verification
    to reject a legitimate in-place ``Shape.Fill.UserPicture`` replacement.
    """

    if container.tag == f"{{{NS['p']}}}pic":
        properties = container.find("./p:nvPicPr/p:cNvPr", NS)
        blip = container.find("./p:blipFill/a:blip", NS)
    elif container.tag == f"{{{NS['p']}}}sp":
        properties = container.find("./p:nvSpPr/p:cNvPr", NS)
        blips = container.findall("./p:spPr/a:blipFill/a:blip", NS)
        if len(blips) != 1:
            raise ValueError(
                "target is a filled shape but does not contain exactly one embedded picture fill"
            )
        blip = blips[0]
    else:
        raise ValueError("target is neither a native picture nor a picture-filled shape")
    if properties is None or blip is None:
        raise ValueError("picture properties or embedded blip missing")
    return properties, blip


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
    embedded_sha_seen: dict[str, str] = {}
    source_identity_seen: dict[str, str] = {}
    embedded_sha_reuse: dict[str, list[str]] = {}
    source_identity_reuse: dict[str, list[str]] = {}
    slide_parts = slide_parts_in_presentation_order(deck)
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
                if slide_number < 1 or slide_number > len(slide_parts):
                    raise ValueError(
                        f"slide number is outside 1..{len(slide_parts)}"
                    )
                if str(record.get("shared_visual_slot", "") or "").strip():
                    raise ValueError("shared_visual_slot is forbidden for content images")
                identity = source_identity(record)
                if not identity:
                    raise ValueError(
                        "missing stable source identity; add source_id, source_file, "
                        "direct_url, or source_page"
                    )
                previous_identity = source_identity_seen.get(identity)
                if previous_identity:
                    source_identity_reuse.setdefault(identity, [previous_identity]).append(
                        context
                    )
                    errors.append(
                        f"{context}: source identity is already used by {previous_identity}"
                    )
                else:
                    source_identity_seen[identity] = context
                if slide_number not in slide_cache:
                    xml_name = slide_parts[slide_number - 1]
                    rels_name = relationship_part_name(xml_name)
                    slide_cache[slide_number] = ET.fromstring(archive.read(xml_name))
                    rels = ET.fromstring(archive.read(rels_name))
                    rel_cache[slide_number] = {
                        str(node.attrib["Id"]): str(node.attrib["Target"])
                        for node in rels.findall("rel:Relationship", NS)
                    }
                picture = named_shape_container(
                    slide_cache[slide_number], shape_name, occurrence
                )
                properties, blip = image_properties_and_blip(picture)
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
                embedded_bytes = archive.read(media_name)
                embedded_sha = digest(embedded_bytes)
                previous_media = embedded_sha_seen.get(embedded_sha)
                if previous_media:
                    embedded_sha_reuse.setdefault(embedded_sha, [previous_media]).append(
                        context
                    )
                    errors.append(
                        f"{context}: embedded image bytes are reused from {previous_media}"
                    )
                else:
                    embedded_sha_seen[embedded_sha] = context
                source_file = resolve_repo_path(record["file"], repo)
                if embedded_sha != digest(source_file.read_bytes()):
                    errors.append(f"{context}: embedded bytes do not match source asset")
                checked += 1
            except Exception as exc:
                errors.append(f"{context}: {exc}")

    result = {
        "ok": not errors,
        "deck": str(deck),
        "records_checked": checked,
        "embedded_sha_reuse": embedded_sha_reuse,
        "source_identity_reuse": source_identity_reuse,
        "errors": errors,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if errors:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
