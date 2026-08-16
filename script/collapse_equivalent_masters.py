#!/usr/bin/env python3
"""Collapse PowerPoint-created duplicate masters/layouts onto the first family.

PowerPoint can duplicate an otherwise identical custom layout when same-deck
slides are duplicated through COM.  This tool is intentionally conservative:
it proceeds only when every slide master and used slide layout is structurally
equivalent to the first one after relationship IDs are ignored.  It then only
rewires slide-layout relationships and removes now-unused duplicate package
parts; slide XML and slide shapes are never changed.
"""

from __future__ import annotations

import argparse
import copy
import os
import posixpath
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def rels_name(part: str) -> str:
    directory, filename = posixpath.split(part)
    return posixpath.join(directory, "_rels", filename + ".rels")


def normalize_xml(data: bytes) -> bytes:
    root = ET.fromstring(data)
    relationship_attr = f"{{{R_NS}}}id"
    for node in root.iter():
        node.attrib.pop(relationship_attr, None)
        # PowerPoint allocates fresh numeric layout IDs when it duplicates a
        # master.  They carry no visual semantics; relationships identify the
        # actual layouts and are compared separately below.
        if node.tag == f"{{{P_NS}}}sldLayoutId":
            node.attrib.pop("id", None)
    return ET.tostring(root, encoding="utf-8")


def relationships(archive: zipfile.ZipFile, part: str) -> list[tuple[str, str]]:
    name = rels_name(part)
    if name not in archive.namelist():
        return []
    root = ET.fromstring(archive.read(name))
    base = posixpath.dirname(part)
    values = []
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        target = posixpath.normpath(posixpath.join(base, rel.get("Target", "")))
        values.append((rel.get("Type", ""), target))
    return values


def slide_layout(archive: zipfile.ZipFile, slide_part: str) -> str:
    matches = [target for kind, target in relationships(archive, slide_part) if kind.endswith("/slideLayout")]
    if len(matches) != 1:
        raise RuntimeError(f"{slide_part} has {len(matches)} layout relationships")
    return matches[0]


def layout_master(archive: zipfile.ZipFile, layout_part: str) -> str:
    matches = [target for kind, target in relationships(archive, layout_part) if kind.endswith("/slideMaster")]
    if len(matches) != 1:
        raise RuntimeError(f"{layout_part} has {len(matches)} master relationships")
    return matches[0]


def comparable_master(archive: zipfile.ZipFile, master_part: str) -> tuple[bytes, tuple[str, ...]]:
    layouts = tuple(sorted(
        normalize_xml(archive.read(target)).decode("utf-8")
        for kind, target in relationships(archive, master_part)
        if kind.endswith("/slideLayout")
    ))
    return normalize_xml(archive.read(master_part)), layouts


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    source = Path(args.input).resolve()
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(output.suffix + ".tmp")

    with zipfile.ZipFile(source, "r") as archive:
        presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
        pres_rels = relationships(archive, "ppt/presentation.xml")
        master_parts = [target for kind, target in pres_rels if kind.endswith("/slideMaster")]
        if len(master_parts) <= 1:
            if source != output:
                output.write_bytes(source.read_bytes())
            print(f"No duplicate masters: {output}")
            return

        # presentation.xml.rels ordering is not the semantic family order.
        # Resolve the first sldMasterId relationship from presentation.xml;
        # this is the source family's retained master in our compose workflow.
        rel_root = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
        rel_targets = {
            rel.get("Id", ""): posixpath.normpath(posixpath.join("ppt", rel.get("Target", "")))
            for rel in rel_root.findall(f"{{{REL_NS}}}Relationship")
            if rel.get("Type", "").endswith("/slideMaster")
        }
        master_list = presentation.find(f"{{{P_NS}}}sldMasterIdLst")
        if master_list is None or not list(master_list):
            raise RuntimeError("presentation.xml has no master list")
        target_rel_id = list(master_list)[0].get(f"{{{R_NS}}}id", "")
        target_master = rel_targets.get(target_rel_id, "")
        if target_master not in master_parts:
            raise RuntimeError("Cannot resolve the first presentation master relationship")
        master_parts = [target_master, *[part for part in master_parts if part != target_master]]
        baseline = comparable_master(archive, target_master)
        for part in master_parts[1:]:
            if comparable_master(archive, part) != baseline:
                raise RuntimeError(f"Master is not structurally equivalent to first master: {part}")

        target_layouts = [target for kind, target in relationships(archive, target_master) if kind.endswith("/slideLayout")]
        if not target_layouts:
            raise RuntimeError("Target master has no layouts")
        layout_signature_to_target = {normalize_xml(archive.read(part)): part for part in target_layouts}

        slide_parts = sorted(
            (name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
            key=lambda value: int(re.search(r"(\d+)", posixpath.basename(value)).group(1)),
        )
        slide_layout_targets: dict[str, str] = {}
        for slide in slide_parts:
            current = slide_layout(archive, slide)
            signature = normalize_xml(archive.read(current))
            target = layout_signature_to_target.get(signature)
            if target is None:
                raise RuntimeError(f"Used layout is not equivalent to a target-family layout: {current}")
            slide_layout_targets[slide] = target

        obsolete_masters = set(master_parts[1:])
        obsolete_layouts: set[str] = set()
        obsolete_themes: set[str] = set()
        for master in obsolete_masters:
            obsolete_layouts.update(target for kind, target in relationships(archive, master) if kind.endswith("/slideLayout"))
            obsolete_themes.update(target for kind, target in relationships(archive, master) if kind.endswith("/theme"))
        obsolete = set(obsolete_masters) | obsolete_layouts | obsolete_themes
        obsolete |= {rels_name(part) for part in list(obsolete_masters | obsolete_layouts)}

        with zipfile.ZipFile(temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as rewritten:
            for info in archive.infolist():
                name = info.filename
                if name in obsolete:
                    continue
                data = archive.read(name)

                match = re.fullmatch(r"ppt/slides/_rels/(slide\d+\.xml)\.rels", name)
                if match:
                    slide = "ppt/slides/" + match.group(1)
                    root = ET.fromstring(data)
                    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
                        if rel.get("Type", "").endswith("/slideLayout"):
                            rel.set("Target", posixpath.relpath(slide_layout_targets[slide], "ppt/slides"))
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif name == "ppt/presentation.xml":
                    root = ET.fromstring(data)
                    master_list = root.find(f"{{{P_NS}}}sldMasterIdLst")
                    if master_list is None:
                        raise RuntimeError("presentation.xml has no master list")
                    for child in list(master_list):
                        if child.get(f"{{{R_NS}}}id", "") != target_rel_id:
                            master_list.remove(child)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif name == "ppt/_rels/presentation.xml.rels":
                    root = ET.fromstring(data)
                    base = "ppt"
                    for rel in list(root.findall(f"{{{REL_NS}}}Relationship")):
                        target = posixpath.normpath(posixpath.join(base, rel.get("Target", "")))
                        if target in obsolete_masters or target in obsolete_themes:
                            root.remove(rel)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                elif name == "[Content_Types].xml":
                    root = ET.fromstring(data)
                    for override in list(root.findall(f"{{{CT_NS}}}Override")):
                        part = override.get("PartName", "").lstrip("/")
                        if part in obsolete:
                            root.remove(override)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                rewritten.writestr(info, data)

    os.replace(temporary, output)
    print(f"Collapsed {len(master_parts)} equivalent masters to one: {output}")


if __name__ == "__main__":
    main()
