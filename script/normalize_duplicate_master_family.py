#!/usr/bin/env python3
"""Collapse duplicated-but-visually-equivalent PPT masters onto one family.

This is a mechanical package rewrite for a reviewed source copy. It does not
redraw slides or alter slide shapes. The caller supplies exact slide numbers
and the target layout relationship already present in the package.
"""

from __future__ import annotations

import argparse
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


def slide_layout_target(archive: zipfile.ZipFile, slide: int) -> str:
    name = f"ppt/slides/_rels/slide{slide}.xml.rels"
    root = ET.fromstring(archive.read(name))
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        if rel.get("Type", "").endswith("/slideLayout"):
            return posixpath.normpath(posixpath.join("ppt/slides", rel.get("Target", "")))
    raise RuntimeError(f"Slide {slide} has no slideLayout relationship")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--slides", type=int, nargs="+", required=True)
    parser.add_argument("--target-layout", type=int, required=True)
    args = parser.parse_args()

    source = Path(args.input).resolve()
    output = Path(args.output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_suffix(output.suffix + ".tmp")
    target = f"../slideLayouts/slideLayout{args.target_layout}.xml"

    with zipfile.ZipFile(source) as archive:
        required = set(args.slides)
        for slide in required:
            current = slide_layout_target(archive, slide)
            if not re.fullmatch(r"ppt/slideLayouts/slideLayout(?:5|9|13)\.xml", current):
                raise RuntimeError(
                    f"Slide {slide} is not attached to an expected duplicate layout: {current}"
                )

        obsolete = set()
        for master in (2, 3, 4):
            obsolete.add(f"ppt/slideMasters/slideMaster{master}.xml")
            obsolete.add(f"ppt/slideMasters/_rels/slideMaster{master}.xml.rels")
            obsolete.add(f"ppt/theme/theme{master}.xml")
        for layout in range(5, 17):
            obsolete.add(f"ppt/slideLayouts/slideLayout{layout}.xml")
            obsolete.add(f"ppt/slideLayouts/_rels/slideLayout{layout}.xml.rels")

        with zipfile.ZipFile(
            temporary, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9
        ) as rewritten:
            for info in archive.infolist():
                name = info.filename
                if name in obsolete:
                    continue
                data = archive.read(name)

                match = re.fullmatch(r"ppt/slides/_rels/slide(\d+)\.xml\.rels", name)
                if match and int(match.group(1)) in required:
                    root = ET.fromstring(data)
                    changed = False
                    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
                        if rel.get("Type", "").endswith("/slideLayout"):
                            rel.set("Target", target)
                            changed = True
                    if not changed:
                        raise RuntimeError(f"No layout relationship changed for {name}")
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

                elif name == "ppt/presentation.xml":
                    root = ET.fromstring(data)
                    master_list = root.find(f"{{{P_NS}}}sldMasterIdLst")
                    if master_list is None:
                        raise RuntimeError("presentation.xml has no sldMasterIdLst")
                    children = list(master_list)
                    for child in children[1:]:
                        master_list.remove(child)
                    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)

                elif name == "ppt/_rels/presentation.xml.rels":
                    root = ET.fromstring(data)
                    for rel in list(root.findall(f"{{{REL_NS}}}Relationship")):
                        if rel.get("Type", "").endswith(("/slideMaster", "/theme")):
                            target_value = rel.get("Target", "")
                            if re.search(r"(?:slideMaster|theme)[234]\.xml$", target_value):
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
    print(output)


if __name__ == "__main__":
    main()
