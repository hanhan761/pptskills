from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
import zipfile

from template_ppt import (
    NS,
    find_repo_root,
    named_shape_container,
    pptx_patch,
    read_json,
    resolve_repo_path,
    slide_xml_name,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Attach verified caption/source metadata to existing PPT picture shapes"
    )
    parser.add_argument("--input", required=True)
    parser.add_argument("--sources", required=True)
    parser.add_argument("--output")
    parser.add_argument("--repo-root")
    args = parser.parse_args()

    repo = find_repo_root(Path(args.repo_root) if args.repo_root else None)
    input_path = resolve_repo_path(args.input, repo)
    output_path = resolve_repo_path(args.output, repo) if args.output else input_path
    payload = read_json(resolve_repo_path(args.sources, repo))
    records = payload.get("images") if isinstance(payload, dict) else None
    if not isinstance(records, list) or not records:
        raise SystemExit("Sources JSON must contain a non-empty 'images' list")

    def patcher(
        source_zip: zipfile.ZipFile, names: set[str]
    ) -> tuple[dict[str, bytes], dict[str, bytes]]:
        replacements: dict[str, bytes] = {}
        by_slide: dict[int, list[dict[str, Any]]] = {}
        for record in records:
            by_slide.setdefault(int(record["slide"]), []).append(record)
        for slide_number, slide_records in by_slide.items():
            xml_name = slide_xml_name(slide_number)
            if xml_name not in names:
                raise SystemExit(f"Slide {slide_number} does not exist")
            root = ET.fromstring(source_zip.read(xml_name))
            for record in slide_records:
                shape_name = str(record["shape"])
                occurrence = int(record.get("occurrence", 1))
                picture = named_shape_container(root, shape_name, occurrence)
                if picture.tag != f"{{{NS['p']}}}pic":
                    raise SystemExit(
                        f"Slide {slide_number} shape {shape_name!r} occurrence "
                        f"{occurrence} is not a picture"
                    )
                properties = picture.find("./p:nvPicPr/p:cNvPr", NS)
                if properties is None:
                    raise SystemExit(
                        f"Slide {slide_number} shape {shape_name!r} has no cNvPr"
                    )
                caption = str(record["caption"]).strip()
                credit = str(record["credit"]).strip()
                source_page = str(record["source_page"]).strip()
                properties.set("title", caption)
                properties.set(
                    "descr",
                    f"{caption}；来源：{credit}；原始页面：{source_page}",
                )
            replacements[xml_name] = ET.tostring(
                root, encoding="utf-8", xml_declaration=True
            )
        return replacements, {}

    pptx_patch(input_path, output_path, patcher)
    print(output_path)


if __name__ == "__main__":
    main()
