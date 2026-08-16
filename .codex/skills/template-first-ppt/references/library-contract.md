# Master-family and content-module contract

## Storage

```text
待提取/
模板/
  母版/
    索引.json
    <家族>/
      母版源.pptx
      family.json
  模块/
    内容/
      <模块>.module.json
  封面/        # legacy candidate sources
  目录/        # legacy candidate sources
  内容/        # legacy candidate sources
  学习记录.json
  策展记录.json
模板研发部门/
  create1/                 # exactly one current review PPTX; one page by default or explicit N-page batch
  _internal/               # briefs, sources, renders, review records
  已审核/                  # frozen approved source pages, not automatically active
已提取/
task/
script/
test/
```

`test/` or another user-facing delivery directory contains exactly one final PPTX. Plans, modules, contracts, provenance, images, source manifests, renders, and code stay in the internal workspace.

The master source is a byte-faithful copy of a reviewed deck. It can contain multiple source slides because it is the family seed, not a one-page content template. Legacy template PPTX files remain candidate sources only and must not be directly mixed into a new deck.

## Template R&D lifecycle

Template R&D is an isolated exception for creating new below-title content-module candidates when the user explicitly asks for them. Default to one page. Create a multi-page batch only when the user explicitly requests a number, and keep every page independently reviewable. It does not loosen production rules.

Use these states:

- `draft`: one candidate is being prepared for first review;
- `revise`: the user requested changes to that same candidate;
- `approved`: the user explicitly accepted the candidate for formalization;
- `rejected`: the candidate must not enter the library.

For every R&D request:

1. Select a target family and duplicate its approved content-shell page. Keep the master identity and fixed shell unchanged.
2. Create one editable module below the title boundary. Prefer stable replacement slots, meaningful real-image frames, caption/source positions, and a structure that can become separately reviewed sibling variants. The review slide must be filled as a realistic finished example; visible template instructions such as “适用范围” and “迁移方式” belong in internal metadata, not on the slide.
3. Apply a large-and-full visual gate before review: the core content should occupy the safe region decisively, small cards should be consolidated, and body text must remain projection-readable. First measure a user-calibrated page and at least two reviewed or active templates with comparable geometry. Use their relative title/body/conclusion hierarchy as the authority, not a single global font setting. For `public-blue-white`, the current empirical reference is 24–26 pt title bands, 18–28 pt section labels, 16–20 pt wide body text, about 11.5–15 pt narrow-card body text when a reviewed source supports it, 18–20 pt conclusions, and 9.5–11 pt captions. Geometry and rendered line behavior decide the final value.
4. Keep exactly one current review PPTX in `模板研发部门/create1/`. Use one page by default; for an explicit N-template request, use exactly N independently identified candidate pages in that one file. Overwrite it during review instead of accumulating process versions. Store non-PPT evidence and per-page states in `模板研发部门/_internal/`.
5. Keep the candidate out of build plans and active indexes until explicit user approval.
6. On approval, freeze every approved candidate as an immutable one-page source in `模板研发部门/已审核/`, create one normal module contract per page in `模板/模块/内容/`, and record the source deck/slide/hash, target master fingerprint, exact shape references, z-order, geometry signature, text capacities, fixed title lines, mixed-run roles, picture slots, captions, sources, and compatible families. Unapproved pages in the same batch remain isolated.
7. Test-mount and run `verify-module`, `verify-layered`, `verify-master`, plus rendered visual review. Only a fully passing candidate can be marked active in `模板/模块/索引.json`.

“Easy to extend” means a reusable communication job and clear sibling-variant direction. It never permits dynamic shape insertion, movement, scaling, or text reflow in production. Each two-image, three-image, four-image, or other geometry variant is a separate candidate with its own approval and contract.

User approval decides whether a candidate is worth preserving; contracts and validation decide whether it is safe to reuse. A file in `已审核/` is not automatically an active module.

## Master-family contract

`模板/母版/<家族>/family.json` records the one visual identity shared by a deck:

```json
{
  "schema_version": 1,
  "family_id": "public-blue-white",
  "display_name": "通用蓝白",
  "default": true,
  "source_asset": "模板/母版/通用蓝白/母版源.pptx",
  "source_sha256": "...",
  "slide_size_emu": {"width": 12192000, "height": 6858000},
  "powerpoint_master": {
    "part": "/ppt/slideMasters/slideMaster1.xml",
    "master_count": 1,
    "all_source_slides_use_this_master": true
  },
  "approved_layouts": [
    {"name": "标题和内容", "roles": ["cover", "agenda", "section"]},
    {"name": "4_标题幻灯片", "roles": ["content_shell"]}
  ],
  "role_pages": {
    "cover": {"source_slide": 1, "layout": "标题和内容"},
    "content_shell": {"source_slide": 3, "layout": "4_标题幻灯片"}
  },
  "shell_contract": {
    "content_top_emu": 650240,
    "fixed_layout_shapes": ["直接连接符 59", "矩形: 单圆角 205"],
    "fixed_slide_shapes": ["矩形 1", "文本框 2"],
    "content_module_policy": {
      "preserve_original_absolute_geometry": true,
      "preserve_z_order": true,
      "allow_scale": false,
      "allow_relayout": false,
      "reject_if_outside_safe_region": true
    }
  }
}
```

The `part` path is family evidence, not a portable identifier to compare across separately packaged PPTX files. Final validation checks that every output slide resolves to the same master part inside the output package and that its layout is approved by name.

If the source lacks a closing page, use:

```json
"closing": {
  "native_source_slide": null,
  "strategy": "cover-derived"
}
```

This permits only duplication of the same-family cover page followed by normal text/image replacement.

## Content-module contract

A module is the reusable shape set below a family's fixed shell. Store metadata in `<模块>.module.json`. The contract may point to a source PPTX instead of duplicating media; the source must remain available and hash-locked.

```json
{
  "schema_version": 1,
  "module_id": "five-stage-process-double-summary",
  "display_name": "五节点流程与双总结",
  "communication_job": "流程拆解",
  "source": {
    "deck": "已提取/example.pptx",
    "sha256": "...",
    "slide": 12,
    "master_family": "source-family-id"
  },
  "shape_refs": [
    {"name": "矩形 8", "occurrence": 1},
    {"name": "文本框 21", "occurrence": 1}
  ],
  "bounds_emu": {
    "left": 304165,
    "top": 743585,
    "right": 11698605,
    "bottom": 6500000
  },
  "geometry_sha256": "...",
  "compatible_master_families": ["public-blue-white"],
  "constraints": {
    "image_only": false,
    "text_slots": []
  }
}
```

Every listed top-level shape keeps its source z-order and absolute geometry. All visible bounds must be at or below the target family's `content_top_emu` and within the slide edges. Compatibility must be proven by a test mount and full render; matching page size alone is insufficient.

## Text-capacity slots

Count visible characters after removing whitespace; punctuation counts. Use `occurrence` for duplicate names. Mark every required slot. Determine fixed line counts from rendered visual intent or explicit user correction, not from XML paragraphs or automatic wrapping.

Use `style_roles` to point to representative 1-based source runs for mixed-style text. Preserve the template's font, size, weight, and color.

```json
{
  "shape": "标题 1",
  "occurrence": 1,
  "required": true,
  "min_chars": 16,
  "max_chars": 28,
  "line_count": 2,
  "line_min_chars": [8, 8],
  "line_max_chars": [14, 14],
  "style_roles": {"base": 1, "accent": 2},
  "required_style_roles": ["base", "accent"]
}
```

A genuinely text-free module uses `"image_only": true` and an empty `text_slots` list.

## Layered build plan

Every slide declares the same family. Role pages have a same-family source role.正文 pages add a module.

```json
{
  "master_family": "public-blue-white",
  "slides": [
    {
      "purpose": "封面",
      "shell": {"role": "cover", "source_slide": 1},
      "content": {"type": "family_role", "role": "cover"}
    },
    {
      "purpose": "研究路径",
      "shell": {"role": "content_shell", "source_slide": 3},
      "content": {
        "type": "module",
        "module": "模板/模块/内容/五节点流程与双总结.module.json"
      }
    }
  ]
}
```

Build by copying the selected `母版源.pptx`, duplicating role/shell slides inside that presentation, mounting modules, and deleting unused source pages. Never create a blank presentation and never insert a whole slide from another family.

Build both same-source and compatible cross-source modules mechanically with:

```powershell
python script/template_ppt.py compose-family --plan <plan.json> --output <working.pptx>
```

For a same-source module, the command duplicates the registered family source slide by stable SlideID. For a cross-source module, it duplicates the target family's content-shell slide, deletes only its variable正文 shapes, and copies only the external module's registered shapes in source z-order. The external slide, layout, and master are never inserted. Cross-source contracts must include the source master fingerprint, and `verify-layered` must compare target-shell geometry and module geometry independently before `verify-master` proves that only the target master remains.

## Layered provenance

Keep `<working-output>.provenance.json` beside the internal deck until verification:

```json
{
  "schema_version": 2,
  "master_family": "public-blue-white",
  "master_source": "模板/母版/通用蓝白/母版源.pptx",
  "master_source_sha256": "...",
  "slides": [
    {
      "output_slide": 2,
      "purpose": "研究路径",
      "shell": {"role": "content_shell", "source_slide": 3},
      "content_module": {
        "contract": "模板/模块/内容/五节点流程与双总结.module.json",
        "source_slide": 12,
        "geometry_sha256": "..."
      }
    }
  ]
}
```

Provenance is internal evidence, not a deliverable.

## Replacement maps

Use exact shape names from `inspect`. A fixed-line title uses explicit `lines`; a mixed-style title uses `rich_lines` and only roles defined by its capacity contract.

```json
{
  "replacements": [
    {"slide": 2, "shape": "副标题 2", "lines": ["第一行标题", "第二行标题"]},
    {
      "slide": 3,
      "shape": "标题 1",
      "rich_lines": [[
        {"text": "面向商业航天", "style": "base"},
        {"text": "低成本", "style": "accent"}
      ]]
    }
  ]
}
```

Picture replacement uses only existing picture shapes and keeps the frame fixed:

```json
{
  "replacements": [{
    "slide": 2,
    "shape": "图片 5",
    "image": "task/task1/images/factory.png",
    "fit": "cover",
    "focal_x": 0.6,
    "focal_y": 0.45,
    "caption": "RAMFIRE nozzle hot-fire test",
    "credit": "NASA, 2023",
    "source_page": "https://www.nasa.gov/..."
  }]
}
```

When caption fields are present, `replace-image` must also update the picture's native `cNvPr` `title` and `descr`; changing the displayed media while leaving old alternative text and source metadata is a failed replacement.

When a mounted cross-source module contains a shape whose name collides with a fixed target-shell shape, keep `occurrence` aligned with the source module contract and use `output_occurrence` to select the actual same-named shape in the assembled slide. Never rename or replace the fixed shell shape to resolve the collision.

When a source picture contains an SVG alternate image, linked image, or a malformed media extension content type, replacement must also remove the stale alternate/link reference and normalize the media ContentType. The rendered slide—not merely the changed `r:embed` relationship—is the authority for whether the replacement succeeded.

For 100+ slide decks, composition must be resumable and committed in short batches. Duplicate family shells and mount registered module shapes into a staging copy, save successfully, then promote that batch to the working deck while recording checkpoint state. Restart PowerPoint between bounded batches. If one module repeatedly crashes at the same mount point, mark it inactive or restricted and select another verified module that performs the same communication job; never repair the failure by resizing, decomposing, or redrawing the module.

## Image-source manifest

Use real images by default. Record one verified source entry per replacement:

```json
{
  "images": [{
    "slide": 2,
    "shape": "图片 5",
    "file": "task/task1/images/factory.png",
    "caption": "RAMFIRE nozzle hot-fire test / NASA, 2023",
    "source_page": "https://www.nasa.gov/...",
    "credit": "NASA",
    "depicted_subject": "RAMFIRE additively manufactured nozzle during a hot-fire test",
    "verified": true
  }]
}
```

Reject an image if the subject, source page, credit, or usage status cannot be established. The on-slide caption and manifest must agree.

### Zero-reuse image contract

Default to one independent original visual asset per replaceable content-image frame across the whole deliverable. The following are the same asset, not new assets:

- an identical file under another name or extension;
- recompression, resizing, flipping, recoloring, or format conversion;
- multiple crops or screenshots of the same original;
- a full paper/PDF page and a figure crop taken from that same page.

The task-local source manifest must have one record per picture frame. Before replacement and before delivery, audit normalized file paths, source-file SHA256, perceptual hashes, and source identities. Perceptual near-matches require visual review and must be replaced unless the user explicitly requested a same-image comparison. Store the audit internally, never in the delivery directory.

```powershell
python script/audit_ppt_image_uniqueness.py `
  --sources <sources.json> `
  --repo-root <repo-root> `
  --expected-count <replaceable-picture-frame-count> `
  --output <internal-image-uniqueness-audit.json>
```

## Required validation

A deck passes only when:

- every slide resolves to one actual PowerPoint master in the output package;
- every slide layout is approved by the selected family;
- fixed shell geometry matches its family role source;
- every正文 page has a module contract and unchanged module geometry signature;
- every module remains within the family safe region without scaling or relayout;
- the module-rhythm audit shows no unexplained adjacent duplicate and no more than two uses of one module in any five consecutive正文 pages; any approved parallel set contains at most three adjacent pages and records its shared comparison dimension;
- text capacities, title lines, and style roles pass;
- all images and captions match the source manifest;
- the image-source manifest has one record per replaceable picture frame, and path, SHA256, perceptual-hash, source-identity, and visual near-duplicate checks show no unexplained reuse;
- every rendered page has been visually reviewed;
- the delivery directory contains only the final PPTX.

Use these mechanical gates before visual review:

```powershell
python script/template_ppt.py verify-module --module <module.json> --family <family.json>
python script/template_ppt.py verify-layered --input <working-output.pptx>
python script/template_ppt.py verify-master --input <working-output.pptx> --family <family.json>
```
