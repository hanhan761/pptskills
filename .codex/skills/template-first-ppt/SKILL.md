---
name: template-first-ppt
description: Build, revise, curate, or develop user-reviewed PowerPoint templates with one consistent master family and reusable below-title content modules. Use when creating or editing PPT/PPTX files, learning a reference deck's master shell, extracting content modules, maintaining this project's master/module library, developing one-page or explicitly requested batch candidates in 模板研发部门, or replacing slide copy and real images without inventing production layouts.
---

# Template-first PPT

Build every deck from two verified layers:

1. A single **master family** supplies the real PowerPoint master, approved layouts, cover, agenda, section states, closing treatment, and the fixed header/title shell.
2. A **content module** supplies the already-designed shapes below the title shell.

Never draw a slide or mix unrelated full-page masters. The stable visual identity comes from the first layer; only the second layer varies by communication job.

## Establish the workspace

1. Locate the nearest ancestor containing both `AGENTS.md` and `模板/`.
2. Read `AGENTS.md` before changing a presentation.
3. Read [references/library-contract.md](references/library-contract.md) before learning a master family, extracting modules, or writing a build plan.
4. Use repository-root tools in `script/` only for mechanical copying, replacement, indexing, rendering, and validation. Keep code and internal evidence out of the delivery directory.

Complete this stage only when the repository, task directory, master-family index, and relevant contracts are known.

## Use the correct mental model

“Master family” means more than the XML `slideMaster`. It includes:

- exactly one real PowerPoint master;
- its approved slide layouts;
- visually fixed slide-level shell shapes, such as a title tab, line, logo, page marker, or title text box.

Determine the shell/content boundary from full-deck renders and repeated visual behavior. Do not infer it solely from XML location: an object can live on the slide while still being part of the fixed shell.

Treat existing one-slide files in `模板/封面`, `模板/目录`, and `模板/内容` as legacy candidate sources. Do not compose them directly across different masters.

## Learn a master family

1. Render and inspect every source slide.
2. Confirm page size, actual master count, layout usage, repeated shell shapes, and the visual top boundary of the content region.
3. Choose representative role pages for cover, agenda, section transition, content shell, and closing. If a role is absent, record it as absent; only derive a closing from an existing same-family role page through text/image replacement.
4. Copy the source deck unchanged to `模板/母版/<家族>/母版源.pptx`. Do not reconstruct it.
5. Write `family.json` using the schema in the library contract. Record the source hash, one real master, approved layouts, role pages, fixed shell shapes, content boundary, and brand tokens observed in the source.
6. Register the family in `模板/母版/索引.json`. Only a visually reviewed and structurally verified family can be active or default.
7. Run `python script/template_ppt.py verify-master --input <母版源.pptx> --family <family.json>`.

Stop if the source uses multiple masters without a clear single family, or if fixed shell and content cannot be separated reliably.

## Curate content modules

1. Review every candidate page, but retain only representative modules. Give each page one decision: `keep`, `merge`, or `discard`.
2. For a正文 candidate, separate the fixed shell from shapes whose visible bounds lie wholly inside the source family's content region. Preserve the module shapes, z-order, absolute geometry, text runs, image frames, and theme references.
3. Classify candidates by communication job and geometry before topic or color. Merge same-layout variants and reject unfinished, locked, or weakly reusable content.
4. Write a module contract recording source deck, source slide, source family, exact shape identities, bounding box, geometry signature, compatibility, and all text-capacity slots.
5. Determine title and text line counts from rendered visual intent or explicit user correction. Record mixed-color run roles and required emphasis roles.
6. Run `python script/template_ppt.py verify-module --module <module.json> --family <family.json>` to validate source hash, exact shape set, geometry signature, page size, compatibility, and safe-region containment.
7. Test-mount the module at its original coordinates on a copied content-shell page. Reject it if any shape crosses the target safe region, requires scaling/reflow, or visually conflicts with the shell.
8. Render and compare the mounted result before adding the module to the active index.

Complete this branch only when every retained module has provenance, a valid contract, an unchanged geometry signature, and a verified compatible master family.

## Develop candidate modules

Use this branch only when the user explicitly requests a new template in `模板研发部门/`. It is a narrow research exception, not permission to improvise layouts inside a production deck.

1. Default to one candidate. If the user explicitly requests N templates, define exactly N distinct communication jobs and keep each page independently reviewable and contractable.
2. Select one target master family and duplicate its approved content-shell page. Keep its real master, layout, fixed header/title shell, theme, typography, and palette unchanged. Do not develop a new master or role page unless the user explicitly authorizes that separate scope.
3. Create one native, editable, below-title content-module candidate per communication job. In an explicit batch, make the geometries materially different rather than recoloring one card system. Prefer a stable text hierarchy and 1–4 meaningful real-image frames with caption/source positions when imagery adds evidence. Give every visible content-image frame in the review deck an independent original asset; a renamed, recompressed, recolored, flipped, or cropped copy is still the same asset. Never use a whole-slide image or flattened SVG as the page.
4. Present every candidate as a finished sample slide, not a wireframe or template instruction page. Fill it with substantive example copy and real evidence. Keep extension notes, replaceable-field descriptions, sibling variants, and phrases such as “适用” or “迁移方式” in the internal review record, not on the visible slide.
5. Apply the user's “large, full, few, strong” standard: use the safe region decisively, consolidate small floating boxes, keep one dominant visual, and use projection-readable type. Before setting sizes, measure at least one user-calibrated slide and two reviewed/active templates with comparable geometry. Match their relative hierarchy instead of imposing one global size. For `public-blue-white`, a current empirical reference is 24–26 pt title bands, 18–28 pt section labels, 16–20 pt wide body text, approximately 11.5–15 pt narrow-card body text when a reviewed source supports it, 18–20 pt conclusions, and 9.5–11 pt captions. If a larger size breaks a narrow horizontal column, compress the copy or convert the candidate to a vertical label/body hierarchy; do not flatten every role to a similar size.
6. Make the communication structure easy to extend through separately reviewed sibling variants; do not create a module whose production use depends on adding, moving, resizing, or reflowing shapes.
7. Render and inspect every candidate. Reject a page that is structurally valid but visually sparse, fragmented, or dependent on small text. Keep only one current review PPTX in `模板研发部门/create1/`: one page by default, or exactly N pages for an explicitly requested batch. Keep briefs, sources, renders, and per-page review states in `模板研发部门/_internal/`. Revise the same review file in place while candidates are `draft` or `revise`.
8. Do not use or index any candidate until the user explicitly approves that page or the whole named batch. On approval, freeze each approved page as its own source in `模板研发部门/已审核/`, write an independent module/capacity contract, test-mount it, run module/layer/master validation, render it again, and only then add it to the active module index.

User approval is necessary but does not replace mechanical validation. A rejected or unverified candidate remains isolated from production.

## Build a deck

1. Turn the request into a slide-by-slide narrative and define each page's communication job, copy, and real-image need.
2. Select exactly one `master_family` before selecting content modules. Use the default family only when the user did not specify another style.
3. Confirm that the family supplies every required role. Cover, agenda, section pages, content shell, and closing must all come from this family.
4. Select one verified content module for every正文 page. If no compatible module exists, report the missing module type and stop the production build. Offer the separate template-R&D branch only when the user wants a new structure; resume production only after that candidate is approved, contracted, verified, and indexed.
5. Write the layered build plan from the library contract. Every slide records the same `master_family`, a shell role/source, and a content module or same-family role page. Audit module rhythm before composition: do not repeat the same module on adjacent正文 pages or more than twice in any five正文-page window. Permit at most three adjacent uses only for an explicit parallel comparison/case set, and record the shared dimension and reason.
6. Run `python script/template_ppt.py compose-family --plan <plan.json> --output <working.pptx>`. The command always starts from the selected family's `母版源.pptx`: same-source modules are duplicated inside that family by stable slide ID, while compatible cross-source modules copy only their registered below-title shapes onto a duplicated target content shell. It must never insert the source slide, source layout, or source master.
7. Mount each content module at its recorded absolute coordinates and source z-order without scaling or rearranging shapes. For cross-source modules, require a source-master fingerprint in the contract and verify that the final deck still contains only the target master.
8. Replace only existing text and pictures. Honor capacity contracts, explicit title lines, and original run-style roles. Use real traceable images with accurate captions and a task-local source manifest. Default to zero content-image reuse across the entire deck: bind one independent original asset to each replaceable picture frame. Treat copies, renames, recompression, format conversion, crop variants, flips, color edits, screenshots, and a paper page plus a crop from that same page as reuse rather than new material.
9. Audit image uniqueness before replacement and again against the final source manifest. Require unique normalized paths, source-file SHA256 values, perceptual hashes, and source identities; review every perceptual near-match. Run `python script/audit_ppt_image_uniqueness.py --sources <sources.json> --repo-root <repo> --expected-count <picture-frame-count> --output <internal-audit.json>`. Any unexplained duplicate is a hard failure. A repeated evidence image is allowed only when the user explicitly requests that same-image comparison and the exact pages and reason are recorded.
10. Run `verify-module` for every selected module, `python script/template_ppt.py verify-layered --input <working.pptx>`, and `verify-master` for the complete deck. Any master mismatch, unapproved layout, shell drift, module geometry drift, stale provenance, or failed image-uniqueness audit is a failure.
11. Render every page and inspect the complete deck for a uniform header/title system, text fullness, overflow, fixed title lines, keyword color grammar, clipping, overlap, image crop, caption accuracy, low resolution, typos, style jumps, narrative continuity, and visual reuse that hashing may not expose.
12. Deliver exactly one clearly named `.pptx`. Keep plans, provenance, source manifests, images, previews, and code internal, then remove them from the delivery directory.

## Hard stops

- Do not use more than one master family in an output deck.
- Do not call a collection of similar colors a “master”; verify the actual PowerPoint master, approved layout, and fixed slide-level shell together.
- Do not directly compose legacy full-page templates from different sources.
- Do not invent a layout in a production deck. The only layout-development exception is one user-requested candidate below an existing family shell in `模板研发部门/`, and it remains unusable until explicit approval and full validation.
- Do not use one familiar module as a long-deck fallback. Fail the rhythm audit and select another verified module—or enter gated template R&D—when adjacent or five-page-window repetition has no explicit parallel rationale.
- Do not add, delete, move, resize, regroup, or restyle module shapes. Mounting a registered module at its recorded coordinates is the only pre-authorized structural operation.
- Do not shrink fonts, expand text boxes, or rely on accidental wrapping.
- Do not flatten mixed-color titles or invent a new palette.
- Do not use an unverified or AI-generated image when real documentary or technical imagery is expected.
- Do not reuse one content image in two frames or pages by copying, renaming, re-encoding, cropping, flipping, recoloring, or extracting a crop from a document page already used elsewhere. Fixed non-evidentiary master-family branding is outside this content-image rule; semantic photographs, screenshots, diagrams, and illustrations inside modules are not.
- Do not invent captions, depicted subjects, credits, dates, or locations.
- Do not trust an image replacement until the rendered slide shows the new subject. Normalize the media content type and remove stale linked/SVG alternate image references when present.
- Do not generate a whole-slide image to bypass the template rule.
- Do not extract or build through a blank presentation.
- Do not declare completion until every slide uses the selected real master, all contracts pass, and every rendered slide has been visually reviewed.
