---
name: ppt6
description: "Create editable PowerPoint decks through a mandatory two-stage workflow: absorb an existing template, generate a visual draft with imagegen, wait for human approval, then rebuild the approved design with editable objects and real sourced images. Use only when the active model is GPT-6."
---

# PPT6

PPT6 is a gated PowerPoint workflow for one slide, a small set of slides, or a complete deck. It separates visual direction from editable production:

1. understand the task and produce image-based draft pages;
2. stop for human review;
3. after explicit approval, rebuild the approved pages as an editable `.pptx`.

Do not skip the review gate, silently replace a rejected draft, or deliver the draft image as the final PowerPoint.

## Model gate

PPT6 is GPT-6-only. Before doing any image generation, template analysis, external image search, or PPT production:

- obtain the active model ID from runtime metadata or the caller;
- run `scripts/check_model_gate.py --model-id <active-model-id>` when a model ID is available;
- accept only an ID beginning with `gpt-6` or `gpt6`;
- if the ID is another model or cannot be verified, stop and report that PPT6 cannot run. Do not fall back to another model or continue partially.

The gate is a capability boundary, not a request to change models. PPT6 cannot upgrade or switch the active model.

## Phase 1: visual draft

Create a task-local run directory and record:

- audience, purpose, page count or scope, aspect ratio, language, and factual constraints;
- the narrative spine and the role of every planned page;
- whether a source/template deck exists;
- image slots, real-image search needs, logo needs, and source requirements.

When a template is supplied, inspect and render it before drafting. Absorb its visual grammar: page size, master/layout family, title hierarchy, spacing, color, typography, image treatment, footer/chrome, and repeated content structures. Preserve meaning-bearing template elements in the later editable build. Read [references/workflow.md](references/workflow.md) for the template-audit record.

Use the built-in `$imagegen` workflow for visual exploration and for an original logo/mark when appropriate. Use exact copy in the draft; if imagegen cannot render text reliably, compose a deterministic text overlay on top of the generated visual so the human reviews the real wording rather than hallucinated lettering. A generated logo is an image asset, not permission to flatten the whole slide.

For photographs, screenshots, devices, places, people, experiments, and events, find real and traceable source assets. Do not use an AI-generated scene as evidence or as a substitute for a real source image. Record source URL, creator or institution, title/description, retrieval date, and usage notes.

Produce reviewable PNG/JPEG page images, a contact sheet for multi-page work, a copy manifest, an asset/source manifest, and `review_state.json`. The draft may be a visual composition, but it must be concrete enough to judge hierarchy, image choice, crop, density, and wording.

## Human approval gate

Stop after Phase 1 and show the image draft. Ask for explicit approval or revision instructions. Approval must identify the scope: every page, named pages, or the complete draft. “看一下”“大概可以” or silence is not approval.

Record approval in `review_state.json` and validate it with `scripts/validate_review_state.py --state <path> --phase build`. If any page is rejected or marked for revision, remain in Phase 1 and regenerate only the affected draft assets/pages. Never create or overwrite the final PPT before the validator passes.

## Phase 2: editable PPT production

Run the model gate again, then validate the human approval state. Build from the supplied template/source deck when one exists; otherwise use the approved visual draft as a reference and reconstruct the page with native editable objects.

The final deck must use:

- native text boxes for all editable text;
- native shapes, tables, charts, and diagrams wherever the object is meant to be edited;
- real sourced photos or screenshots inside editable image frames;
- a generated logo only as an explicitly approved image asset, never as a full-slide screenshot;
- the source template’s geometry and visual system when a template was supplied.

Never place the approved full-slide draft image as the final slide background to disguise an uneditable build. A background photo or logo image may remain raster, but the slide’s text, structure, and meaning-bearing layout must remain editable.

Replace draft proxies with final real images, preserve their source records, and keep captions/credits accurate. Render the finished deck and compare it with the approved draft for composition, text, crops, and visual hierarchy. Run a final editability check: no page may consist only of one full-slide image, and every required text element must be addressable as a native object.

## Copy and visual quality

Write concise Chinese (or the requested language): state the judgment first, then evidence or action; remove filler, defensive framing, repeated claims, generic “AI” phrasing, and unexplained buzzwords. Prefer short sentences and concrete verbs. Do not invent facts, citations, logos, or image provenance. Read [references/copy-and-asset-rules.md](references/copy-and-asset-rules.md) before drafting.

## Required stop conditions

Stop and report instead of improvising when:

- the active model is not verifiably GPT-6;
- the template cannot be read or rendered reliably;
- a required real image or source cannot be verified;
- the image draft has not received explicit human approval;
- the approved visual cannot be reconstructed as editable objects without flattening or silently changing its intent.

Keep intermediate drafts, prompts, manifests, approvals, renders, and QA evidence inside the task-local run directory. Deliver only the final `.pptx` unless the user asks for the review package too.

For the state schema and approval examples, read [references/review-state.md](references/review-state.md).
