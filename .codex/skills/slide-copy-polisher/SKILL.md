---
name: slide-copy-polisher
description: Polish PowerPoint slide copy, captions, and speaker notes into direct, clean Chinese while preserving facts, evidence, capacity contracts, fixed title lines, and the existing layout.
---

# Slide-copy-polisher

Use this skill for PPT page text, module text, captions, speaker notes, and text replacement maps. It changes wording only. It does not change the master family, module geometry, font size, text boxes, images, title-line contracts, or note placement.

## Style contract

- State the page's judgment first. Follow with evidence, explanation, consequence, or action.
- Remove defensive preambles, indirect argument, empty qualifications, and self-justifying transitions.
- Do not repeat the same conclusion in the title, body, and conclusion unless each occurrence adds a different function.
- Keep one main action per sentence. Split long sentences into short, complete sentences.
- Do not use em dashes. Prefer full stops, commas, colons, or semicolons.
- Keep facts, numbers, terminology, citation meaning, uncertainty, and conclusion strength unchanged.
- Do not add filler words to meet a capacity minimum. If text does not fit, rewrite the idea or select a compatible module.

Read [references/style-contract.md](references/style-contract.md) for role-specific rules and examples before a substantial polish.

## Workflow

1. Inventory every text slot and label its role: main title, module title, body, evidence label, conclusion, caption, or note. Keep page number, shape name, occurrence, line count, and capacity limits beside the text.
2. Write one plain-language judgment for the page. If the page has several blocks, give each block one distinct job and remove blocks that only restate another block.
3. Rewrite in this order: conclusion, evidence, explanation, action. Replace abstract nouns with concrete verbs when the fact allows it. Remove preambles such as “需要指出的是” and “值得注意的是” unless the phrase itself carries information.
4. Split sentences at natural actions or evidence boundaries. Remove stacked clauses, repeated subjects, and chained qualifiers. Replace em dashes with ordinary punctuation.
5. Preserve the original run roles and explicit lines. Check visible character counts, per-line counts, title line counts, and semantic completeness. Never truncate a sentence automatically.
6. Run the mechanical audit:

   ```powershell
   python .codex/skills/slide-copy-polisher/scripts/audit_prose_style.py `
     --input <copy-map.json> `
     --output <internal-prose-audit.json>
   ```

   Resolve every finding. The audit is a gate, not a replacement for reading the rendered slide.

7. For PPT work, replace only existing text or NotesPage body text, then render and inspect for overflow, accidental line changes, semantic repetition, and mismatch between copy and visual evidence.

## Output contract

Return the polished copy with its original selectors and roles. Keep a concise internal change record when the wording materially changes a claim. Do not expose audit files, source manifests, or working maps in the final delivery directory.

## Hard stops

- Do not invent evidence, citations, numbers, captions, or uncertainty.
- Do not make a claim stronger or weaker just to sound cleaner.
- Do not solve a capacity problem by shrinking text, expanding a box, changing line count, or moving a shape.
- Do not pass text containing unresolved defensive markers, substantive repetition, complex sentences, or em dashes.
