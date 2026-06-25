---
name: pdf-section-to-docx
description: Use when Codex needs to convert a PDF into section-by-section Word documents with full text extraction, OCR for scanned pages, table-to-prose rewriting, heading hierarchy preservation, or section-based file naming. Trigger on requests to split a PDF by headings, export each section to DOCX, keep every title level, or avoid omissions in OCR-based document conversion.
---

# PDF Section To DOCX

## Overview

Convert one PDF into multiple Word files, one per detected section, while preserving heading hierarchy and rewriting every detected table as prose. Keep the document guide title only once at the top of each exported DOCX, remove repeated in-body guide-title lines, and avoid treating parenthesized list items like `（5）...` as new sections.

## Workflow

1. Run the bundled script instead of ad-hoc PDF handling.
2. Let the script extract direct text first and trigger OCR on low-text pages.
3. Let the script rewrite each table into paragraph form before DOCX export.
4. Review `manifest.json` and spot-check generated DOCX files when the source layout is unusual.
5. Stop and report any prerequisite or confidence issue instead of claiming complete extraction.

## Command

Run:

```powershell
python D:\codex-data\skills\pdf-section-to-docx\scripts\pdf_section_to_docx.py <input.pdf>
```

Optional flags:

- `--output-dir <dir>` to choose the export folder; when omitted, output goes to `D:\Codex_output\<pdf-stem>-sections`
- `--ocr-lang chi_sim+eng` to set Tesseract languages
- `--min-direct-chars 80` to tune when OCR starts

## Output Contract

- Emit one `.docx` per detected section.
- Name each file from the current section title after filename sanitization.
- Write the document guide title once at the top of each exported DOCX when the source contains a `...指南` title.
- Remove duplicate in-body occurrences of the same guide title.
- Preserve the full heading chain inside each DOCX by re-adding ancestor headings before section content.
- Convert each detected table into prose paragraphs instead of leaving tabular layout in the DOCX, preferably carrying table title and source into the prose block.
- Write `manifest.json` alongside the DOCX files so the run can be audited.
- When no explicit `--output-dir` is passed, save the run under `D:\Codex_output`.

## Fail-Closed Rules

- If the PDF appears scanned and `tesseract` is unavailable, stop with an explicit prerequisite error.
- If no headings can be recovered, stop and report that section-based splitting was not possible from the source structure.
- Do not summarize, compress, or intentionally omit text to make the output prettier.
- If a PDF layout is ambiguous enough that heading hierarchy may be wrong, say so and recommend manual review of the affected outputs.

## Resources

### scripts/

- `scripts/pdf_section_to_docx.py`: bundled extractor and exporter
- `scripts/test_pdf_section_to_docx.py`: helper-level regression tests

### references/

- `references/ocr-prerequisites.md`: installation and runtime notes for OCR pages
