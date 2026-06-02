---
name: office-reader
description: Use when Codex needs to deep-read Word or PowerPoint files, including .doc, .docx, .ppt, .pptx, legacy Office conversion, full Markdown extraction, structured reports, comments, tracked revisions, tables, speaker notes, embedded media, OCR, rendered pages, screenshots, charts, diagrams, or visual-review gaps.
---

# Office Reader

## Overview

Use this skill to turn Word and PowerPoint files into a full Markdown transcript, a structured manifest, and a reading report. Prefer this skill when the user asks to read, inspect, analyze, summarize, audit, or understand `.doc`, `.docx`, `.ppt`, or `.pptx` files, especially when images, screenshots, charts, or diagrams may contain important evidence.

## Workflow

1. Resolve the input file path and create an output directory near the source unless the user requested another location.
2. Prefer the unified entrypoint: `python scripts/read_office.py <file> --out-dir <out-dir> --mode balanced`.
3. If doing the workflow manually and the file is `.doc` or `.ppt`, run `scripts/convert_legacy_office.ps1` first. It tries Microsoft Office COM, then WPS if present, then LibreOffice.
4. Run the reader that matches the normalized file:
   - `.docx`: `python scripts/read_docx.py <file.docx> --out-dir <out-dir>`
   - `.pptx`: `python scripts/read_pptx.py <file.pptx> --out-dir <out-dir>`
5. Enrich visuals with `python scripts/visual_analysis.py <manifest> --normalized-file <file.docx|pptx> --out-dir <out-dir> --mode balanced`.
6. Build the report with `python scripts/assemble_report.py <basename>.manifest.json`.
7. Read the manifest and report before answering the user. Call out visual-review gaps whenever preview rendering, OCR, or OpenAI vision was skipped or failed.

## Reading Modes

- `--mode fast`: OOXML extraction plus media flags only. Use when the user wants a quick text pass.
- `--mode balanced`: default. Render pages/slides only when visual risk is detected, OCR/vision-analyze selected items, and use cache.
- `--mode complete`: render and analyze every page/slide. Use when the user asks to fully understand image-heavy or evidence-critical documents.

Use `--no-openai-vision` when cloud visual analysis is not allowed. Without `OPENAI_API_KEY`, local OCR still runs when available and the report clearly lists remaining visual gaps.

## Dependency Bootstrap

Run `scripts/bootstrap_deps.ps1 -DryRun -IncludeSystemTools` to inspect missing dependencies without installing anything.

Run `scripts/bootstrap_deps.ps1 -IncludeSystemTools` to create the skill-local `.venv` and install Python packages: `openai`, `markitdown-ocr`, `rapidocr`, and `onnxruntime`.

Run `scripts/bootstrap_deps.ps1 -IncludeSystemTools -InstallSystemTools` only when the user wants missing fallback tools installed. It prefers `winget`/`choco` packages for LibreOffice, Poppler, Tesseract, and optional WPS. Python packages stay inside `C:\Users\Huang\.codex\skills\office-reader\.venv`.

Portable Poppler/Tesseract builds can be used without system install. Put the folders under `scripts/../tools`, the parent skills directory, or the current workspace, or set `OFFICE_READER_TOOL_PATHS` to their `bin` directories.

## Output Contract

Every successful read produces:

- `<basename>.full.md`: full Markdown transcript with headings, slide boundaries, tables, comments, revisions, and notes where extractable.
- `<basename>.manifest.json`: structured extraction data plus `reading_mode`, `visual_analysis`, and `completeness_score`. See `references/output_schema.md`.
- `<basename>.report.md`: structured reading report with summary, read completeness, outline, tables, comments/revisions, notes, visual findings, risks, and artifacts.

If legacy conversion fails, return the conversion JSON and explain which backends were unavailable. Do not pretend a `.doc` or `.ppt` was read when no normalized `.docx` or `.pptx` exists.

## Legacy Conversion Backends

Use `scripts/discover_office_backends.ps1 -InputExtension .doc -Format json` or `.ppt` to inspect available conversion backends. The priority order is fixed:

1. Microsoft Office COM through `Word.Application` or `PowerPoint.Application`.
2. Optional WPS Office fallback, including `WPS_PATH`, `WPP_PATH`, `wps.exe`, `wpp.exe`, App Paths registry entries, common WPS install folders, and WPS COM ProgIDs.
3. LibreOffice/`soffice`.

Detailed behavior is in `references/backend_fallbacks.md`.

## Reading Expectations

For `.docx`, extract document metadata, paragraphs, heading levels, tables, comments, tracked insertions and deletions, move revisions, media references, and a visual-review flag when drawings or media are present. Header, footer, footnote, and endnote text are included in the transcript and manifest with `part_type`/`part` location metadata when those OOXML parts are present. Block-level content controls are unpacked and tagged with `container: content_control`. Textbox-originated paragraphs, revisions, comments, and media are separated from their outer paragraph and tagged with `container: textbox`.

For `.pptx`, extract presentation metadata, slide order, slide text, tables, speaker notes, comments, media references, and per-slide visual-review flags when media is present. PPTX visual findings include an object inventory for images, charts, SmartArt, OLE objects, video, and audio when the relationship metadata is present in the slide XML.

XML extraction does not fully read text inside images, screenshots, rasterized charts, SmartArt, or complex embedded objects. In `balanced` or `complete` mode, the visual pipeline renders pages/slides when possible, applies local OCR, optionally asks OpenAI vision for diagram/chart/screenshot interpretation, caches results, and writes findings back into `visual_findings`.

The `completeness_score` is a conservative coverage signal. Do not treat OOXML media hints, fast-mode placeholders, or rendered-page bookkeeping as proof that visual content was understood. Report `unverified_visual_count` and score signals when answering the user.

## Preview Backend Health Memory

Preview rendering remembers unhealthy Microsoft Office COM behavior separately from legacy conversion. If Word or PowerPoint COM preview times out or returns invalid output, `scripts/render_preview.ps1` records that backend as unhealthy for the normalized extension and later preview rendering skips directly to fallback backends such as LibreOffice.

This does not change `.doc` or `.ppt` legacy conversion priority: Microsoft Office COM remains first, followed by optional WPS and LibreOffice.

The default preview health file is `.office-reader-cache/preview-backend-health.json` under the skill directory. Set `OFFICE_READER_PREVIEW_HEALTH_PATH` only for tests or when you need an isolated cache.

## Quick Commands

```powershell
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\read_office.py C:\path\file.pptx --out-dir C:\path\out --mode complete
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced --no-openai-vision
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -DryRun -IncludeSystemTools
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\discover_office_backends.ps1 -InputExtension .doc -Format json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\convert_legacy_office.ps1 -InputPath C:\path\file.doc -OutputDir C:\path\out
python scripts\read_docx.py C:\path\file.docx --out-dir C:\path\out
python scripts\read_pptx.py C:\path\file.pptx --out-dir C:\path\out
python scripts\visual_analysis.py C:\path\out\file.manifest.json --normalized-file C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\assemble_report.py C:\path\out\file.manifest.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\render_preview.ps1 -InputPath C:\path\file.pptx -OutputDir C:\path\preview
```

## Common Mistakes

- Do not use this skill for PDF or XLSX files; it is intentionally scoped to Word and PowerPoint.
- Do not rely only on Markdown when comments, revisions, speaker notes, or media matter. Check the manifest and report.
- Do not silently skip legacy conversion. If all backends are missing, report that Office COM, WPS, and LibreOffice were unavailable.
- Do not claim image-only charts or screenshots were fully read from XML. Use `balanced` or `complete` mode and report any remaining visual gaps.
- Do not install global Python packages for this skill. Use `bootstrap_deps.ps1` so dependencies stay in the skill-local `.venv`.
