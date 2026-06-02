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
3. If doing the workflow manually and the file is `.doc` or `.ppt`, run `scripts/convert_legacy_office.ps1` first. It tries Microsoft Office COM, then WPS if present, then LibreOffice. Each backend runs in an isolated worker with a 45-second default timeout before fallback.
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

Visual cache entries are isolated by reading mode and analysis profile. A local-only run does not suppress a later OpenAI Vision run, and cloud-derived summaries are not reused as local-only results.

## Dependency Bootstrap

Run `scripts/bootstrap_deps.ps1 -DryRun -IncludeSystemTools` to inspect missing dependencies without installing anything.

Run `scripts/bootstrap_deps.ps1 -IncludeSystemTools` to create the skill-local `.venv` and install Python packages: `openai`, `markitdown-ocr`, `rapidocr`, and `onnxruntime`.

Run `scripts/bootstrap_deps.ps1 -IncludeSystemTools -InstallSystemTools` only when the user wants missing fallback tools installed. It prefers `winget`/`choco` packages for LibreOffice, Poppler, Tesseract, and optional WPS. Python packages stay inside `C:\Users\Huang\.codex\skills\office-reader\.venv`.

When using `read_office.py`, pass `--install-missing-deps --install-system-tools` together to request system fallback installation. `--install-system-tools` alone is rejected.

After attempting system-tool installation, bootstrap reruns discovery. Inspect each returned `system_tools` entry for final `available`, `source`, `install_attempted`, `install_exit_code`, and `install_result` values. Do not treat a package-manager invocation as proof that the tool became available.

Portable Poppler/Tesseract builds can be used without system install. Put the folders under `scripts/../tools`, the parent skills directory, or the current workspace, or set `OFFICE_READER_TOOL_PATHS` to their `bin` directories.

Set `SOFFICE_PATH` or `LIBREOFFICE_PATH` to a LibreOffice executable or installation directory when LibreOffice is installed outside the usual locations. Backend discovery, dependency dry-run, and conversion honor the same overrides.

## Output Contract

Every successful read produces:

- `<basename>.full.md`: full Markdown transcript with headings, slide boundaries, tables, comments, revisions, and notes where extractable.
- `<basename>.manifest.json`: structured extraction data plus `reading_mode`, `visual_analysis`, and non-fatal OOXML `warnings`. See `references/output_schema.md`.
- `<basename>.report.md`: structured reading report with summary, outline, tables, comments/revisions, notes, visual findings, risks, and artifacts.

If unknown same-name artifacts already exist in the selected output directory, preserve them and use an `office-reader-run-<guid>` subdirectory. Same-source office-reader reruns may reuse their prior artifact paths.

If legacy conversion fails, return the conversion JSON and explain which backends were unavailable. Do not pretend a `.doc` or `.ppt` was read when no normalized `.docx` or `.pptx` exists.

## Legacy Conversion Backends

Use `scripts/discover_office_backends.ps1 -InputExtension .doc -Format json` or `.ppt` to inspect available conversion backends. The priority order is fixed:

1. Microsoft Office COM through `Word.Application` or `PowerPoint.Application`.
2. Optional WPS Office fallback, including `WPS_PATH`, `WPP_PATH`, `wps.exe`, `wpp.exe`, App Paths registry entries, common WPS install folders, and WPS COM ProgIDs.
3. LibreOffice/`soffice`.

Detailed behavior is in `references/backend_fallbacks.md`.

## Reading Expectations

For `.docx`, extract document metadata, paragraphs, heading levels, tables, comments, tracked insertions and deletions, media references, and a visual-review flag when drawings or media are present.

For `.pptx`, extract presentation metadata, slide order, slide text, tables, speaker notes, comments, media references, and per-slide visual-review flags when media is present.

XML extraction does not fully read text inside images, screenshots, rasterized charts, SmartArt, or complex embedded objects. In `balanced` or `complete` mode, the visual pipeline renders pages/slides when possible, applies local OCR, optionally asks OpenAI vision for diagram/chart/screenshot interpretation, caches results, and writes findings back into `visual_findings`.

If optional comments, notes, metadata, or relationship XML is malformed, preserve readable body content and inspect manifest `warnings` plus the report risks section before treating the read as complete.

## Quick Commands

```powershell
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\read_office.py C:\path\file.pptx --out-dir C:\path\out --mode complete
python scripts\read_office.py C:\path\file.doc --out-dir C:\path\out --mode fast --legacy-timeout-seconds 45
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced --no-openai-vision
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -DryRun -IncludeSystemTools
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\discover_office_backends.ps1 -InputExtension .doc -Format json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\convert_legacy_office.ps1 -InputPath C:\path\file.doc -OutputDir C:\path\out -TimeoutSeconds 45
python scripts\read_docx.py C:\path\file.docx --out-dir C:\path\out
python scripts\read_pptx.py C:\path\file.pptx --out-dir C:\path\out
python scripts\visual_analysis.py C:\path\out\file.manifest.json --normalized-file C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\assemble_report.py C:\path\out\file.manifest.json
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\render_preview.ps1 -InputPath C:\path\file.pptx -OutputDir C:\path\preview
python scripts\smoke_office_reader.py --doc C:\path\file.doc --docx C:\path\file.docx --pptx C:\path\file.pptx --derive-ppt-from-pptx --visual-timeout-seconds 180
```

## Local Smoke Validation

Use `scripts/smoke_office_reader.py` for repeatable real-document validation. It runs `.doc`, `.docx`, `.ppt`, and `.pptx` in `fast` mode, then reruns modern `.docx` and `.pptx` inputs in `complete --no-openai-vision` mode. Pass `--derive-ppt-from-pptx` when a local `.ppt` fixture is unavailable.

Keep real source documents, derived `.ppt` files, caches, and smoke output local. Derived PPT fixtures use unique `<stem>.derived-<guid>.ppt` names to preserve existing files. Do not commit them to the skill repository.

For long Word documents, raise `--visual-timeout-seconds` during smoke validation so Office COM preview has time to fail cleanly and LibreOffice fallback can still render pages.

PowerShell JSON output is ASCII-safe so Unicode and space-containing paths remain parseable when stdout is redirected across Windows console code pages.

## Common Mistakes

- Do not use this skill for PDF or XLSX files; it is intentionally scoped to Word and PowerPoint.
- Do not rely only on Markdown when comments, revisions, speaker notes, or media matter. Check the manifest and report.
- Do not silently skip legacy conversion. If all backends are missing, report that Office COM, WPS, and LibreOffice were unavailable.
- Do not claim image-only charts or screenshots were fully read from XML. Use `balanced` or `complete` mode and report any remaining visual gaps.
- Do not install global Python packages for this skill. Use `bootstrap_deps.ps1` so dependencies stay in the skill-local `.venv`.
