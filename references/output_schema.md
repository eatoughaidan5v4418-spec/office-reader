# Output Schema

`office-reader` writes a manifest JSON file beside the Markdown transcript and report.

The unified `scripts/read_office.py` command also prints a small JSON object to stdout with `full_markdown`, `manifest`, and `report` paths. Stdout/stderr are configured as UTF-8 so callers can decode output paths reliably when paths contain Chinese characters or spaces. PowerShell 5 callers should read generated JSON files with `-Encoding UTF8`.

## Top-Level Fields

- `source`: original file path and name.
- `normalized_file`: actual `.docx` or `.pptx` file read after legacy conversion.
- `conversion`: conversion status, backend, output path, and messages. For legacy `.doc/.ppt` text fallback, `status` is `text_fallback` and `output_path` points to the extracted `.legacy-text.txt` file.
- `document_type`: `docx` or `pptx`.
- `reading_mode`: `fast`, `balanced`, or `complete`.
- `metadata`: package metadata from `docProps/core.xml` when present.
- `structure`: paragraphs/headings for Word or ordered slides for PowerPoint. Word entries may include `part_type` and `part` for non-body sources such as headers, footers, footnotes, and endnotes, and `container` values such as `content_control` or `textbox` when text came from those containers.
- `tables`: extracted table rows with document, slide, or Word part location.
- `comments`: Word comments or PowerPoint comments. Word comments may include `anchor_text` when `commentRangeStart/commentRangeEnd` identifies the commented span. Word comments referenced inside a table cell may include `table_index`, `row_index`, `cell_index`, `part_type`, and `part`.
- `revisions`: Word tracked insertions and deletions. Word revisions inside a table cell may include `table_index`, `row_index`, `cell_index`, `part_type`, and `part`.
- `notes`: PowerPoint speaker notes.
- `visual_analysis`: visual pipeline status, selected mode, rendered page count, analyzed item count, cache hits, backends, and messages.
- `visual_findings`: flags for media, drawings, image-heavy content, OCR text, rendered-page observations, vision summaries, diagram summaries, confidence, backend, duration, and cache status.
- `embedded_media`: extracted packaged media records with package member, extracted path, hash, content type, cache status, and optional EMF PNG preview path.
- `completeness_score`: conservative extraction coverage score. It combines text coverage, table coverage, verified visual coverage, OCR coverage, OpenAI vision use, unverified visual count, and score signals.
- `artifacts`: paths to generated `.full.md`, `.manifest.json`, and report files.

## Revision Markers In Markdown

- Insertions are rendered as `{+inserted text+}`.
- Deletions are rendered as `{-deleted text-}`.
- Moved-from text is rendered as `{~moved from: text~}`.
- Moved-to text is rendered as `{~moved to: text~}`.

## Visual Review Flag

`requires_visual_review: true` means OOXML text extraction saw media or drawings. In `balanced` and `complete` modes, the visual pipeline tries to add these optional fields:

- `page_index` or `slide_index`: rendered location when known.
- `image_path`: rendered page/slide image used for OCR/vision.
- `ocr_text`: text recovered by local OCR.
- `vision_summary`: image/chart/screenshot interpretation from OpenAI vision when enabled.
- `diagram_summary`: extracted structure or conclusion for charts, architecture diagrams, and screenshots.
- `confidence`: `low`, `medium`, or `high`.
- `backend`: `rapidocr`, `tesseract`, `openai:<model>`, `ooxml-media`, or a combination.
- `duration_ms`: elapsed time for this item.
- `cache_hit`: whether this result came from `.office-reader-cache`.

For DOCX media relationships found in the body, a non-body Word part, a table cell, a block-level content control, or a textbox, entries under `visual_findings[].relationships` may include `part_type`, `part`, `container`, `table_index`, `row_index`, `cell_index`, `paragraph_index`, `paragraph_text`, `nearest_heading`, `nearby_text_before`, `nearby_text_after`, and detected `caption`.

`embedded_media[]` records are generated when package members under `word/media` or `ppt/media` are referenced. Fields:

- `member`: OOXML package member such as `word/media/image1.emf`.
- `path`: extracted local copy under `embedded_media/`.
- `sha256`: package member hash for deduplication and cache identity.
- `content_type`: file extension-derived type such as `png`, `jpeg`, `emf`, or `mp4`.
- `cache_hit`: whether the extracted copy already existed.
- `preview_path` and `preview_format`: optional cached PNG preview for EMF files when conversion succeeds.

For PPTX slides, entries under `visual_findings[].objects` may include a structured visual object inventory:

- `object_type`: `image`, `chart`, `smartart`, `ole`, `video`, `audio`, or another visual risk type.
- `slide_index`: slide location.
- `name`, `alt_text`, and `title`: non-visual DrawingML metadata when present.
- `relationship_id`, `relationship_type`, `target`, and `target_mode`: relationship identifier, relationship type, resolved package target, and external/internal target mode when present.
- `geometry`: EMU coordinates with `x`, `y`, `cx`, and `cy` when the object has an `a:xfrm`.
- `relationships`: SmartArt relationship list with role names such as `data_model`, `layout`, `quick_style`, and `colors`.
- `prog_id`: OLE program identifier such as `Excel.Sheet.12` when present.

The report repeats this inventory under `Visual Findings` so a caller can see which slide contains image/chart/diagram/embed/media risk before opening the JSON manifest.

If these fields are empty, the report must list the visual gap instead of claiming the image content was fully read.

## Completeness Score

`completeness_score` is a conservative coverage signal, not a guarantee that every fact was understood. Visual items count as verified only when local OCR (`rapidocr` or `tesseract`) produced text or OpenAI vision produced a summary. OOXML media hints, fast-mode placeholders, and rendered-page bookkeeping do not count as visual interpretation.

Fields:

- `overall`: 0-100 weighted score.
- `text_coverage`: readable structure entries divided by extracted structure entries.
- `table_coverage`: tables with readable cells divided by extracted tables. If no tables exist, this is 100.
- `visual_coverage`: visually verified items divided by items requiring visual review.
- `ocr_confidence`: OCR-backed items divided by items requiring visual review.
- `openai_vision_enabled`: whether an OpenAI vision backend contributed to any visual finding.
- `unverified_visual_count`: visual-review items still lacking OCR or OpenAI vision confirmation.
- `signals`: short reasons that explain score limits and remaining risks.

## Visual Analysis Messages

`visual_analysis.messages` carries dependency and backend diagnostics that should be surfaced in reports. Preview backend health decisions, such as skipping unhealthy Word or PowerPoint COM preview and falling back to LibreOffice, are recorded here.
