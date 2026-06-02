# Output Schema

`office-reader` writes a manifest JSON file beside the Markdown transcript and report.

The unified `scripts/read_office.py` command also prints a small JSON object to stdout with `full_markdown`, `manifest`, and `report` paths. That stdout JSON is ASCII escaped so callers can decode it as UTF-8 across Windows consoles even when paths contain Chinese characters or spaces.

## Top-Level Fields

- `source`: original file path and name.
- `normalized_file`: actual `.docx` or `.pptx` file read after legacy conversion.
- `conversion`: conversion status, backend, output path, and messages.
- `document_type`: `docx` or `pptx`.
- `reading_mode`: `fast`, `balanced`, or `complete`.
- `metadata`: package metadata from `docProps/core.xml` when present.
- `structure`: paragraphs/headings for Word or ordered slides for PowerPoint.
- `tables`: extracted table rows with document or slide location.
- `comments`: Word comments or PowerPoint comments.
- `revisions`: Word tracked insertions and deletions.
- `notes`: PowerPoint speaker notes.
- `visual_analysis`: visual pipeline status, selected mode, rendered page count, analyzed item count, cache hits, backends, and messages.
- `visual_findings`: flags for media, drawings, image-heavy content, OCR text, rendered-page observations, vision summaries, diagram summaries, confidence, backend, duration, and cache status.
- `artifacts`: paths to generated `.full.md`, `.manifest.json`, and report files.

## Revision Markers In Markdown

- Insertions are rendered as `{+inserted text+}`.
- Deletions are rendered as `{-deleted text-}`.

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

If these fields are empty, the report must list the visual gap instead of claiming the image content was fully read.

## Visual Analysis Messages

`visual_analysis.messages` carries dependency and backend diagnostics that should be surfaced in reports. Preview backend health decisions, such as skipping unhealthy Word or PowerPoint COM preview and falling back to LibreOffice, are recorded here.
