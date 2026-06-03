# Output Schema

`office-reader` writes a manifest JSON file beside the Markdown transcript and report.

The unified `scripts/read_office.py` command also prints a small JSON object to stdout with `full_markdown`, `manifest`, and `report` paths. When comments or tracked revisions are extracted, stdout also includes `review_items` and `review_items_csv`. When `--query` is used, stdout also includes `query_results`. Stdout/stderr are configured as UTF-8 so callers can decode output paths reliably when paths contain Chinese characters or spaces. PowerShell 5 callers should read generated JSON files with `-Encoding UTF8`.

## Top-Level Fields

- `source`: original file path and name.
- `normalized_file`: actual `.docx` or `.pptx` file read after legacy conversion.
- `conversion`: conversion status, backend, output path, and messages. For legacy `.doc/.ppt` text fallback, `status` is `text_fallback` and `output_path` points to the extracted `.legacy-text.txt` file.
- `document_type`: `docx` or `pptx`.
- `reading_mode`: `fast`, `balanced`, or `complete`.
- `metadata`: package metadata from `docProps/core.xml` when present.
- `structure`: paragraphs/headings for Word or ordered slides for PowerPoint. Word entries may include `part_type` and `part` for non-body sources such as headers, footers, footnotes, and endnotes, and `container` values such as `content_control` or `textbox` when text came from those containers.
- `tables`: extracted table rows with document, slide, or Word part location. DOCX tables may also include `caption`, `headers`, `nearby_text_before`, `nearby_text_after`, and `merged_cells`.
- `comments`: Word comments or PowerPoint comments. Word comments may include `anchor_text` when `commentRangeStart/commentRangeEnd` identifies the commented span. Word comments referenced inside a table cell may include `table_index`, `row_index`, `cell_index`, `part_type`, and `part`. PowerPoint comments may include `author_id`, `author`, and `initials` when `ppt/commentAuthors.xml` is present.
- `revisions`: Word tracked insertions and deletions. Word revisions inside a table cell may include `table_index`, `row_index`, `cell_index`, `part_type`, and `part`.
- `notes`: PowerPoint speaker notes.
- `visual_analysis`: visual pipeline status, selected mode, rendered page count, analyzed item count, cache hits, backends, and messages.
- `visual_findings`: flags for media, drawings, image-heavy content, OCR text, rendered-page observations, vision summaries, diagram summaries, confidence, backend, duration, and cache status.
- `embedded_media`: extracted packaged media records with package member, extracted path, hash, content type, cache status, derived label, optional EMF PNG preview path, and optional context records.
- `query`: present when `--query` is used. Contains the query string, normalized tokens, total match count, returned match excerpts, truncation flag, source type, and location metadata.
- `completeness_score`: conservative extraction coverage score. It combines text coverage, table coverage, verified visual coverage, OCR coverage, OpenAI vision use, unverified visual count, and score signals.
- `artifacts`: paths to generated `.full.md`, `.manifest.json`, report, optional review-item, query, media, and preview files.

## Review Items

When comments or tracked revisions are extracted, `read_office.py` writes `<basename>.review-items.json` plus UTF-8-BOM `<basename>.review-items.csv`, records the paths under `manifest.artifacts.review_items` and `manifest.artifacts.review_items_csv`, includes `review_items` and `review_items_csv` in stdout, and lists both in the report artifacts.

`<basename>.review-items.json` fields:

- `source`: original source path/name copied from the manifest.
- `document_type`: manifest document type.
- `total_items`: total comments and revisions exported.
- `counts`: comment and revision counts.
- `items`: flat review queue. Comment items include `kind: "comment"`, `comment_id`, `author`, `initials`, `date`, `text`, `anchor_text`, `location`, and `status: "open"`. Revision items include `kind: "revision"`, `revision_type`, `author`, `date`, `text`, `location`, and `status: "pending"`.

`location` preserves available source coordinates such as paragraph, slide, table row/cell, Word part, and container.

`<basename>.review-items.csv` has fixed spreadsheet-friendly columns:

- `id`, `kind`, `status`, `comment_id`, `revision_type`, `author`, `initials`, `date`, `text`, `anchor_text`
- `paragraph_index`, `slide_index`, `table_index`, `row_index`, `cell_index`, `part_type`, `part`, `container`

## Query Results

When `read_office.py --query "<text>"` is provided, the reader writes `<basename>.query.json`, records the same data under `manifest.query`, adds `artifacts.query_results`, includes `query_results` in stdout, and adds a `Query Results` section to the report.

The query scan covers extracted structure text, tables, comments, comment anchors, revisions, speaker notes, OCR/vision text fields, media relationship context, and embedded-media labels. Matching is case-insensitive and requires all whitespace-separated query tokens to appear in a candidate text field.

`<basename>.query.json` fields:

- `query`: original query text.
- `tokens`: normalized case-folded tokens.
- `total_matches`: total matched candidates before limit truncation.
- `matches`: returned match excerpts with `source_type`, `location`, and `text`.
- `truncated`: whether additional matches were omitted by `--query-limit`.

Query results are lookup aids over already extracted text. They do not prove that unverified image-only content was OCR-read or visually understood.

## Evidence Report

When `read_office.py --evidence-report` or `assemble_report.py --evidence` is used, the report includes an `Evidence Index` section. The section lists source-backed manifest entries for structure, tables, comments, revisions, speaker notes, media relationships, visual object inventory records, and OCR findings. Each entry includes a compact location such as paragraph index, slide index, table row/cell, Word part, media target, or object relationship metadata when available.

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

For DOCX media relationships found in the body, a non-body Word part, a table cell, a block-level content control, or a textbox, entries under `visual_findings[].relationships` may include `part_type`, `part`, `container`, `table_index`, `row_index`, `cell_index`, `paragraph_index`, `paragraph_text`, `nearest_heading`, `nearby_text_before`, `nearby_text_after`, `media_source`, detected `caption`, `object_id`, `name`, `alt_text`, `title`, and `geometry`. `media_source` is usually `drawingml` for `a:blip` images or `vml` for older `v:imagedata`/OLE-style images such as embedded Visio EMF previews. DrawingML `geometry` records inline/anchor extent as EMU `cx`/`cy` when available; VML metadata may include shape ids and titles.

DOCX table captions are matched from the same cell, same row, or nearest preceding caption row within the same table. This is intended for common Word layout tables where an image paragraph is separate from its figure-caption paragraph.

DOCX table semantic fields:

- `caption`: nearest preceding table caption paragraph such as `Table 2-1 ...` or Chinese `表...` when detected.
- `headers`: inferred header cells, usually the first row, or the second row when the first row is a single merged title row.
- `nearby_text_before` and `nearby_text_after`: closest surrounding structure text.
- `merged_cells`: row/cell coordinates for `w:gridSpan` and `w:vMerge` cells, including `grid_span` or `v_merge` where present.

`embedded_media[]` records are generated when package members under `word/media` or `ppt/media` are referenced. Fields:

- `member`: OOXML package member such as `word/media/image1.emf`.
- `path`: extracted local copy under `embedded_media/`.
- `sha256`: package member hash for deduplication and cache identity.
- `content_type`: file extension-derived type such as `png`, `jpeg`, `emf`, or `mp4`.
- `cache_hit`: whether the extracted copy already existed.
- `label`: best available short label derived from caption, alt text, title, object name, nearest heading, or package member.
- `preview_path` and `preview_format`: optional cached PNG preview for EMF files when conversion succeeds.
- `contexts`: optional relationship/object context copied from `visual_findings`, such as caption, nearest heading, table cell, slide, alt text, or DOCX `media_source`.
- `ocr_text` and `ocr_backend`: optional fields populated when `--media-ocr selected` or `--media-ocr all` recovers text from the extracted image or EMF preview.

When image-like media can be opened or previewed, the visual pipeline also writes:

- `media_summary.json`: compact list of media items with paths, preview paths, labels, hashes, OCR text/backend when present, and contexts.
- `media_contact_sheet.jpg`: tiled thumbnail sheet for quick human inspection.

These files are triage aids. They do not prove OCR or semantic understanding unless corresponding `ocr_text`, OpenAI `vision_summary`, or `diagram_summary` fields are present in `visual_findings`. Media OCR adds `source_type: embedded_media_ocr` entries to `visual_findings[]` and increments `visual_analysis.media_ocr_count`.

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
