# Output Schema

`office-reader` writes a manifest JSON file beside the Markdown transcript and report.

## Top-Level Fields

- `source`: original file path and name.
- `normalized_file`: actual `.docx` or `.pptx` file read after legacy conversion.
- `conversion`: conversion status, backend, output path, and messages.
- `document_type`: `docx` or `pptx`.
- `reading_mode`: `fast`, `balanced`, or `complete`.
- `metadata`: package metadata from `docProps/core.xml` when present.
- `structure`: paragraphs/headings for Word or ordered slides for PowerPoint.
- `tables`: extracted table rows with document or slide location.
- `comments`: Word comments or PowerPoint comments. PowerPoint comments include `author_id` and the mapped `author` from `ppt/commentAuthors.xml` when available. PowerPoint package comments without a slide relationship use `slide_index: null`.
- `revisions`: Word tracked insertions and deletions.
- `notes`: PowerPoint speaker notes.
- `warnings`: non-fatal extraction warnings, such as malformed optional OOXML parts that were skipped while preserving readable body content.
- `visual_analysis`: visual pipeline status, selected mode, rendered page count, analyzed item count, cache hits, backends, and messages.
- `visual_findings`: flags for media, drawings, image-heavy content, OCR text, rendered-page observations, vision summaries, diagram summaries, confidence, backend, duration, and cache status.
- `artifacts`: paths to generated `.full.md`, `.manifest.json`, and report files.

## Revision Markers In Markdown

- Insertions are rendered as `{+inserted text+}`.
- Deletions are rendered as `{-deleted text-}`.
- Word `<w:tab/>`, `<w:br/>`, and `<w:cr/>` separators are preserved as tabs or line breaks during text extraction.
- PowerPoint slide body text is aggregated by shape paragraph so formatted runs remain joined. Table cell text stays in `tables` rather than being duplicated into slide body text.
- PowerPoint speaker notes prefer `p:ph type="body"` placeholders so slide-number and footer placeholders do not pollute note text. Packages without a body placeholder fall back to all note text.

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
- `origin`: generated page-level findings use `rendered-page`.

Visual cache keys include a schema version, reading mode, rendered item hash, and analysis profile. Local-only runs and OpenAI Vision runs do not reuse each other's cache entries; OpenAI Vision cache entries are also separated by model name.

Unreadable visual cache JSON is ignored and recomputed. Fresh results replace the damaged entry atomically.

On a cache hit, analysis text is reused but run-specific location fields such as `image_path`, `page_index`, and `slide_index` are refreshed from the current rendered page.

On enrichment reruns, prior `origin: "rendered-page"` findings are replaced before the current rendered pages are analyzed. They do not accumulate across runs.

Rendered page images are written under run-specific `page_images/render-<guid>` directories. Reruns preserve existing page image files and update `artifacts.page_images` to the current run.

DOCX/PPTX extraction reuses existing artifact paths only when the prior manifest belongs to the same source and document type and its recorded Markdown/manifest artifact paths exactly match the candidate outputs. Unknown or incomplete same-name artifacts are preserved; new artifacts are written under `office-reader-run-<guid>`.

If these fields are empty, the report must list the visual gap instead of claiming the image content was fully read.
