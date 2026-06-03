# Office Reader Iteration Log

## 2026-06-03 - Legacy conversion COM timeout and health memory

### Problems Found

- Legacy `.doc`/`.ppt` conversion used Office COM directly and lacked preview-style timeout isolation.
- A slow or stuck Office COM conversion could block the conversion path before WPS/LibreOffice fallback.
- Legacy conversion did not remember unhealthy Office COM behavior, so later runs could repeat the same slow failure.
- Conversion health diagnostics had no path-safety guard comparable to preview health memory.

### Changes Completed

- Added `convert_legacy_office.ps1 -TimeoutSeconds` and `-ContinueAfterComFailure`.
- Added isolated Office COM conversion worker with timeout handling.
- Added `.office-reader-cache/conversion-backend-health.json` health memory, overridable with `OFFICE_READER_CONVERSION_HEALTH_PATH`.
- Added path safety validation so conversion health memory cannot overwrite arbitrary JSON/non-JSON files.
- Auto-priority conversion skips Office COM when health memory marks it unhealthy for `.doc` or `.ppt`, then tries WPS/LibreOffice fallback backends.
- `read_office.py` now forwards `--visual-timeout-seconds` to legacy conversion and enables fallback continuation for the built-in converter.
- Updated `SKILL.md`, `README.md`, and `references/backend_fallbacks.md`.

### Verification

- TDD red tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_skips_unhealthy_office_com_from_health_memory tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_rejects_health_path_that_would_overwrite_regular_file -v`
  - Initial result: health skip failed; invalid health path initially risked slow backend probing.
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_office_com_timeout_records_health_and_falls_back -v`
  - Initial result: failed because Office COM conversion did not run through a timeout-isolated worker.
- Focused tests after implementation:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_refuses_to_overwrite_existing_normalized_file tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_cleans_libreoffice_profile_directory tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_skips_unhealthy_office_com_from_health_memory tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_rejects_health_path_that_would_overwrite_regular_file tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_office_com_timeout_records_health_and_falls_back tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_libreoffice_timeout_is_structured -v`
  - Result: `OK`
- Legacy fallback JSON safety:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_legacy_text_fallback_manifest_is_powershell_json_safe -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\read_office.py tests\test_office_reader.py`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 47 tests
- Real legacy DOC smoke:
  - Source: `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.doc`
  - Command: `read_office.py ... --mode balanced --visual-timeout-seconds 5 --query "HT32" --no-openai-vision`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\legacy-timeout-round-20260603-ht32-doc-v2`
  - Result: success
  - Conversion status: `text_fallback`
  - Conversion backend: `word-com-text`
  - Query matches: `16`
  - Conversion messages included Office COM health-memory skip, LibreOffice 5-second timeout, and Word COM text fallback success.
  - Manifest re-read checks:
    - Python `json.loads`: `OK`
    - PowerShell `Get-Content -Encoding UTF8 | ConvertFrom-Json`: `OK`

### Remaining Risks

- Worker timeout protects the conversion wrapper, but Office/WPS COM automation can still leave external Office processes that require OS-level cleanup in rare cases.
- LibreOffice fallback is now timeout-bounded, but a timed-out conversion may produce no normalized `.docx/.pptx`; balanced/complete mode can still fall back to searchable legacy text when available.
- Health memory is intentionally conservative; users can clear `.office-reader-cache/conversion-backend-health.json` after fixing Office installation problems.

### Next Round Direction

- Add review-item export for comments and tracked revisions.

## 2026-06-03 - DOCX table semantic hints

### Problems Found

- DOCX tables only exposed raw rows and location metadata.
- Reports could list table columns, but did not expose table captions, inferred headers, nearby explanatory context, or merged-cell signals.
- Query mode could search raw table cells, but not table captions or surrounding context as separate evidence-bearing fields.

### Changes Completed

- Added DOCX table semantic extraction:
  - `caption` from nearest preceding table-caption paragraph.
  - `headers` inferred from the first row, or second row when the first row is a single merged title row.
  - `nearby_text_before` and `nearby_text_after`.
  - `merged_cells` for `w:gridSpan` and `w:vMerge`.
- Updated Table Index and Evidence Index report output to show captions, headers, context, and merged-cell counts.
- Added query candidates for table caption, headers, and nearby context.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- TDD red test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_table_semantics -v`
  - Initial result: failed with missing `caption` metadata.
- Focused tests after implementation:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_table_semantics tests.test_office_reader.OfficeReaderTests.test_report_assembler_evidence_mode_lists_source_locations tests.test_office_reader.OfficeReaderTests.test_unified_reader_query_writes_matches_to_manifest_report_and_stdout -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\read_docx.py scripts\assemble_report.py scripts\read_office.py tests\test_office_reader.py`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 42 tests
- Real DOCX table semantics smoke:
  - Source: `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.docx`
  - Command: `read_office.py ... --mode fast --query "表" --evidence-report --no-openai-vision`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\table-semantics-round-20260603-ht32`
  - Result: success
  - Tables: `7`
  - Tables with headers: `7`
  - Tables with before/after context: `7/7`
  - Tables with merged-cell signals: `4`
  - Query matches: `24`
  - Note: this real document did not expose table-caption paragraphs matching the caption heuristic; synthetic regression covers the table-caption path.

### Remaining Risks

- Header inference is heuristic and may be wrong for unusual multi-level table headers.
- Vertical merge continuation text is not expanded across rows; `merged_cells` records the signal for downstream review.

### Next Round Direction

- Add legacy `.doc/.ppt` conversion timeout isolation and backend health memory, mirroring preview rendering protections.

## 2026-06-03 - Evidence index report mode

### Problems Found

- Structured reports summarized extracted content but did not provide a compact source-backed evidence index.
- Callers reviewing summaries had to jump between manifest JSON sections to locate paragraph, table, comment, revision, media, visual object, and OCR evidence.
- The unified reader had no single flag to request a source-location-heavy report.

### Changes Completed

- Added `assemble_report.py --evidence`.
- Added `read_office.py --evidence-report` forwarding through the unified entrypoint.
- Evidence Index includes structure, tables, comments, revisions, speaker notes, DOCX/PPTX media relationships, visual object inventory records, and OCR findings.
- Evidence lines include compact locations such as paragraph index, slide index, table row/cell, Word part, media target, and object metadata when available.
- Updated `SKILL.md`, `README.md`, and `references/output_schema.md`.

### Verification

- TDD red test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_report_assembler_evidence_mode_lists_source_locations -v`
  - Initial result: failed because `assemble_report.py` did not recognize `--evidence`.
- Focused tests after implementation:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_report_assembler_evidence_mode_lists_source_locations tests.test_office_reader.OfficeReaderTests.test_unified_reader_evidence_report_forwards_to_report_assembler -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\assemble_report.py scripts\read_office.py tests\test_office_reader.py`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 41 tests
- Real DOCX evidence smoke:
  - Source: `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.docx`
  - Command: `read_office.py ... --mode fast --query "HT32" --evidence-report --no-openai-vision`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\evidence-round-20260603-ht32`
  - Result: success
  - Evidence Index present: `yes`
  - Evidence lines: `128`
  - Sample evidence line: `Structure p8 ... 参赛编号:20260034 使用合泰芯片型号：HT32F52352`

### Remaining Risks

- Evidence Index is an index of extracted evidence, not an argument checker or citation verifier.
- It lists compact excerpts and locations; users still need the full Markdown/manifest for long-form traceability.

### Next Round Direction

- Add stronger table semantics: captions, header detection, merged-cell signals, and nearby explanatory context.

## 2026-06-03 - DOCX image object metadata inventory

### Problems Found

- DOCX media relationships carried useful location and caption context, but modern DrawingML object metadata was not exposed.
- Word image alt text, object name/title, object id, and inline/anchor extent were invisible to manifest consumers.
- VML/EMF preview relationships exposed `media_source: vml`, but not shape id/title metadata.

### Changes Completed

- Added DrawingML image metadata extraction from `wp:inline` / `wp:anchor`:
  - `object_id`
  - `name`
  - `alt_text`
  - `title`
  - `geometry` with EMU `cx`/`cy` extent when available.
- Added fallback picture metadata extraction from `pic:cNvPr`.
- Added VML shape metadata extraction for object id/name/title when available.
- Propagated new context fields into embedded media contexts, media summary labels, query candidates, and report media context lines.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- TDD red test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_drawingml_image_metadata -v`
  - Initial result: failed with missing `name` metadata.
- Focused tests after implementation:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_vml_imagedata_media_relationships tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_drawingml_image_metadata -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\read_docx.py scripts\visual_analysis.py scripts\read_office.py scripts\assemble_report.py scripts\common_ooxml.py tests\test_office_reader.py`
  - Result: `OK`
- Focused regression set:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_vml_imagedata_media_relationships tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_drawingml_image_metadata tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_extracts_embedded_media_and_report_lists_context -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 39 tests
- Real DOCX metadata smoke:
  - Source: `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.docx`
  - Command: `read_office.py ... --mode fast --query "HT32" --no-openai-vision`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\docx-metadata-round-20260603-ht32`
  - Result: success
  - DOCX media relationships: `6`
  - Relationships with `name`: `6`
  - Relationships with `geometry`: `6`
  - Query matches: `23`
  - Embedded media extracted: `6`

### Remaining Risks

- DOCX shape crop, rotation, wrap mode, z-order, and absolute position are not modeled yet.
- Complex embedded OLE/Visio object internals are still represented through preview media metadata, not parsed semantically.

### Next Round Direction

- Add evidence-style report mode that ties summaries/findings to paragraph, table, slide, media, comment, and revision locations.

## 2026-06-03 - Fast embedded media OCR option

### Problems Found

- Fast mode extracted embedded images and EMF previews but did not OCR them.
- Query mode could search OCR fields, but fast-mode image text remained invisible unless full page rendering/OCR was used.
- `media_summary.json` and the report did not have a place to expose text recovered from individual embedded media files.

### Changes Completed

- Added `visual_analysis.py --media-ocr off|selected|all`.
- Added `read_office.py --media-ocr off|selected|all` forwarding through the unified entrypoint.
- `selected` OCRs a bounded set of extracted image-like media items/EMF previews; `all` attempts every image-like extracted media item.
- Media OCR writes `ocr_text` and `ocr_backend` to `embedded_media[]` and `media_summary.json`.
- Added `source_type: embedded_media_ocr` visual findings so query mode can find media OCR text.
- Reports now show `Media OCR` lines under Embedded Media.
- Updated `SKILL.md`, `README.md`, and `references/output_schema.md`.

### Verification

- TDD red test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_media_ocr_adds_text_to_embedded_media -v`
  - Initial result: failed because `visual_analysis.py` did not recognize `--media-ocr`.
- Focused test after implementation:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_media_ocr_adds_text_to_embedded_media -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\visual_analysis.py scripts\read_office.py scripts\assemble_report.py tests\test_office_reader.py`
  - Result: `OK`
- Focused regression pair:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_media_ocr_adds_text_to_embedded_media tests.test_office_reader.OfficeReaderTests.test_unified_reader_query_writes_matches_to_manifest_report_and_stdout -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 38 tests
- Real PowerPoint media OCR smoke:
  - Source: `C:\Users\Huang\Documents\标题幻灯片选项 1.pptx`
  - Command: `read_office.py ... --mode fast --media-ocr selected --query "标题" --no-openai-vision`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\media-ocr-round-20260603-title-pptx-v2`
  - Result: success
  - Visual analysis status: `partial`
  - Embedded media: `29`
  - Media OCR count: `6`
  - Query matches: `4`
  - Report includes `Media OCR via tesseract` / `Media OCR via rapidocr` lines.

### Remaining Risks

- OCR quality depends on RapidOCR/Tesseract availability and image quality.
- `selected` uses a bounded item count to keep fast mode fast; use `all` for exhaustive embedded-media OCR.
- Media OCR does not interpret diagrams semantically unless OCR text is enough; OpenAI vision or complete mode is still needed for complex chart/diagram meaning.

### Next Round Direction

- Add stronger DOCX image object inventory metadata, especially alt text, shape names, dimensions, and OLE-style object hints.

## 2026-06-03 - Unified reader query lookup artifact

### Problems Found

- `--mode fast` was useful for simple lookups, but callers still had to open the Markdown/report and manually search for relevant excerpts.
- The unified stdout JSON did not expose a focused lookup artifact for automation.
- Reports had no dedicated section showing query hits and their manifest locations.

### Changes Completed

- Added `read_office.py --query "<text>"` and `--query-limit`.
- Added `<basename>.query.json` with normalized tokens, match count, source type, location metadata, excerpts, and truncation flag.
- Added `manifest.query`, `artifacts.query_results`, and stdout `query_results`.
- Query scan covers extracted structure text, tables, comments, comment anchors, revisions, speaker notes, OCR/vision text fields, media relationship context, and embedded-media labels.
- Added a Query Results section to `assemble_report.py`.
- Updated `SKILL.md`, `README.md`, and `references/output_schema.md`.

### Verification

- TDD red test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_unified_reader_query_writes_matches_to_manifest_report_and_stdout -v`
  - Initial result: failed because `read_office.py` did not recognize `--query`.
- Focused test after implementation:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_unified_reader_query_writes_matches_to_manifest_report_and_stdout -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\read_office.py scripts\assemble_report.py tests\test_office_reader.py`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 37 tests
- Real PowerPoint query smoke:
  - Source: `C:\Users\Huang\Documents\标题幻灯片选项 1.pptx`
  - Command: `read_office.py ... --mode fast --query "标题" --no-openai-vision`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\query-round-20260603-title-pptx`
  - Result: success
  - Query matches: `4`
  - Report includes `## Query Results` and slide hits such as `标题幻灯片选项 1`.

### Remaining Risks

- Query mode searches already extracted text fields. It does not OCR unverified image-only content by itself.
- Matching is intentionally simple all-token substring matching; fuzzy matching and ranking are not implemented yet.

### Next Round Direction

- Add lightweight media-level OCR for selected extracted images so fast visual lookups can cover common screenshot/diagram text without rendering every page.

## 2026-06-03 - Embedded media manifest labels

### Problems Found

- In the previous three-document smoke, `media_summary.json` had useful derived labels, but `manifest.embedded_media[]` did not expose the same top-level `label`.
- Downstream scripts that only inspect the manifest had to duplicate label derivation from contexts or incorrectly report zero labeled media.

### Changes Completed

- Added top-level `label` to each `embedded_media[]` record after media contexts are attached.
- Kept `media_summary.json` labels and contact-sheet labels derived from the same source.
- Updated the report assembler to prefer `embedded_media[].label` before falling back to context fields.
- Added a regression assertion that manifest and media-summary labels are present for extracted DOCX media.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- Syntax check:
  - `python -m py_compile scripts\visual_analysis.py scripts\assemble_report.py tests\test_office_reader.py`
  - Result: `OK`
- Focused test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_extracts_embedded_media_and_report_lists_context -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 36 tests
- Real three-document fast slice:
  - Huang Man: `27` media, `27` manifest labels, `27` summary labels, `0` mismatches, `8` VML/EMF items.
  - HT32-CC: `6` media, `6` manifest labels, `6` summary labels, `0` mismatches.
  - Huang Dengke final: `3` media, `3` manifest labels, `3` summary labels, `0` mismatches.
  - Output root: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\label-round-20260603-105816`.

### Remaining Risks

- Labels are contextual identifiers, not proof that the image content itself was OCR-read or semantically understood.
- Image/diagram text still requires balanced or complete visual analysis with OCR/vision where appropriate.

### Next Round Direction

- Continue auditing real-document slices for non-invented gaps, especially selected image/OCR behavior or lightweight visual block interpretation.

## 2026-06-03 - DOCX VML/EMF media relationship context

### Problems Found

- The Huang Man real DOCX contained 8 EMF/Visio-style flowcharts under `word/media`, but they were referenced through VML `v:imagedata` instead of DrawingML `a:blip`.
- The DOCX reader only scanned DrawingML image relationships, so those VML images could be extracted as packaged media but lacked relationship ids, captions, nearby text, and source metadata.
- `media_summary.json` could not distinguish modern DrawingML images from older VML/OLE-style image previews.

### Changes Completed

- Added DOCX media relationship scanning for both:
  - DrawingML `a:blip` images as `media_source: drawingml`.
  - VML `v:imagedata r:id` images as `media_source: vml`.
- Propagated `media_source` through DOCX relationship contexts and embedded media summaries.
- Added a synthetic DOCX regression fixture for a VML Visio/EMF image with a nearby figure caption.
- Updated `SKILL.md` and `references/output_schema.md` to document VML media-source context.

### Verification

- Syntax check:
  - `python -m py_compile scripts\read_docx.py scripts\visual_analysis.py tests\test_office_reader.py`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 36 tests
- Real three-document fast slice:
  - Huang Man: `27` media summary items, `27` labeled, `8` VML/EMF items detected.
  - HT32-CC: `6` media summary items, `6` labeled.
  - Huang Dengke final: `3` media summary items, `3` labeled.
  - Output root: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\vml-round-20260603-105201`.

### Remaining Risks

- VML shape geometry, OLE object metadata, and internal Visio semantics are still not modeled deeply.
- Fast mode extracts and labels visual media but still does not OCR or semantically understand the image content.
- Media labels are written to `media_summary.json`; `manifest.embedded_media[]` carries contexts but not a top-level `label` field.

### Next Round Direction

- Add an optional lightweight media OCR/vision block pass for selected extracted images, or add top-level embedded-media labels in the manifest for easier downstream counting.

## 2026-06-03 - Legacy fast text fallback for simple lookup tasks

### Problems Found

- Simple lookup questions on legacy `.doc` files could spend minutes attempting full conversion before falling back to plain text manually.
- `read_office.py` treated legacy conversion failure as a hard failure, even when text-only extraction would answer narrow tasks such as "find experiment eight thinking questions."
- `.ppt` legacy fast lookup had the same risk pattern.

### Changes Completed

- Added `scripts/extract_legacy_text.ps1`:
  - `.doc`: read-only Word COM text extraction.
  - `.ppt`: read-only PowerPoint COM slide/notes text extraction.
  - structured JSON success/failure output.
- Added `read_office.py --mode fast` legacy behavior:
  - `.doc` and `.ppt` go directly to text fallback instead of full conversion.
  - produces normal `.full.md`, `.manifest.json`, and `.report.md` artifacts.
- Added `balanced`/`complete` fallback behavior:
  - if full legacy conversion fails, `read_office.py` tries text fallback before returning failure.
  - manifest records `conversion.status: text_fallback`.
  - completeness score stays conservative and warns that layout/media/tables may be incomplete.
- Added BOM stripping for PowerShell UTF-8 text output.
- Added test override environment variables for deterministic conversion/text fallback tests:
  - `OFFICE_READER_CONVERT_LEGACY_SCRIPT`
  - `OFFICE_READER_LEGACY_TEXT_EXTRACTOR`
- Updated `SKILL.md`, `README.md`, `references/output_schema.md`, and `references/backend_fallbacks.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_unified_reader_fast_mode_uses_legacy_text_fallback_for_doc tests.test_office_reader.OfficeReaderTests.test_unified_reader_fast_mode_uses_legacy_text_fallback_for_ppt tests.test_office_reader.OfficeReaderTests.test_unified_reader_falls_back_to_legacy_text_when_conversion_fails -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 31 tests

### Remaining Risks

- Text fallback still depends on Office COM availability for legacy `.doc/.ppt`.
- Text fallback is not a substitute for full deep reading: tables, comments, revisions, media, and layout can be incomplete.
- No explicit natural-language query argument exists yet; callers still choose `--mode fast` for narrow lookup tasks.

### Next Round Direction

- Add a lightweight query/excerpt helper or continue with DOCX drawing/image metadata inventory.

## 2026-06-02 - DOCX move revision extraction

### Problems Found

- `w:moveFrom` and `w:moveTo` were not modeled as revisions.
- Moved text could be flattened into paragraph text without making the review action clear.

### Changes Completed

- Added `move_from` and `move_to` revision extraction in `paragraph_text_with_revisions()`.
- Added explicit Markdown markers:
  - `{~moved from: text~}`
  - `{~moved to: text~}`
- Added a synthetic DOCX regression fixture for moved-from and moved-to text.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_move_revisions tests.test_office_reader.OfficeReaderTests.test_docx_reader_outputs_markdown_and_manifest -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 28 tests

### Remaining Risks

- DOCX drawing geometry/alt text for images is still less detailed than PPTX visual object inventory.
- Formatting-only revisions are not yet modeled.
- COM worker stdout/stderr temp files are still retained in some paths.

### Next Round Direction

- Add DOCX drawing/image metadata inventory or clean COM worker stdout/stderr files.

## 2026-06-02 - DOCX comment range anchors

### Problems Found

- Word comments were associated with `commentReference` ids but did not expose the text span covered by `commentRangeStart/commentRangeEnd`.
- Reports and manifests could say a paragraph had a comment without showing what exact phrase the reviewer commented on.

### Changes Completed

- Added comment range scanning for DOCX paragraphs with optional subtree skipping.
- Added `anchor_text` to Word comment records when a matching range span is present.
- Preserved existing paragraph/table/content-control/textbox location metadata.
- Report comments now show the anchor text when present.
- Added a synthetic DOCX regression fixture covering `commentRangeStart`, covered text, `commentRangeEnd`, and `commentReference`.
- Updated `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_comment_range_anchor_text tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_textbox_content_with_origin tests.test_office_reader.OfficeReaderTests.test_docx_reader_outputs_markdown_and_manifest -v`
  - Result: `OK`
- Report focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_report_assembler_uses_manifest_counts_and_outline tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_comment_range_anchor_text -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 27 tests

### Remaining Risks

- Move revisions (`w:moveFrom`, `w:moveTo`) are not yet modeled separately.
- DOCX drawing geometry/alt text for images is still less detailed than PPTX visual object inventory.

### Next Round Direction

- Add move revision extraction.

## 2026-06-02 - DOCX textbox origin metadata

### Problems Found

- Text inside `w:txbxContent` could be folded into the outer paragraph by recursive text extraction.
- Textbox comments, revisions, and image relationships could lose their true shape/textbox origin.
- Outer paragraphs that hosted a textbox could duplicate textbox text or media in the wrong location.

### Changes Completed

- Added optional subtree skipping to `paragraph_text_with_revisions()` so callers can avoid mixing textbox content into the outer paragraph.
- Added DOCX textbox block extraction from `w:txbxContent`.
- Tagged textbox-originated paragraphs, revisions, comments, and media relationships with `container: textbox`.
- Added subtree-aware comment and image scanning so outer paragraphs do not duplicate textbox references.
- Added a synthetic DOCX regression fixture with outer text, textbox text, insertion revision, comment reference, and textbox image relationship.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_textbox_content_with_origin tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_block_level_content_controls tests.test_office_reader.OfficeReaderTests.test_docx_reader_outputs_markdown_and_manifest -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 26 tests

### Remaining Risks

- Comment range anchors are still not captured as `anchor_text`.
- Move revisions (`w:moveFrom`, `w:moveTo`) are not yet modeled separately.
- DOCX drawing geometry/alt text for images is still less detailed than PPTX visual object inventory.

### Next Round Direction

- Add DOCX comment range anchor text extraction, then model move revisions.

## 2026-06-02 - DOCX block-level content controls

### Problems Found

- Block-level Word content controls (`w:sdt`) were skipped because `read_docx.py` only processed direct paragraph/table children.
- Template/report fields inside `w:sdtContent` could omit paragraphs, tables, comments, revisions, and media references.

### Changes Completed

- Added recursive processing for block-level `w:sdt/w:sdtContent`.
- Tagged extracted paragraphs, tables, comments, revisions, and media relationships from content controls with `container: content_control`.
- Added a synthetic DOCX regression fixture with a content-control paragraph, table, insertion revision, comment reference, and image relationship.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_block_level_content_controls tests.test_office_reader.OfficeReaderTests.test_docx_reader_outputs_markdown_and_manifest -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 25 tests

### Remaining Risks

- Textboxes and shape-originated text still need first-class origin metadata.
- Comment range anchors are still not captured as `anchor_text`.
- Move revisions (`w:moveFrom`, `w:moveTo`) are not yet modeled separately.

### Next Round Direction

- Add textbox-origin detection for DOCX paragraphs/media, then improve comment range anchors.

## 2026-06-02 - DOCX supplemental part extraction

### Problems Found

- `read_docx.py` only processed `word/document.xml` body children.
- Header, footer, footnote, and endnote text could be silently omitted from Markdown and manifest output.
- Media relationships inside header/footer parts were not location-aware.

### Changes Completed

- Added DOCX supplemental part discovery for:
  - headers referenced from `word/_rels/document.xml.rels`
  - footers referenced from `word/_rels/document.xml.rels`
  - `word/footnotes.xml`
  - `word/endnotes.xml`
- Added shared part processing for Word paragraphs, tables, comments, revisions, and drawing media references.
- Added `part_type` and `part` location metadata to Word structure entries, tables, revisions, comments, and media relationship hints.
- Added a synthetic DOCX regression fixture covering header text, footer text, footnote text, endnote text, and a header image relationship.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_outputs_markdown_and_manifest tests.test_office_reader.OfficeReaderTests.test_docx_reader_extracts_headers_footers_footnotes_and_endnotes -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 24 tests

### Remaining Risks

- Block-level content controls (`w:sdt`) are still not fully unpacked as first-class containers.
- Textboxes and shape-originated text can still be flattened or missed, especially outside the main document body.
- Comment range anchors and move revisions are not yet modeled.

### Next Round Direction

- Add DOCX block-level content control extraction and then textbox-origin metadata.

## 2026-06-02 - LibreOffice temp profile cleanup

### Problems Found

- `render_preview.ps1` created `.lo_profile_<guid>` directories in the preview output directory and did not clean them after LibreOffice finished.
- `convert_legacy_office.ps1` created the same temporary LibreOffice profile directories during `.doc/.ppt` fallback conversion and left them behind.
- While testing this cleanup, `render_preview.ps1` also exposed a structured-output reliability bug: `Get-CimInstance Win32_Process` can be denied under managed permissions, causing the script to exit with no JSON stdout instead of returning a structured failure/fallback result.

### Changes Completed

- Added best-effort cleanup of LibreOffice temporary profile directories in preview rendering.
- Added best-effort cleanup of LibreOffice temporary profile directories in legacy conversion.
- Made Office COM automation process cleanup best-effort in `render_preview.ps1`; access-denied diagnostics are added to messages instead of breaking JSON output.
- Added regression coverage:
  - preview LibreOffice fallback leaves no `.lo_profile_*` directories.
  - legacy conversion LibreOffice fallback leaves no `.lo_profile_*` directories.
  - timeout health-memory test continues to receive structured JSON even when process inspection is unavailable.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_render_preview_timeout_records_unhealthy_com_backend tests.test_office_reader.OfficeReaderTests.test_render_preview_com_worker_failure_can_continue_to_later_backends tests.test_office_reader.OfficeReaderTests.test_legacy_conversion_cleans_libreoffice_profile_directory -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 23 tests

### Remaining Risks

- COM worker stdout/stderr temp files are still retained in some paths.
- Legacy `.doc/.ppt` Office COM conversion still lacks the preview-style timeout isolation.
- DOCX headers, footers, footnotes, endnotes, block-level content controls, and textbox-originated content still need extraction and positioning.

### Next Round Direction

- Add DOCX non-body part extraction for headers, footers, footnotes, and endnotes, with synthetic OOXML coverage.

## 2026-06-02 - PPTX visual object inventory v2

### Problems Found

- PPTX inventory v1 covered images and charts, but SmartArt, OLE objects, video, and audio were still easy to miss or collapse into coarse media hints.
- Chart objects did not carry frame-level non-visual metadata or geometry.
- Reports did not expose `visual_findings[].objects`, so a caller had to open the manifest JSON to see object-level visual risk.

### Changes Completed

- Extended `read_pptx.py` visual object inventory:
  - SmartArt detection from diagram `graphicData` and `dgm:relIds`.
  - SmartArt relationship roles for data model, layout, quick style, and colors.
  - OLE detection with `prog_id`, relationship id, relationship type, target, and target mode.
  - Video detection from `a:videoFile`.
  - Audio detection from `a:audioFile` and `a:wavAudioFile`.
  - Relationship type and target mode are now included for PPTX visual objects when available.
  - Chart objects now reuse frame metadata and geometry when available.
- Added object summaries to `assemble_report.py` under `Visual Findings`.
- Added a synthetic PPTX regression fixture for SmartArt, OLE, external video, and embedded audio.
- Updated `SKILL.md`, `README.md`, and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_pptx_reader_inventories_complex_visual_objects tests.test_office_reader.OfficeReaderTests.test_report_assembler_uses_manifest_counts_and_outline -v`
  - Result: `OK`
- Syntax check:
  - `python -m py_compile scripts\read_pptx.py scripts\assemble_report.py`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 22 tests
- Real document smoke:
  - Source: `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.docx`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\ht32-posture-final-cc-v2`
  - Result: success
  - `structure`: 84
  - `tables`: 7
  - `visual_analysis.status`: `completed`
  - `rendered_page_count`: 8
  - `analyzed_item_count`: 8
  - `completeness_score.overall`: 92
  - `unverified_visual_count`: 1

### Remaining Risks

- DOCX headers, footers, footnotes, endnotes, block-level content controls, and textbox-originated content still need deeper extraction and positioning.
- PPTX group transforms, crop rectangles, rotation, z-order, and layout/master inherited objects are still not modeled.
- LibreOffice `.lo_profile_*` temporary directories should be cleaned in preview rendering and legacy conversion.

### Next Round Direction

- Run full test suite and real-document smoke for this round, then commit and push.
- Next implementation round: clean LibreOffice temporary profile directories or add DOCX non-body part extraction, depending on which is lower risk to land first.

## 2026-06-02 - PPTX visual object inventory v1

### Problems Found

- `read_pptx.py` only recorded `a:blip` media relationship hints.
- PPTX image object names, alt text, and EMU geometry were not exposed in the manifest.
- Chart-only visual risk objects did not produce structured visual object records.

### Changes Completed

- Added chart and diagram namespaces to `common_ooxml.py`.
- Added `visual_findings[].objects` for PPTX slides.
- Extracted image object metadata:
  - `name`
  - `alt_text`
  - `title`
  - `relationship_id`
  - resolved package `target`
  - EMU `geometry`
- Extracted chart visual-risk objects with relationship id and resolved target.
- Kept legacy `visual_findings[].media` for compatibility.
- Added a synthetic PPTX regression fixture covering image alt text, image geometry, and chart risk.
- Updated `references/output_schema.md`.

### Verification

- Focused test:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_pptx_reader_extracts_slides_notes_comments_and_tables -v`
  - Result: `OK`

### Remaining Risks

- SmartArt, OLE, video, and audio objects are still not inventoried in v1.
- Group transforms, crop, rotation, z-order, and layout/master inherited objects are not handled yet.

### Next Round Direction

- Extend PPTX object inventory to SmartArt, OLE, video, and audio with minimal synthetic fixtures.

## 2026-06-03 - Three-document slice: DOCX layout-table figure captions

### Problems Found

- User clarified the real slice set should be:
  - `D:\QQ文件\(黄曼)基于STM32多参数水质实时监测系统设计与实现 .docx`
  - `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.docx`
  - `C:\Users\Huang\Desktop\Proj\发电厂课程设计\03_成果输出\最终提交材料\黄登科-202407124102-定稿材料\黄登科-202407124102-定稿.docx`
- The earlier user-provided Huang Dengke path missed the final child directory; the actual `.docx` is inside `黄登科-202407124102-定稿材料`.
- Three-document fast slice found that Huang Man had only `3/27` media items labeled, while HT32-CC and Huang Dengke were fully labeled.
- Root cause: Word layout tables often store image paragraphs separately from caption paragraphs. `read_docx.py` only used paragraph text or a single unique table caption, so multi-image tables produced missing or wrong labels.
- The `is_caption_text()` regex had non-ASCII source text that was vulnerable to encoding display/copy damage.

### Changes Completed

- Rewrote caption detection using Unicode escapes for Chinese `图` and `表`.
- Added table-caption candidate extraction from table rows/cells.
- Added table-aware media caption matching:
  - same cell
  - same row
  - nearest preceding caption row
  - single caption fallback
- Added regression test for multiple figures in one layout table so the second image binds to the second caption instead of the first.
- Updated `SKILL.md` and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_matches_multiple_table_figures_by_nearest_caption_row tests.test_office_reader.OfficeReaderTests.test_docx_reader_uses_table_cell_paragraph_caption_for_media tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_extracts_embedded_media_and_report_lists_context -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 35 tests
- Real three-document fast slice:
  - Huang Man: `27` media, labeled count improved from `3` to `19`.
  - HT32-CC: `6` media, `6` labeled.
  - Huang Dengke final: `3` media, `3` labeled.
  - Warm-cache rerun times: Huang Man about `636 ms`, HT32-CC about `460 ms`, Huang Dengke about `500 ms`.
- Specific Huang Man dual-image table check:
  - `word/media/image22.jpeg` -> `图5-6-1 正常采集与OLED显示测试图`
  - `word/media/image23.png` -> `图5-6-2 正常采集与OLED显示测试图`
  - `word/media/image24.jpeg` -> `图5-7-1 按键修改阈值界面演示图`
  - `word/media/image25.jpeg` -> `图5-7-2 按键修改阈值界面演示图`

### Remaining Risks

- Huang Man's EMF flowcharts still lack captions because those package members are not exposed through the current `a:blip` relationship scan.
- Fast mode still does not OCR or semantically interpret image content.
- Figure matching is heuristic; unusual table layouts can still confuse nearest-caption assignment.

### Next Round Direction

- Inspect EMF/drawing relationships that are present in package media but missing from `visual_findings[].relationships`.
- Add a narrow media-level OCR path for selected extracted images while keeping fast mode lightweight.

## 2026-06-03 - Media contact sheet and summary index

### Problems Found

- Embedded media extraction produced individual files, but there was no single fast artifact for visually scanning all diagrams/screenshots.
- The report listed extracted media paths, but automation had to inspect the manifest to find labels and contexts.
- Test fixtures for newly added figure captions contained encoding-damaged Chinese strings, making the tests less portable across Windows consoles.

### Changes Completed

- Added `media_summary.json` with media member, extracted path, preview path, hash, content type, derived label, and context records.
- Added `media_contact_sheet.jpg` for image-like media and EMF PNG previews.
- Attached DOCX/PPTX relationship context records directly to matching `embedded_media[]` items.
- Report now lists contact sheet and media summary artifact paths and includes labels when available.
- Replaced fragile non-ASCII caption fixture strings in the new tests with ASCII figure labels while keeping separate Chinese-path coverage.
- Updated `SKILL.md`, `README.md`, and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_extracts_embedded_media_and_report_lists_context tests.test_office_reader.OfficeReaderTests.test_docx_reader_binds_media_to_caption_and_context tests.test_office_reader.OfficeReaderTests.test_docx_reader_uses_table_cell_paragraph_caption_for_media -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 34 tests
- Real document smoke:
  - Source: `D:\QQ文件\(黄曼)基于STM32多参数水质实时监测系统设计与实现 .docx`
  - Command: `read_office.py ... --mode fast --no-openai-vision`
  - Result: success
  - Time: about `1.42s`
  - Produced `media_contact_sheet.jpg` and `media_summary.json`
  - Contact sheet visually showed EMF flowcharts, simulation screens, PCB layout, hardware photos, sensor circuits, and the mobile UI screenshot.
  - `media_summary.json` contained `27` media items; report labels included `图3-1 系统硬件电路原理图`, `图3-2 单片机最小系统框`, and `图5-6-1 正常采集与OLED显示测试图`.

### Remaining Risks

- Contact sheets are triage artifacts, not OCR/vision interpretation.
- Some EMF flowcharts still lack captions/context because the Word package does not expose a simple paragraph relationship for them in current extraction.
- Cross-row/cross-cell caption matching for layout tables remains incomplete.

### Next Round Direction

- Improve context propagation for image-only table cells by scanning neighboring table rows/cells for captions.
- Add optional OCR over contact-sheet source images or selected media blocks for fast visual text extraction.
- Add PPTX media-summary label parity from slide object names, alt text, and slide titles.

## 2026-06-03 - Embedded media extraction and DOCX figure context

### Problems Found

- Fast mode could identify that a DOCX contained media, but did not extract packaged images for quick inspection.
- DOCX image relationships lacked reliable figure context, so callers had to infer image-to-caption mapping from order.
- Images placed inside one-cell tables often carried their figure caption in the same paragraph, but that caption was not attached to the media relationship.
- EMF diagrams required a manual conversion step before visual inspection.
- `read_office.py` UTF-8 stdout improvements initially exposed a Windows child-process decoding issue; sub-script paths could be decoded with replacement characters when the child printed in the local code page.

### Changes Completed

- Added DOCX media relationship context:
  - `paragraph_index`
  - `paragraph_text`
  - `nearest_heading`
  - `nearby_text_before`
  - `nearby_text_after`
  - detected `caption`
  - table cell and container metadata preserved.
- Added embedded media extraction in `visual_analysis.py` for OOXML `word/media` and `ppt/media` members referenced by findings.
- Added cached EMF-to-PNG preview conversion on Windows using GDI+ when possible.
- Added `embedded_media[]` manifest records and `embedded_media_count` in `visual_analysis`.
- Updated reports with `Embedded Media` and `Media context` sections.
- Made `read_office.py` print UTF-8 JSON while decoding child-process output with UTF-8/local-codepage fallback.
- Added regression tests for:
  - DOCX image caption/context binding.
  - table-cell paragraph captions.
  - fast-mode embedded media extraction and report output.
- Updated `SKILL.md`, `README.md`, and `references/output_schema.md`.

### Verification

- Focused tests:
  - `python -m unittest tests.test_office_reader.OfficeReaderTests.test_docx_reader_uses_table_cell_paragraph_caption_for_media tests.test_office_reader.OfficeReaderTests.test_docx_reader_binds_media_to_caption_and_context tests.test_office_reader.OfficeReaderTests.test_visual_analysis_fast_extracts_embedded_media_and_report_lists_context -v`
  - Result: `OK`
- Full test suite:
  - `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 34 tests
- Real document smoke:
  - Source: `D:\QQ文件\(黄曼)基于STM32多参数水质实时监测系统设计与实现 .docx`
  - Command: `read_office.py ... --mode fast --no-openai-vision`
  - Result: success
  - Time: about `1.28s`
  - Extracted embedded media: `27`
  - EMF PNG previews: `8`
  - Report now includes detected captions such as `图3-1 系统硬件电路原理图`, `图3-2 单片机最小系统框`, and `图5-6-1 正常采集与OLED显示测试图`.

### Remaining Risks

- Fast mode extracts media and context but still does not OCR or semantically interpret image content.
- Figure context is heuristic; some Word table layouts split image and caption across separate cells/rows, so not every caption is detected.
- PowerShell 5 must read generated UTF-8 JSON files with `-Encoding UTF8`; otherwise the shell may misdecode Chinese text.
- EMF conversion depends on Windows GDI+ and can fail for unusual vector content.

### Next Round Direction

- Add a lightweight contact-sheet generator and optional media-summary JSON so simple image-content checks can be faster and less manual.
- Improve cross-row/cross-cell caption matching for DOCX figures stored in layout tables.
- Add PPTX embedded media extraction context parity where slide object inventory already knows slide/object metadata.

## 2026-06-02 - Output no-clobber and preview reuse

### Problems Found

- `render_preview.ps1` used a fixed `<basename>.pdf` output path and could overwrite an existing preview PDF or report success from a stale file.
- `convert_legacy_office.ps1` used fixed `<basename>.docx/.pptx` output paths and could overwrite or falsely accept an existing normalized file.
- A first no-clobber pass made repeated `read_office.py` runs degrade visual analysis because the visual pipeline treated an existing preview PDF as a hard render failure.

### Changes Completed

- `render_preview.ps1` now refuses to overwrite an existing preview PDF and returns structured JSON failure.
- `convert_legacy_office.ps1` now refuses to overwrite an existing normalized legacy-conversion output and returns structured JSON failure.
- `visual_analysis.py` now reuses an existing preview PDF when `render_preview.ps1` reports no-clobber, preserving repeated-run behavior without overwriting the PDF.
- Added regression tests for preview PDF no-clobber, legacy conversion no-clobber, and visual-analysis preview PDF reuse.
- Updated `references/backend_fallbacks.md` to document no-clobber behavior and the need for fresh output directories when direct conversion/rendering is rerun.

### Verification

- `python -m unittest discover -s tests -v`
  - Result: `OK`
  - Count: 21 tests
- Real document smoke:
  - Source: `C:\Users\Huang\Desktop\Proj\发电厂课程设计\03_成果输出\最终提交材料\黄登科-202407124102-定稿材料\黄登科-202407124102-定稿.docx`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\huang-202407124102-final`
  - Result: success
  - `visual_analysis.status`: `completed`
  - `rendered_page_count`: 8
  - `cache_hits`: 8
  - `completeness_score.overall`: 94
- Real document smoke:
  - Source: `C:\Users\Huang\Desktop\CC\第十八届合泰杯复赛报告书_基于HT32的无感式智能体态与脊柱健康监测垫_终稿.docx`
  - Output: `C:\Users\Huang\Documents\2123Near\office-reader-real-smoke\ht32-posture-final-cc`
  - Result: success
  - `structure`: 84
  - `tables`: 7
  - `visual_analysis.status`: `completed`
  - `rendered_page_count`: 8
  - `completeness_score.overall`: 92
  - Note: a same-named older file also exists under `C:\Users\Huang\Desktop\HT\其他`; this round validated the newer `Desktop\CC` copy.

### Remaining Risks

- `render_preview.ps1` still leaves COM worker stdout/stderr files and LibreOffice profile directories in some paths.
- `convert_legacy_office.ps1` still lacks COM worker timeout isolation for legacy `.doc/.ppt` conversion.
- PPTX visual objects still lack a structured inventory for alt text, geometry, charts, SmartArt, OLE, video, and audio.

### Next Round Direction

- Implement PPTX visual object inventory v1:
  - image alt text and object name
  - EMU geometry for image objects
  - chart, SmartArt, OLE, video, and audio visual-risk flags
  - compatibility with existing `visual_findings[].media`
