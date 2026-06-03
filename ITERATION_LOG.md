# Office Reader Iteration Log

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
