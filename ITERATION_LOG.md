# Office Reader Iteration Log

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
