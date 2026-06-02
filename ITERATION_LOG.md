# Office Reader Iteration Log

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
