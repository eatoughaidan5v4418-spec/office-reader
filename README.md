# office-reader

Codex skill for deep-reading Microsoft Word and PowerPoint files.

It reads `.doc`, `.docx`, `.ppt`, and `.pptx` into:

- `<basename>.full.md`
- `<basename>.manifest.json`
- `<basename>.report.md`

The default path is Microsoft Office COM first for legacy conversion, then optional WPS, then LibreOffice. Preview rendering uses Office COM when healthy and falls back to LibreOffice; slow or invalid COM preview runs are remembered in `.office-reader-cache/preview-backend-health.json`. Visual deep reading combines OOXML extraction, embedded media extraction, rendered page/slide analysis, local OCR, optional OpenAI vision, and cache reuse.

Legacy `.doc`/`.ppt` Office COM conversion also runs in an isolated worker with timeout protection and health memory. Slow or failed Office COM conversion attempts are recorded in `.office-reader-cache/conversion-backend-health.json`, so later conversion runs can skip directly to WPS/LibreOffice fallback backends while reporting the reason.

Reports include a conservative `completeness_score` so callers can see text/table/visual coverage and remaining unverified visual items. It is a coverage signal, not proof that every visual fact was understood.

For PowerPoint files, the manifest and report include a per-slide visual object inventory for images, charts, SmartArt, OLE objects, video, and audio when those objects are visible in slide relationships. These records are risk/location hints; rendered OCR or OpenAI vision is still required before claiming the visual content itself was understood.

For Word files, media relationships include nearby context when OOXML exposes it: paragraph text, table cell coordinates, nearest heading, adjacent text, and detected figure/table captions. The visual pipeline extracts packaged `word/media` or `ppt/media` files into `embedded_media/`; EMF files get cached PNG previews on Windows when GDI+ conversion succeeds. Image-like media are also indexed in `media_summary.json` and `media_contact_sheet.jpg` for quick visual triage.

## Quick Start

```powershell
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\read_office.py C:\path\file.pptx --out-dir C:\path\out --mode complete --no-openai-vision
python scripts\read_office.py C:\path\legacy.doc --out-dir C:\path\out --mode fast
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode fast --query "ARR source"
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode fast --media-ocr selected --query "sensor"
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced --evidence-report
```

`read_office.py` configures stdout/stderr as UTF-8 and prints a small JSON object with generated paths. Automation should decode stdout as UTF-8; PowerShell 5 users should read manifest/report files with `-Encoding UTF8`.

Use `--mode fast` for simple lookup tasks such as finding a section, question, or keyword. For legacy `.doc` and `.ppt`, fast mode uses a text-only fallback first. If full legacy conversion fails in `balanced` or `complete`, the unified reader also falls back to searchable text artifacts and marks `conversion.status` as `text_fallback`.

Use `--query "<text>"` with any mode to generate `<basename>.query.json`, add `manifest.query`, include `query_results` in stdout, and add a Query Results section to the report. Query mode searches extracted structure text, tables, comments, comment anchors, revisions, speaker notes, OCR/vision text fields, media context, and embedded-media labels.

Use `--media-ocr selected` to OCR selected extracted embedded images or EMF previews without rendering full pages. Use `--media-ocr all` to attempt every image-like extracted media item. Media OCR results are written to `embedded_media[].ocr_text`, `media_summary.json`, `visual_findings[]`, and the report, so `--query` can find text recovered from embedded images.

Use `--evidence-report` to append an Evidence Index to the report. It lists source-backed entries from structure, tables, comments, revisions, speaker notes, media relationships, visual object inventory, and OCR findings with manifest locations.

## Dependencies

Python dependencies are installed into the skill-local `.venv`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -IncludeSystemTools
```

Inspect missing optional tools without installing:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -DryRun -IncludeSystemTools
```

Portable Poppler/Tesseract builds are supported. Put executable folders under `tools/`, the current workspace, or set `OFFICE_READER_TOOL_PATHS` to the `bin` directories.

## Validation

```powershell
python -m unittest discover -s tests -v
```
