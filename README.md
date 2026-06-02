# office-reader

Codex skill for deep-reading Microsoft Word and PowerPoint files.

It reads `.doc`, `.docx`, `.ppt`, and `.pptx` into:

- `<basename>.full.md`
- `<basename>.manifest.json`
- `<basename>.report.md`

The default path is Microsoft Office COM first for legacy conversion, then optional WPS, then LibreOffice. Preview rendering uses Office COM when healthy and falls back to LibreOffice; slow or invalid COM preview runs are remembered in `.office-reader-cache/preview-backend-health.json`. Visual deep reading combines OOXML extraction, rendered page/slide analysis, local OCR, optional OpenAI vision, and cache reuse.

Reports include a conservative `completeness_score` so callers can see text/table/visual coverage and remaining unverified visual items. It is a coverage signal, not proof that every visual fact was understood.

For PowerPoint files, the manifest and report include a per-slide visual object inventory for images, charts, SmartArt, OLE objects, video, and audio when those objects are visible in slide relationships. These records are risk/location hints; rendered OCR or OpenAI vision is still required before claiming the visual content itself was understood.

## Quick Start

```powershell
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\read_office.py C:\path\file.pptx --out-dir C:\path\out --mode complete --no-openai-vision
```

`read_office.py` prints ASCII-safe JSON on stdout so automation can parse output paths reliably even when source or output paths contain Chinese characters or spaces.

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
