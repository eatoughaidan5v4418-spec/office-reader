# office-reader

Codex skill for deep-reading Microsoft Word and PowerPoint files.

It reads `.doc`, `.docx`, `.ppt`, and `.pptx` into:

- `<basename>.full.md`
- `<basename>.manifest.json`
- `<basename>.report.md`

The default path is Microsoft Office COM first for legacy conversion/rendering, then optional WPS, then LibreOffice. Visual deep reading combines OOXML extraction, rendered page/slide analysis, local OCR, optional OpenAI vision, and cache reuse.

## Quick Start

```powershell
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\read_office.py C:\path\file.pptx --out-dir C:\path\out --mode complete --no-openai-vision
```

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
