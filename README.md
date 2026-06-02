# office-reader

Codex skill for deep-reading Microsoft Word and PowerPoint files.

It reads `.doc`, `.docx`, `.ppt`, and `.pptx` into:

- `<basename>.full.md`
- `<basename>.manifest.json`
- `<basename>.report.md`

The default path is Microsoft Office COM first for legacy conversion/rendering, then optional WPS, then LibreOffice. Visual deep reading combines OOXML extraction, rendered page/slide analysis, local OCR, optional OpenAI vision, and cache reuse.

Legacy `.doc` and `.ppt` conversion runs each backend in an isolated worker. A backend that exceeds the default 45-second budget is stopped before the reader continues to the next fallback.

## Quick Start

```powershell
python scripts\read_office.py C:\path\file.docx --out-dir C:\path\out --mode balanced
python scripts\read_office.py C:\path\file.pptx --out-dir C:\path\out --mode complete --no-openai-vision
python scripts\read_office.py C:\path\file.doc --out-dir C:\path\out --mode fast --legacy-timeout-seconds 45
python scripts\read_office.py C:\path\file.docx --install-missing-deps --install-system-tools
```

`--install-system-tools` requires `--install-missing-deps`. The combination explicitly forwards system fallback installation to `bootstrap_deps.ps1`; omit it when only skill-local Python packages should be installed.

## Dependencies

Python dependencies are installed into the skill-local `.venv`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -IncludeSystemTools
```

Inspect missing optional tools without installing:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -DryRun -IncludeSystemTools
```

After a system-tool install request, bootstrap reruns discovery before returning JSON. Check each `system_tools` entry for final `available`, `source`, `install_attempted`, `install_exit_code`, and `install_result` values instead of assuming that an invoked package manager completed the setup.

Portable Poppler/Tesseract builds are supported. Put executable folders under `tools/`, the current workspace, or set `OFFICE_READER_TOOL_PATHS` to the `bin` directories.

For LibreOffice outside the usual install locations, set `SOFFICE_PATH` or `LIBREOFFICE_PATH` to the executable or installation directory.

## Local Real-Document Smoke Check

```powershell
python scripts\smoke_office_reader.py --doc C:\path\file.doc --docx C:\path\file.docx --pptx C:\path\file.pptx --derive-ppt-from-pptx --visual-timeout-seconds 180
```

The smoke harness keeps real inputs and derived legacy files local. Derived PPT fixtures use unique `<stem>.derived-<guid>.ppt` names so repeated checks do not overwrite an existing file. Do not commit them.
Use a larger `--visual-timeout-seconds` value for long Word documents whose complete-mode preview export may need to fall back from Office COM to LibreOffice.

## Validation

```powershell
python -m unittest discover -s tests -v
```
