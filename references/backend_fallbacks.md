# Backend Fallbacks

Legacy `.doc` and `.ppt` inputs must be normalized before reading. The conversion script emits JSON so the calling agent can report exactly what happened.

## Priority Order

1. Microsoft Office COM.
2. WPS Office, only as an optional fallback.
3. LibreOffice/`soffice`.

Microsoft Office COM is the preferred backend on this machine because the user wants Microsoft Office rather than WPS.

## WPS Discovery

WPS is optional. `discover_office_backends.ps1` still checks:

- Explicit `-WpsPath` and `-WppPath` arguments.
- Environment variables `WPS_PATH`, `WPP_PATH`, `WPS_EXE`, and `WPP_EXE`.
- `PATH` commands such as `wps.exe`, `wpp.exe`, `kwps.exe`, `kwpp.exe`, and `wpsoffice.exe`.
- Windows App Paths registry entries.
- Common `WPS Office` and `Kingsoft\WPS Office` install directories.
- WPS COM ProgIDs such as `KWPS.Application`, `WPS.Application`, `KWPP.Application`, and `WPP.Application`.

## Conversion Behavior

The converter attempts Microsoft Office COM first. If Office COM is unavailable or fails, it tries WPS-compatible COM, then LibreOffice. Some WPS builds expose executables but not COM automation; in that case the script records the failure and continues to LibreOffice.

LibreOffice is used through headless `soffice --convert-to` only after Office COM and WPS fail or are missing.

Legacy conversion refuses to overwrite an existing normalized `.docx` or `.pptx` output. Use a fresh output directory when rerunning conversion for the same source.

## Failure Contract

When all backends fail, the script exits non-zero and prints JSON with `status: "failed"` plus messages for every attempted backend. Surface those messages to the user.

## Preview Rendering

`scripts/render_preview.ps1` exports normalized `.docx` or `.pptx` files to PDF for visual inspection. It uses Microsoft Office COM when available and falls back to LibreOffice. It returns structured JSON and enforces a timeout to avoid stuck Office automation workers.

Preview rendering refuses to overwrite an existing `<basename>.pdf` in the preview output directory. Use a fresh preview directory when rerunning a render.

Preview rendering keeps a small backend health file at `.office-reader-cache/preview-backend-health.json` by default. When Office COM preview times out or returns invalid JSON, the script marks `office-com` unhealthy for that normalized extension. Later preview runs skip COM and try fallback preview backends first, preserving the skip reason in JSON messages. A successful COM preview records the backend as healthy again.

Set `OFFICE_READER_PREVIEW_HEALTH_PATH` to isolate this file for tests or one-off diagnostics. The path must end in `.json`; if it already exists, it must be a preview health JSON object. This prevents accidentally overwriting an arbitrary user file. The health memory applies only to preview rendering and must not change the legacy `.doc`/`.ppt` conversion order.

The visual deep-read pipeline uses preview rendering in `balanced` and `complete` modes. `balanced` skips rendering when only lightweight OOXML drawing references exist and no packaged media is present, while `complete` tries to render every normalized document.

## Dependency Installation

`scripts/bootstrap_deps.ps1` creates a skill-local Python virtual environment and installs OCR/vision Python packages there. It can also inspect or install system tools through `winget`/`choco`:

- LibreOffice: conversion/rendering fallback after Microsoft Office COM.
- Poppler: optional PDF utility fallback.
- Tesseract: optional local OCR fallback after RapidOCR.
- WPS Office: optional legacy conversion fallback only when the user wants it installed.

Portable Poppler/Tesseract builds can be placed under the skill `tools` directory, the parent skills directory, or the current workspace. The scripts also honor `OFFICE_READER_TOOL_PATHS` with Windows path-list separators.
