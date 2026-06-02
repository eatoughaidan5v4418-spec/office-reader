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

Each backend attempt runs in an isolated worker process with a 45-second default timeout. Use a positive integer with `-TimeoutSeconds` on `convert_legacy_office.ps1` or `--legacy-timeout-seconds` on `read_office.py` to adjust the budget. On timeout, the converter stops that worker process tree and any newly launched Office automation PID that COM reparented outside the tree, records a backend-specific message, and continues to the next fallback. Office automation PIDs that existed before the worker started are preserved.

The outer converter performs discovery once and skips worker startup for backends already reported unavailable. Available backends still run in isolated workers, preserving timeout and cleanup boundaries while avoiding avoidable PowerShell startup cost for missing optional WPS installations.

If the selected output directory already contains the normalized `<basename>.docx` or `<basename>.pptx`, the converter preserves that file and writes the new normalized artifact into a run-specific `legacy-normalized-<guid>` subdirectory.

LibreOffice discovery accepts `SOFFICE_PATH` or `LIBREOFFICE_PATH` as either an executable path or installation directory. Discovery, dependency dry-run, preview rendering, and conversion use the same override convention.

## Failure Contract

When all backends fail, the script exits non-zero and prints JSON with `status: "failed"` plus messages for every attempted backend. Surface those messages to the user.

## Preview Rendering

`scripts/render_preview.ps1` exports normalized `.docx` or `.pptx` files to PDF for visual inspection. It uses Microsoft Office COM when available and falls back to LibreOffice. It returns structured JSON and enforces a timeout to avoid stuck Office automation workers. Timeout cleanup stops the current preview worker tree plus newly launched reparented Office automation PIDs, while preserving the pre-worker automation snapshot.

For normalized `.docx` preview only, a Word COM timeout records a private machine-local health entry for seven days. Later previews try LibreOffice first while that entry is active, then retry the normal COM-first path if LibreOffice is missing or fails. A successful Word COM preview clears the degraded entry. The default cache is `%LOCALAPPDATA%\office-reader\preview-backend-health.json`; set `OFFICE_READER_PREVIEW_HEALTH_PATH` to override it for isolated runs. This adaptive preview preference does not change legacy conversion order.

If the preview output directory already contains `<basename>.pdf`, the renderer preserves it and writes the new PDF into a run-specific `preview-render-<guid>` subdirectory.

The visual deep-read pipeline uses preview rendering in `balanced` and `complete` modes. `balanced` skips rendering when only lightweight OOXML drawing references exist and no packaged media is present, while `complete` tries to render every normalized document.

## Dependency Installation

`scripts/bootstrap_deps.ps1` creates a skill-local Python virtual environment and installs OCR/vision Python packages there. It can also inspect or install system tools through `winget`/`choco`:

- LibreOffice: conversion/rendering fallback after Microsoft Office COM.
- Poppler: optional PDF utility fallback.
- Tesseract: optional local OCR fallback after RapidOCR.
- WPS Office: optional legacy conversion fallback only when the user wants it installed.

Portable Poppler/Tesseract builds can be placed under the skill `tools` directory, the parent skills directory, or the current workspace. The scripts also honor `OFFICE_READER_TOOL_PATHS` with Windows path-list separators.

After `-InstallSystemTools`, bootstrap reruns tool discovery and returns the final status. Each `system_tools` entry records `install_attempted`, `install_exit_code`, and `install_result` alongside `available` and `source`. A package-manager command that exits successfully but does not produce a discoverable executable is reported as `not_detected_after_install`.
