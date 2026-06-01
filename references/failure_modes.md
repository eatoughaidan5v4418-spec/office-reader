# Failure Modes

## Missing Visual Dependencies

If OCR, rendering, or OpenAI vision dependencies are missing, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -DryRun -IncludeSystemTools
```

Use the dry-run output to decide whether to install local Python packages or system tools. Missing visual dependencies must be recorded in `visual_analysis.messages` and surfaced in the report.

## Unsupported File Type

Return a clear scope message. This skill covers `.doc`, `.docx`, `.ppt`, and `.pptx` only.

## Legacy Conversion Missing

If Microsoft Office COM, WPS, and LibreOffice are all unavailable, explain that `.doc` or `.ppt` cannot be normalized on this machine yet. Mention that Microsoft Office COM is the preferred path for this skill on this machine.

## Damaged OOXML Package

If the ZIP package cannot be opened or required parts such as `word/document.xml` or `ppt/presentation.xml` are missing, report the damaged part and preserve any partial artifacts if they were written.

## Visual-Only Content

Charts, screenshots, scanned slides, SmartArt, and text baked into images may not appear in XML extraction. Treat `visual_findings` as a required follow-up signal for high-stakes reads.

If `--mode fast` was used, say that rendered-page OCR/vision was intentionally skipped.

If `OPENAI_API_KEY` is missing, local OCR can still run but complex chart/diagram interpretation may remain incomplete.

If the user provides portable Poppler or Tesseract, put the executable directory in `OFFICE_READER_TOOL_PATHS` or place it under the skill `tools` directory/current workspace so bootstrap and visual analysis can discover it.

## Comments And Notes

Some producer tools store modern comments in nonstandard locations. If the manifest shows no comments but the rendered document visibly has comments, inspect the OOXML package manually for additional comment parts.
