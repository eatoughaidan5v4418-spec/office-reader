# Failure Modes

## Missing Visual Dependencies

If OCR, rendering, or OpenAI vision dependencies are missing, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\bootstrap_deps.ps1 -DryRun -IncludeSystemTools
```

Use the dry-run output to decide whether to install local Python packages or system tools. Missing visual dependencies must be recorded in `visual_analysis.messages` and surfaced in the report.

After `-InstallSystemTools`, use the returned final `system_tools` status rather than assuming an invoked `winget` or `choco` command succeeded. `install_result: "failed"` records a non-zero package-manager exit code. `install_result: "not_detected_after_install"` means the command exited successfully but bootstrap still could not discover the executable.

When `read_office.py --install-missing-deps` invokes bootstrap, malformed JSON output, non-object JSON output, or `status: "failed"` is treated as a hard failure. Do not continue document reading after an unverified dependency setup result.

Bootstrap rechecks required Python imports after `pip` completes. If any required package is still missing, it returns `status: "failed"` with the missing package names and exits non-zero.

If RapidOCR or Tesseract is available but returns no text for a rendered page, report that the page may be blank or image-only. Do not mislabel an empty OCR result as a missing dependency.

## Unsupported File Type

Return a clear scope message. This skill covers `.doc`, `.docx`, `.ppt`, and `.pptx` only.

## Legacy Conversion Missing

If Microsoft Office COM, WPS, and LibreOffice are all unavailable, explain that `.doc` or `.ppt` cannot be normalized on this machine yet. Mention that Microsoft Office COM is the preferred path for this skill on this machine.

## Legacy Conversion Timeout

Each legacy conversion backend has an isolated 45-second default timeout. Timeout CLI arguments must be positive integers. If a backend hangs, report its timeout message and continue to the next fallback. Cleanup stops the timed-out worker process tree and newly launched Office automation PIDs that COM reparented outside that tree. Do not kill Office automation PIDs that existed before the worker started.

The unified reader accepts a successful legacy conversion only when `output_path` is a string naming an existing `.docx` or `.pptx` artifact inside the selected output directory. Redirected, missing, malformed, or wrong-extension normalized artifacts are rejected.

Use `scripts/smoke_office_reader.py` for local real-document validation. Keep source documents, derived `.ppt` files, caches, and smoke artifacts out of the repository.

When `--derive-ppt-from-pptx` is used, the smoke harness writes a unique `<stem>.derived-<guid>.ppt` fixture. Repeated local validation does not overwrite an existing derived file.

Smoke PPT derivation snapshots PowerPoint automation PIDs before launching COM and performs best-effort cleanup of newly launched PIDs afterward, including timeout paths. Preexisting PowerPoint automation is preserved.

If PPT derivation fails or times out after PowerPoint partially writes the unique generated target, the harness removes only that run's partial `.ppt` file. Existing fixtures remain untouched.

If a smoke child reader exits successfully but returns malformed JSON, omits the manifest path, redirects the manifest outside that run's smoke output directory, or produces an unreadable or non-object manifest, record that run as failed in the smoke JSON summary instead of aborting the entire harness.

If a unified reader run exceeds the smoke harness command timeout, the harness stops only Office automation and LibreOffice PIDs that were absent from its pre-run snapshot before recording the failed run.

## Damaged OOXML Package

If the ZIP package cannot be opened or required Word parts such as `word/document.xml` are missing, report the damaged part and preserve any partial artifacts if they were written.

For PPTX packages, if `ppt/presentation.xml`, its slide relationships, or referenced slide members are missing or unreadable but `ppt/slides/slideN.xml` members remain, extract those slides in numeric order and record a manifest warning. Treat the result as partial recovery rather than a fully healthy package. If `ppt/presentation.xml` is unavailable and no fallback slide members exist, fail with a damaged-package error instead of emitting a successful zero-slide result. If slides were selected but every slide XML part is unreadable, fail instead of emitting a successful zero-slide result.

## Unicode And Space-Containing Paths

Python and PowerShell entrypoints support Unicode and space-containing paths. JSON-emitting PowerShell scripts escape non-ASCII characters as `\uXXXX` so redirected stdout remains parseable across Windows console code pages. If a path appears garbled in a terminal display, parse the JSON output before concluding that the artifact path is invalid.

## Internal Temporary Artifacts

Legacy conversion and preview rendering delete their internal worker stdout/stderr files after collecting structured messages. LibreOffice profiles are removed only when they are inside the selected output directory and match the generated `.lo_profile_<guid>` naming pattern. Do not recursively delete broader output directories during cleanup.

Legacy conversion and preview timeout cleanup stop the current worker process tree and only the Office automation PIDs absent from the pre-worker snapshot. Do not terminate preexisting Office automation processes.

Legacy conversion does not overwrite an existing normalized `.docx` or `.pptx`. When a same-name file already exists in the output directory, the new artifact is written under a run-specific `legacy-normalized-<guid>` subdirectory.

After a legacy worker fails or times out, the converter removes only the current run's normalized target before trying the next backend. Final conversion failure also removes that partial target.

The unified Python reader accepts legacy conversion only when the PowerShell JSON is parseable, has `status: "success"`, and points to an existing normalized `.docx` or `.pptx` file with the expected extension. A zero process exit code alone is not sufficient.

Preview rendering does not overwrite an existing same-name PDF. The renderer writes a new artifact under a run-specific `preview-render-<guid>` subdirectory and returns that path in its JSON result.

Before preview rendering falls back from Office COM to LibreOffice, it removes only the current run's target PDF. A partial COM export cannot be mistaken for a successful LibreOffice artifact. Final preview failure also removes that partial target.

The visual Python wrapper accepts a preview success result only when it contains an existing PDF artifact inside the current `preview` directory. Missing or redirected artifacts are rejected before rasterization.

If preview rendering returns malformed JSON, a valid non-object JSON payload, non-list `messages` or `artifacts` fields, or non-string artifact paths, visual analysis records a structured failed preview result instead of raising an unhandled parser/type error.

Visual page rasterization writes each run under `page_images/render-<guid>`. It does not overwrite existing `page-*.png` files in the output directory. If rasterization writes no pages, it removes the empty run directory.

Unreadable `.office-reader-cache/*.json` entries are ignored and replaced by a fresh analysis result. A damaged cache entry must not abort document reading.

Visual enrichment appends its Markdown section only when `artifacts.full_markdown` resolves inside the selected output directory. A malformed or manually supplied manifest cannot redirect that write outside `--out-dir`.

The unified reader accepts a DOCX/PPTX child-reader result only when stdout points to an existing, readable JSON-object `.manifest.json` file inside the selected output directory. Its declared Markdown artifact must also exist inside that directory, and its declared manifest artifact must match the returned manifest path. Run-specific nested artifact directories are allowed; redirected, malformed, non-object, and inconsistent manifests are rejected.

The unified reader also requires visual analysis to return the original manifest path, then revalidates the manifest JSON object and its artifact paths after enrichment. Report assembly must return the same-directory `.report.md` path derived from that manifest. Redirected or corrupted subprocess artifacts are rejected.

When `assemble_report.py` is called directly with an unreadable or non-object manifest, it exits non-zero with a concise `Failed to assemble report` diagnostic instead of emitting a Python traceback.

DOCX/PPTX readers preserve unknown same-name `.full.md`, `.manifest.json`, and `.report.md` files in the selected output directory. They write new extraction artifacts under `office-reader-run-<guid>` unless an existing manifest proves the artifacts belong to the same source and document type and records the exact candidate extraction artifact paths.

Generated JSON, Markdown, and report text use same-directory temporary files followed by atomic replacement. A failed write should not leave a partially written final artifact or a generated `.tmp` file.

## Visual-Only Content

Charts, screenshots, scanned slides, SmartArt, and text baked into images may not appear in XML extraction. Treat `visual_findings` as a required follow-up signal for high-stakes reads.

If `--mode fast` was used, say that rendered-page OCR/vision was intentionally skipped.

If `OPENAI_API_KEY` is missing, local OCR can still run but complex chart/diagram interpretation may remain incomplete.

If the user provides portable Poppler or Tesseract, put the executable directory in `OFFICE_READER_TOOL_PATHS` or place it under the skill `tools` directory/current workspace so bootstrap and visual analysis can discover it.

## Comments And Notes

Some producer tools store modern comments in nonstandard locations. If the manifest shows no comments but the rendered document visibly has comments, inspect the OOXML package manually for additional comment parts.

PowerPoint packages may contain `ppt/comments/*.xml` parts without a slide relationship. The reader retains these orphan comments once with `slide_index: null`; it does not guess a slide or duplicate the same comment across slides.

Malformed optional OOXML parts, such as comments, notes, metadata, or relationship files, are skipped so readable body content can still be extracted. Check manifest `warnings` and the report risks section before treating the extraction as complete. Malformed required DOCX body XML remains a hard failure.

OOXML relationship targets are resolved only inside the package. ZIP-root escapes such as `../../../...`, external URLs, network paths, and `TargetMode="External"` relationships are ignored rather than treated as package parts. DOCX and PPTX extraction record a warning when they reject one of these targets.
