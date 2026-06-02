#!/usr/bin/env python3
"""Unified entrypoint for office-reader."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path

from common_ooxml import atomic_write_text


SCRIPT_DIR = Path(__file__).resolve().parent


def positive_int(value: str) -> int:
    number = int(value)
    if number <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def run_command(command: list[str]) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    return subprocess.run(
        command,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        env=env,
    )


def convert_legacy(source: Path, out_dir: Path, timeout_seconds: int) -> dict:
    proc = run_command(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(SCRIPT_DIR / "convert_legacy_office.ps1"),
            "-InputPath",
            str(source),
            "-OutputDir",
            str(out_dir),
            "-TimeoutSeconds",
            str(timeout_seconds),
        ]
    )
    malformed_output = False
    try:
        result = json.loads(proc.stdout)
        if not isinstance(result, dict):
            raise ValueError("Legacy conversion output is not a JSON object.")
    except (json.JSONDecodeError, ValueError):
        malformed_output = True
        result = {
            "required": True,
            "status": "failed",
            "backend": "",
            "output_path": "",
            "messages": [proc.stderr.strip() or proc.stdout.strip() or "Legacy conversion failed without JSON output."],
        }
    if proc.returncode != 0 or malformed_output or result.get("status") != "success":
        raise RuntimeError(json.dumps(result, ensure_ascii=False, indent=2))
    output_path = result.get("output_path", "")
    normalized = Path(output_path).resolve() if isinstance(output_path, str) and output_path else None
    expected_suffix = ".docx" if source.suffix.lower() == ".doc" else ".pptx"
    usable = bool(normalized and normalized.suffix.lower() == expected_suffix and normalized.is_file())
    if usable:
        try:
            normalized.relative_to(out_dir.resolve())
        except ValueError:
            usable = False
    if not usable:
        messages = result.get("messages")
        if not isinstance(messages, list):
            messages = []
            result["messages"] = messages
        messages.append(
            f"Legacy conversion returned an unusable output_path; expected an existing {expected_suffix} file inside "
            f"{out_dir.resolve()}: {output_path}"
        )
        raise RuntimeError(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def validate_manifest_artifacts(manifest_path: Path, out_dir: Path, producer: str) -> Path:
    manifest_path = manifest_path.resolve()
    if not manifest_path.is_file() or not manifest_path.name.endswith(".manifest.json"):
        raise RuntimeError(f"{producer} did not produce manifest: {manifest_path}")
    try:
        manifest_path.relative_to(out_dir.resolve())
    except ValueError as exc:
        raise RuntimeError(f"{producer} returned manifest outside output directory: {manifest_path}") from exc
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"{producer} returned unreadable manifest: {manifest_path}") from exc
    if not isinstance(manifest, dict):
        raise RuntimeError(f"{producer} returned non-object manifest: {manifest_path}")
    artifacts = manifest.get("artifacts")
    if not isinstance(artifacts, dict):
        raise RuntimeError(f"{producer} manifest did not contain an artifacts object: {manifest_path}")
    try:
        full_markdown = Path(artifacts["full_markdown"]).resolve()
        declared_manifest = Path(artifacts["manifest"]).resolve()
        full_markdown.relative_to(out_dir.resolve())
    except (KeyError, OSError, TypeError, ValueError) as exc:
        raise RuntimeError(f"{producer} manifest contained unusable artifact paths: {manifest_path}") from exc
    if not full_markdown.is_file() or not full_markdown.name.endswith(".full.md") or declared_manifest != manifest_path:
        raise RuntimeError(f"{producer} manifest contained unusable artifact paths: {manifest_path}")
    return manifest_path


def run_reader(normalized: Path, out_dir: Path) -> Path:
    ext = normalized.suffix.lower()
    if ext == ".docx":
        script = SCRIPT_DIR / "read_docx.py"
    elif ext == ".pptx":
        script = SCRIPT_DIR / "read_pptx.py"
    else:
        raise ValueError(f"Unsupported normalized extension: {ext}")

    proc = run_command([sys.executable, str(script), str(normalized), "--out-dir", str(out_dir)])
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
    output_lines = proc.stdout.strip().splitlines()
    if not output_lines:
        raise RuntimeError("Reader did not return a manifest path.")
    return validate_manifest_artifacts(Path(output_lines[-1]), out_dir, "Reader")


def assemble_report(manifest_path: Path) -> Path:
    proc = run_command([sys.executable, str(SCRIPT_DIR / "assemble_report.py"), str(manifest_path)])
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
    output_lines = proc.stdout.strip().splitlines()
    if not output_lines:
        raise RuntimeError("Report assembler did not return a report path.")
    report_path = Path(output_lines[-1]).resolve()
    expected_report = manifest_path.with_name(manifest_path.name.replace(".manifest.json", ".report.md")).resolve()
    if not report_path.is_file() or report_path != expected_report:
        raise RuntimeError(f"Report assembler did not produce report: {report_path}")
    return report_path


def bootstrap_deps(include_system_tools: bool = False, install_system_tools: bool = False) -> dict:
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_DIR / "bootstrap_deps.ps1"),
    ]
    if include_system_tools or install_system_tools:
        command.append("-IncludeSystemTools")
    if install_system_tools:
        command.append("-InstallSystemTools")
    proc = run_command(command)
    malformed_output = False
    try:
        result = json.loads(proc.stdout)
        if not isinstance(result, dict):
            raise ValueError("Dependency bootstrap output is not a JSON object.")
    except (json.JSONDecodeError, ValueError):
        malformed_output = True
        result = {
            "status": "failed",
            "messages": [proc.stderr.strip() or proc.stdout.strip() or "Dependency bootstrap failed without JSON output."],
        }
    if proc.returncode != 0 or malformed_output or result.get("status") == "failed":
        raise RuntimeError(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def run_visual_analysis(
    manifest_path: Path,
    normalized: Path,
    out_dir: Path,
    mode: str,
    enable_openai_vision: bool,
    timeout_seconds: int,
) -> Path:
    command = [
        sys.executable,
        str(SCRIPT_DIR / "visual_analysis.py"),
        str(manifest_path),
        "--normalized-file",
        str(normalized),
        "--out-dir",
        str(out_dir),
        "--mode",
        mode,
        "--timeout-seconds",
        str(timeout_seconds),
    ]
    if not enable_openai_vision:
        command.append("--no-openai-vision")
    proc = run_command(command)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
    output_lines = proc.stdout.strip().splitlines()
    if not output_lines:
        raise RuntimeError("Visual analysis did not return a manifest path.")
    returned_manifest = Path(output_lines[-1]).resolve()
    if not returned_manifest.is_file() or returned_manifest != manifest_path.resolve():
        raise RuntimeError(f"Visual analysis did not update the expected manifest: {returned_manifest}")
    return validate_manifest_artifacts(returned_manifest, out_dir, "Visual analysis")


def update_conversion(manifest_path: Path, source: Path, conversion: dict | None) -> None:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["source"] = {"path": str(source), "name": source.name}
    if conversion:
        manifest["conversion"] = conversion
        output_path = conversion.get("output_path")
        if output_path:
            normalized = Path(output_path)
            manifest["normalized_file"] = {"path": str(normalized), "extension": normalized.suffix.lower()}
    atomic_write_text(manifest_path, json.dumps(manifest, ensure_ascii=False, indent=2))


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Read .doc/.docx/.ppt/.pptx into office-reader artifacts.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out-dir", type=Path, default=None)
    parser.add_argument("--mode", choices=["fast", "balanced", "complete"], default="balanced")
    parser.add_argument("--no-openai-vision", action="store_true")
    parser.add_argument("--install-missing-deps", action="store_true")
    parser.add_argument("--install-system-tools", action="store_true")
    parser.add_argument("--visual-timeout-seconds", type=positive_int, default=90)
    parser.add_argument("--legacy-timeout-seconds", type=positive_int, default=45)
    args = parser.parse_args()

    source = args.source.resolve()
    if not source.exists():
        print(f"Input file does not exist: {source}", file=sys.stderr)
        return 2
    out_dir = (args.out_dir or source.parent / f"{source.stem}.office-reader").resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    ext = source.suffix.lower()
    conversion = None
    normalized = source
    try:
        if args.install_system_tools and not args.install_missing_deps:
            print("--install-system-tools requires --install-missing-deps.", file=sys.stderr)
            return 2
        if args.install_missing_deps:
            bootstrap_deps(include_system_tools=args.install_system_tools, install_system_tools=args.install_system_tools)
        if ext in {".doc", ".ppt"}:
            conversion = convert_legacy(source, out_dir, args.legacy_timeout_seconds)
            normalized_path = conversion.get("output_path")
            if not normalized_path:
                raise RuntimeError("Legacy conversion did not return an output_path.")
            normalized = Path(normalized_path).resolve()
        elif ext not in {".docx", ".pptx"}:
            print("office-reader supports only .doc, .docx, .ppt, and .pptx files.", file=sys.stderr)
            return 2

        manifest_path = run_reader(normalized, out_dir)
        update_conversion(manifest_path, source, conversion)
        manifest_path = run_visual_analysis(
            manifest_path,
            normalized,
            out_dir,
            mode=args.mode,
            enable_openai_vision=not args.no_openai_vision,
            timeout_seconds=args.visual_timeout_seconds,
        )
        report_path = assemble_report(manifest_path)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1

    final_manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    result = {
        "full_markdown": str(final_manifest.get("artifacts", {}).get("full_markdown", "")),
        "manifest": str(manifest_path),
        "report": str(report_path),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
