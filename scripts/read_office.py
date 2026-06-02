#!/usr/bin/env python3
"""Unified entrypoint for office-reader."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent


def run_command(command: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        text=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )


def convert_legacy(source: Path, out_dir: Path) -> dict:
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
        ]
    )
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        result = {
            "required": True,
            "status": "failed",
            "backend": "",
            "output_path": "",
            "messages": [proc.stderr.strip() or proc.stdout.strip() or "Legacy conversion failed without JSON output."],
        }
    if proc.returncode != 0:
        raise RuntimeError(json.dumps(result, ensure_ascii=True, indent=2))
    return result


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
    manifest_path = Path(proc.stdout.strip().splitlines()[-1])
    if not manifest_path.exists():
        raise RuntimeError(f"Reader did not produce manifest: {manifest_path}")
    return manifest_path


def assemble_report(manifest_path: Path) -> Path:
    proc = run_command([sys.executable, str(SCRIPT_DIR / "assemble_report.py"), str(manifest_path)])
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
    report_path = Path(proc.stdout.strip().splitlines()[-1])
    if not report_path.exists():
        raise RuntimeError(f"Report assembler did not produce report: {report_path}")
    return report_path


def bootstrap_deps(include_system_tools: bool = False) -> dict:
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(SCRIPT_DIR / "bootstrap_deps.ps1"),
    ]
    if include_system_tools:
        command.append("-IncludeSystemTools")
    proc = run_command(command)
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        result = {
            "status": "failed",
            "messages": [proc.stderr.strip() or proc.stdout.strip() or "Dependency bootstrap failed without JSON output."],
        }
    if proc.returncode != 0:
        raise RuntimeError(json.dumps(result, ensure_ascii=True, indent=2))
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
    return Path(proc.stdout.strip().splitlines()[-1])


def update_conversion(manifest_path: Path, source: Path, conversion: dict | None) -> None:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    manifest["source"] = {"path": str(source), "name": source.name}
    if conversion:
        manifest["conversion"] = conversion
        output_path = conversion.get("output_path")
        if output_path:
            normalized = Path(output_path)
            manifest["normalized_file"] = {"path": str(normalized), "extension": normalized.suffix.lower()}
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Read .doc/.docx/.ppt/.pptx into office-reader artifacts.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out-dir", type=Path, default=None)
    parser.add_argument("--mode", choices=["fast", "balanced", "complete"], default="balanced")
    parser.add_argument("--no-openai-vision", action="store_true")
    parser.add_argument("--install-missing-deps", action="store_true")
    parser.add_argument("--install-system-tools", action="store_true")
    parser.add_argument("--visual-timeout-seconds", type=int, default=90)
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
        if args.install_missing_deps:
            bootstrap_deps(include_system_tools=args.install_system_tools)
        if ext in {".doc", ".ppt"}:
            conversion = convert_legacy(source, out_dir)
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

    result = {
        "full_markdown": str(out_dir / f"{normalized.stem}.full.md"),
        "manifest": str(manifest_path),
        "report": str(report_path),
    }
    print(json.dumps(result, ensure_ascii=True, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
