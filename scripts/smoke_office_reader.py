#!/usr/bin/env python3
"""Run local real-document smoke checks without committing source documents."""

from __future__ import annotations

import argparse
import base64
import json
import os
import subprocess
import sys
import tempfile
import time
import uuid
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent


def positive_int(value: str) -> int:
    number = int(value)
    if number <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return number


def powerpoint_automation_pids() -> set[int] | None:
    command = (
        "Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | "
        "Where-Object { $_.Name -eq 'POWERPNT.EXE' -and $_.CommandLine -match '/Automation|-Embedding' } | "
        "ForEach-Object { $_.ProcessId }"
    )
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=10,
        )
    except Exception:
        return None
    if proc.returncode != 0:
        return None
    return {int(line.strip()) for line in proc.stdout.splitlines() if line.strip().isdigit()}


def stop_new_powerpoint_automation_processes(before: set[int] | None) -> None:
    if before is None:
        return
    after = powerpoint_automation_pids()
    if after is None:
        return
    new_pids = sorted(after - before)
    if not new_pids:
        return
    ids = ",".join(str(pid) for pid in new_pids)
    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"foreach ($id in @({ids})) {{ Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }}",
            ],
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=10,
        )
    except Exception:
        pass


def office_backend_pids() -> set[int] | None:
    command = (
        "Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | "
        "Where-Object { "
        "($_.Name -in @('WINWORD.EXE','POWERPNT.EXE') -and $_.CommandLine -match '/Automation|-Embedding') -or "
        "$_.Name -in @('soffice.exe','soffice.bin') "
        "} | "
        "ForEach-Object { $_.ProcessId }"
    )
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=10,
        )
    except Exception:
        return None
    if proc.returncode != 0:
        return None
    return {int(line.strip()) for line in proc.stdout.splitlines() if line.strip().isdigit()}


def stop_new_office_backend_processes(before: set[int] | None) -> None:
    if before is None:
        return
    after = office_backend_pids()
    if after is None:
        return
    new_pids = sorted(after - before)
    if not new_pids:
        return
    ids = ",".join(str(pid) for pid in new_pids)
    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"foreach ($id in @({ids})) {{ Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }}",
            ],
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=10,
        )
    except Exception:
        pass


def run_command(command: list[str], timeout: int) -> subprocess.CompletedProcess[str]:
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
        timeout=timeout,
        env=env,
    )


def derive_ppt(source: Path, out_dir: Path, timeout: int) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / f"{source.stem}.derived-{uuid.uuid4().hex}.ppt"

    def remove_partial_target() -> None:
        try:
            if target.is_file():
                target.unlink()
        except OSError:
            pass

    script = f"""
$ErrorActionPreference = "Stop"
$before = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {{ $_.Name -eq "POWERPNT.EXE" -and $_.CommandLine -match "/Automation|-Embedding" }} |
    ForEach-Object {{ [int]$_.ProcessId }})
$powerPoint = New-Object -ComObject PowerPoint.Application
$presentation = $powerPoint.Presentations.Open('{str(source).replace("'", "''")}', $true, $false, $false)
try {{
    $presentation.SaveAs('{str(target).replace("'", "''")}', 1)
}} finally {{
    $presentation.Close()
    $powerPoint.Quit()
    Start-Sleep -Milliseconds 200
    Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {{
            $_.Name -eq "POWERPNT.EXE" -and
            $_.CommandLine -match "/Automation|-Embedding" -and
            $_.ProcessId -notin $before
        }} |
        ForEach-Object {{ Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }}
}}
"""
    encoded = base64.b64encode(script.encode("utf-16le")).decode("ascii")
    automation_before = powerpoint_automation_pids()
    try:
        try:
            proc = run_command(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
                timeout=timeout,
            )
        except Exception:
            remove_partial_target()
            raise
    finally:
        stop_new_powerpoint_automation_processes(automation_before)
    if proc.returncode != 0 or not target.is_file():
        remove_partial_target()
        raise RuntimeError(proc.stderr.strip() or proc.stdout.strip() or "PowerPoint COM did not produce a .ppt file.")
    return target


def summarize_manifest(path: Path) -> dict[str, Any]:
    manifest = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(manifest, dict):
        raise ValueError("reader manifest is not a JSON object")
    visual = manifest.get("visual_analysis", {})
    return {
        "document_type": manifest.get("document_type"),
        "conversion": manifest.get("conversion", {}),
        "structure_count": len(manifest.get("structure", [])),
        "table_count": len(manifest.get("tables", [])),
        "comment_count": len(manifest.get("comments", [])),
        "revision_count": len(manifest.get("revisions", [])),
        "note_count": len(manifest.get("notes", [])),
        "visual_status": visual.get("status"),
        "rendered_page_count": visual.get("rendered_page_count", 0),
        "analyzed_item_count": visual.get("analyzed_item_count", 0),
        "visual_messages": visual.get("messages", []),
    }


def run_read(
    source: Path,
    out_dir: Path,
    mode: str,
    legacy_timeout: int,
    visual_timeout: int,
    command_timeout: int,
) -> dict[str, Any]:
    started = time.perf_counter()
    backend_before = office_backend_pids()
    command = [
        sys.executable,
        str(SCRIPT_DIR / "read_office.py"),
        str(source),
        "--out-dir",
        str(out_dir),
        "--mode",
        mode,
        "--no-openai-vision",
        "--legacy-timeout-seconds",
        str(legacy_timeout),
        "--visual-timeout-seconds",
        str(visual_timeout),
    ]
    try:
        proc = run_command(command, timeout=command_timeout)
    except subprocess.TimeoutExpired:
        stop_new_office_backend_processes(backend_before)
        return {
            "source": str(source),
            "mode": mode,
            "status": "failed",
            "duration_ms": int((time.perf_counter() - started) * 1000),
            "messages": [f"Unified reader exceeded smoke timeout of {command_timeout} seconds."],
        }
    duration_ms = int((time.perf_counter() - started) * 1000)
    if proc.returncode != 0:
        return {
            "source": str(source),
            "mode": mode,
            "status": "failed",
            "duration_ms": duration_ms,
            "messages": [proc.stderr.strip() or proc.stdout.strip()],
        }
    try:
        artifacts = json.loads(proc.stdout)
        if not isinstance(artifacts, dict):
            raise ValueError("reader output is not a JSON object")
        manifest_value = artifacts.get("manifest")
        if not manifest_value:
            raise ValueError("reader output did not include a manifest path")
        manifest_path = Path(manifest_value).resolve()
        if not manifest_path.is_file() or not manifest_path.name.endswith(".manifest.json"):
            raise ValueError(f"reader manifest does not exist: {manifest_path}")
        try:
            manifest_path.relative_to(out_dir.resolve())
        except ValueError as exc:
            raise ValueError(f"reader manifest is outside smoke output directory: {manifest_path}") from exc
        manifest_summary = summarize_manifest(manifest_path)
    except (OSError, UnicodeError, ValueError, TypeError) as exc:
        return {
            "source": str(source),
            "mode": mode,
            "status": "failed",
            "duration_ms": duration_ms,
            "messages": [f"Unified reader returned unusable artifacts: {exc}"],
        }
    return {
        "source": str(source),
        "mode": mode,
        "status": "success",
        "duration_ms": duration_ms,
        "artifacts": artifacts,
        "manifest_summary": manifest_summary,
    }


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="Run local office-reader smoke checks with real documents.")
    parser.add_argument("--doc", type=Path)
    parser.add_argument("--docx", type=Path)
    parser.add_argument("--ppt", type=Path)
    parser.add_argument("--pptx", type=Path)
    parser.add_argument("--derive-ppt-from-pptx", action="store_true")
    parser.add_argument("--out-dir", type=Path)
    parser.add_argument("--legacy-timeout-seconds", type=positive_int, default=45)
    parser.add_argument("--visual-timeout-seconds", type=positive_int, default=90)
    parser.add_argument("--command-timeout-seconds", type=positive_int, default=300)
    parser.add_argument("--derive-timeout-seconds", type=positive_int, default=90)
    parser.add_argument("--skip-complete", action="store_true", help="Run only the four-format fast pass.")
    args = parser.parse_args()

    out_dir = (args.out_dir or Path(tempfile.mkdtemp(prefix="office-reader-smoke-"))).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    sources = {
        "doc": args.doc.resolve() if args.doc else None,
        "docx": args.docx.resolve() if args.docx else None,
        "ppt": args.ppt.resolve() if args.ppt else None,
        "pptx": args.pptx.resolve() if args.pptx else None,
    }
    messages: list[str] = []
    if not sources["ppt"] and args.derive_ppt_from_pptx:
        if not sources["pptx"]:
            messages.append("Cannot derive .ppt because --pptx was not provided.")
        else:
            try:
                sources["ppt"] = derive_ppt(sources["pptx"], out_dir / "derived", args.derive_timeout_seconds)
                messages.append(f"Derived local PPT fixture: {sources['ppt']}")
            except Exception as exc:
                messages.append(f"Failed to derive local PPT fixture: {exc}")

    runs: list[dict[str, Any]] = []
    for key in ("doc", "docx", "ppt", "pptx"):
        source = sources[key]
        if not source:
            messages.append(f"Skipped {key}: no input path was provided.")
            continue
        if not source.exists():
            runs.append({"source": str(source), "mode": "fast", "status": "failed", "messages": ["Input file does not exist."]})
            continue
        runs.append(
            run_read(
                source,
                out_dir / f"{key}-fast",
                "fast",
                args.legacy_timeout_seconds,
                args.visual_timeout_seconds,
                args.command_timeout_seconds,
            )
        )

    if not args.skip_complete:
        for key in ("docx", "pptx"):
            source = sources[key]
            if source and source.exists():
                runs.append(
                    run_read(
                        source,
                        out_dir / f"{key}-complete",
                        "complete",
                        args.legacy_timeout_seconds,
                        args.visual_timeout_seconds,
                        args.command_timeout_seconds,
                    )
                )

    result = {
        "status": "success" if runs and all(item["status"] == "success" for item in runs) else "failed",
        "out_dir": str(out_dir),
        "messages": messages,
        "runs": runs,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["status"] == "success" else 1


if __name__ == "__main__":
    raise SystemExit(main())
