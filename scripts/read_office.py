#!/usr/bin/env python3
"""Unified entrypoint for office-reader."""

from __future__ import annotations

import argparse
import locale
import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent


def configure_utf8_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def run_command(command: list[str]) -> subprocess.CompletedProcess[str]:
    proc = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
    )
    stdout = decode_process_bytes(proc.stdout)
    stderr = decode_process_bytes(proc.stderr)
    return subprocess.CompletedProcess(proc.args, proc.returncode, stdout, stderr)


def decode_process_bytes(data: bytes | str | None) -> str:
    if data is None:
        return ""
    if isinstance(data, str):
        return data
    encodings = ["utf-8", locale.getpreferredencoding(False), "mbcs", "gbk"]
    for encoding in encodings:
        if not encoding:
            continue
        try:
            return data.decode(encoding)
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", errors="replace")


def conversion_script_path() -> Path:
    override = os.environ.get("OFFICE_READER_CONVERT_LEGACY_SCRIPT")
    return Path(override) if override else SCRIPT_DIR / "convert_legacy_office.ps1"


def legacy_text_script_path() -> Path:
    override = os.environ.get("OFFICE_READER_LEGACY_TEXT_EXTRACTOR")
    return Path(override) if override else SCRIPT_DIR / "extract_legacy_text.ps1"


def convert_legacy(source: Path, out_dir: Path) -> dict:
    proc = run_command(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(conversion_script_path()),
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
        raise RuntimeError(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def parse_conversion_error(exc: Exception) -> dict:
    try:
        return json.loads(str(exc))
    except json.JSONDecodeError:
        return {
            "required": True,
            "status": "failed",
            "backend": "",
            "output_path": "",
            "messages": [str(exc)],
        }


def extract_legacy_text(source: Path, out_dir: Path) -> dict:
    text_path = out_dir / f"{source.stem}.legacy-text.txt"
    proc = run_command(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(legacy_text_script_path()),
            "-InputPath",
            str(source),
            "-OutputPath",
            str(text_path),
        ]
    )
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        result = {
            "status": "failed",
            "backend": "",
            "output_path": "",
            "messages": [proc.stderr.strip() or proc.stdout.strip() or "Legacy text extraction failed without JSON output."],
        }
    if proc.returncode != 0 or result.get("status") != "success":
        raise RuntimeError(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def legacy_text_completeness(structure: list[dict], visual_required: bool) -> dict:
    text_coverage = 100 if structure else 0
    visual_coverage = 0 if visual_required else 100
    ocr_confidence = 0 if visual_required else 100
    unverified = 1 if visual_required else 0
    signals = [
        "Legacy binary Office file was read through a text-only fallback.",
        "Tables, comments, revisions, media, and layout may be incomplete until conversion succeeds.",
    ]
    if visual_required:
        signals.append("Visual content was not rendered or OCR-verified in legacy text fallback.")
    overall = round((text_coverage * 0.35) + (100 * 0.15) + (visual_coverage * 0.40) + (ocr_confidence * 0.10))
    return {
        "overall": max(0, min(100, int(overall))),
        "text_coverage": text_coverage,
        "table_coverage": 100,
        "visual_coverage": visual_coverage,
        "ocr_confidence": ocr_confidence,
        "openai_vision_enabled": False,
        "unverified_visual_count": unverified,
        "signals": signals,
    }


def write_legacy_text_artifacts(source: Path, out_dir: Path, extraction: dict, conversion: dict | None, mode: str) -> Path:
    text_path = Path(extraction.get("output_path", ""))
    text = text_path.read_text(encoding="utf-8", errors="replace").lstrip("\ufeff") if text_path.exists() else ""
    lines = [line.strip() for line in text.replace("\r", "\n").split("\n")]
    paragraphs = [line for line in lines if line]
    structure = [
        {
            "type": "paragraph",
            "index": index,
            "text": paragraph,
            "part_type": "legacy_text",
        }
        for index, paragraph in enumerate(paragraphs, start=1)
    ]
    markdown_lines = [f"# {source.stem}", ""]
    markdown_lines.extend(paragraphs or [""])
    full_path = out_dir / f"{source.stem}.full.md"
    manifest_path = out_dir / f"{source.stem}.manifest.json"
    ext = source.suffix.lower().lstrip(".")
    messages = list(extraction.get("messages", []))
    if conversion and conversion.get("messages"):
        messages = list(conversion.get("messages", [])) + messages
    fallback_conversion = {
        "required": True,
        "status": "text_fallback",
        "backend": extraction.get("backend", "legacy-text"),
        "output_path": str(text_path),
        "messages": messages,
    }
    if conversion:
        fallback_conversion["conversion_attempt"] = conversion
    visual_required = True
    manifest = {
        "source": {"path": str(source), "name": source.name},
        "normalized_file": {"path": str(text_path), "extension": ".txt"},
        "conversion": fallback_conversion,
        "document_type": ext,
        "reading_mode": mode,
        "metadata": {},
        "structure": structure,
        "tables": [],
        "comments": [],
        "revisions": [],
        "notes": [],
        "visual_findings": [
            {
                "requires_visual_review": visual_required,
                "reason": "legacy text fallback cannot inspect binary Office layout, tables, or embedded media",
                "media_count": 0,
                "media": [],
                "relationships": [],
            }
        ],
        "visual_analysis": {
            "status": "skipped",
            "mode": mode,
            "rendered_page_count": 0,
            "analyzed_item_count": 0,
            "cache_hits": 0,
            "messages": ["Legacy text fallback skipped rendering and OCR."],
        },
        "completeness_score": legacy_text_completeness(structure, visual_required),
        "artifacts": {"full_markdown": str(full_path), "manifest": str(manifest_path), "legacy_text": str(text_path)},
    }
    full_path.write_text("\n\n".join(markdown_lines).strip() + "\n", encoding="utf-8")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
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


def compact_text(value: Any, limit: int = 260) -> str:
    text = " ".join(str(value or "").split())
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "..."


def query_tokens(query: str) -> list[str]:
    return [token.casefold() for token in query.split() if token.strip()]


def text_matches_query(text: str, tokens: list[str]) -> bool:
    haystack = text.casefold()
    return bool(tokens) and all(token in haystack for token in tokens)


def add_query_candidate(candidates: list[dict[str, Any]], source_type: str, location: dict[str, Any], text: Any) -> None:
    normalized = " ".join(str(text or "").split())
    if normalized:
        candidates.append({"source_type": source_type, "location": location, "text": normalized})


def query_candidates(manifest: dict[str, Any]) -> list[dict[str, Any]]:
    candidates: list[dict[str, Any]] = []
    for item in manifest.get("structure", []):
        location = {
            "index": item.get("index"),
            "slide_index": item.get("slide_index"),
            "part_type": item.get("part_type"),
            "part": item.get("part"),
            "container": item.get("container"),
        }
        add_query_candidate(candidates, "structure", location, item.get("title") or item.get("text"))
    for table in manifest.get("tables", []):
        for row_index, row in enumerate(table.get("rows", []), start=1):
            location = {
                "table_index": table.get("index"),
                "row_index": row_index,
                "slide_index": table.get("slide_index"),
                "part_type": table.get("part_type"),
                "part": table.get("part"),
                "container": table.get("container"),
            }
            add_query_candidate(candidates, "table", location, " | ".join(str(cell) for cell in row))
    for comment in manifest.get("comments", []):
        location = {
            "id": comment.get("id"),
            "slide_index": comment.get("slide_index"),
            "table_index": comment.get("table_index"),
            "row_index": comment.get("row_index"),
            "cell_index": comment.get("cell_index"),
            "part_type": comment.get("part_type"),
            "part": comment.get("part"),
            "container": comment.get("container"),
            "anchor_text": comment.get("anchor_text"),
        }
        add_query_candidate(candidates, "comment", location, comment.get("text"))
        add_query_candidate(candidates, "comment_anchor", location, comment.get("anchor_text"))
    for revision in manifest.get("revisions", []):
        location = {
            "type": revision.get("type"),
            "table_index": revision.get("table_index"),
            "row_index": revision.get("row_index"),
            "cell_index": revision.get("cell_index"),
            "part_type": revision.get("part_type"),
            "part": revision.get("part"),
            "container": revision.get("container"),
        }
        add_query_candidate(candidates, "revision", location, revision.get("text"))
    for note in manifest.get("notes", []):
        add_query_candidate(candidates, "speaker_note", {"slide_index": note.get("slide_index")}, note.get("text"))
    for finding in manifest.get("visual_findings", []):
        location = {"page_index": finding.get("page_index"), "slide_index": finding.get("slide_index")}
        for key, source_type in (
            ("ocr_text", "visual_ocr"),
            ("vision_summary", "visual_summary"),
            ("diagram_summary", "diagram_summary"),
            ("reason", "visual_finding"),
        ):
            add_query_candidate(candidates, source_type, location, finding.get(key))
        for rel in finding.get("relationships", []):
            rel_location = {
                "target": rel.get("target"),
                "paragraph_index": rel.get("paragraph_index"),
                "table_index": rel.get("table_index"),
                "row_index": rel.get("row_index"),
                "cell_index": rel.get("cell_index"),
                "slide_index": rel.get("slide_index"),
                "part_type": rel.get("part_type"),
                "part": rel.get("part"),
                "container": rel.get("container"),
            }
            for key in ("caption", "paragraph_text", "nearest_heading", "nearby_text_before", "nearby_text_after", "alt_text", "title", "name"):
                add_query_candidate(candidates, f"media_{key}", rel_location, rel.get(key))
    for item in manifest.get("embedded_media", []):
        location = {"member": item.get("member"), "path": item.get("path")}
        add_query_candidate(candidates, "embedded_media", location, item.get("label") or item.get("member"))
    return candidates


def apply_query(manifest_path: Path, query: str, limit: int = 20) -> Path:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    tokens = query_tokens(query)
    matches = []
    seen = set()
    for candidate in query_candidates(manifest):
        text = candidate.get("text", "")
        if not text_matches_query(text, tokens):
            continue
        identity = (candidate.get("source_type"), json.dumps(candidate.get("location", {}), sort_keys=True), text)
        if identity in seen:
            continue
        seen.add(identity)
        matches.append(
            {
                "source_type": candidate.get("source_type"),
                "location": candidate.get("location", {}),
                "text": compact_text(text, 500),
            }
        )
    query_path = manifest_path.with_name(manifest_path.name.replace(".manifest.json", ".query.json"))
    result = {
        "query": query,
        "tokens": tokens,
        "total_matches": len(matches),
        "matches": matches[:limit],
        "truncated": len(matches) > limit,
    }
    query_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    manifest["query"] = result
    manifest.setdefault("artifacts", {})["query_results"] = str(query_path)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return query_path


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
        raise RuntimeError(json.dumps(result, ensure_ascii=False, indent=2))
    return result


def run_visual_analysis(
    manifest_path: Path,
    normalized: Path,
    out_dir: Path,
    mode: str,
    enable_openai_vision: bool,
    timeout_seconds: int,
    media_ocr: str,
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
        "--media-ocr",
        media_ocr,
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
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="Read .doc/.docx/.ppt/.pptx into office-reader artifacts.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out-dir", type=Path, default=None)
    parser.add_argument("--mode", choices=["fast", "balanced", "complete"], default="balanced")
    parser.add_argument("--no-openai-vision", action="store_true")
    parser.add_argument("--install-missing-deps", action="store_true")
    parser.add_argument("--install-system-tools", action="store_true")
    parser.add_argument("--visual-timeout-seconds", type=int, default=90)
    parser.add_argument("--query", default=None, help="Write a query-results artifact for a quick text lookup.")
    parser.add_argument("--query-limit", type=int, default=20, help="Maximum query matches to include in the artifact.")
    parser.add_argument("--media-ocr", choices=["off", "selected", "all"], default="off")
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
    query_path = None
    try:
        if args.install_missing_deps:
            bootstrap_deps(include_system_tools=args.install_system_tools)
        if ext in {".doc", ".ppt"}:
            if args.mode == "fast":
                extraction = extract_legacy_text(source, out_dir)
                manifest_path = write_legacy_text_artifacts(source, out_dir, extraction, None, args.mode)
                if args.query:
                    query_path = apply_query(manifest_path, args.query, args.query_limit)
                report_path = assemble_report(manifest_path)
                normalized = Path(extraction.get("output_path", source)).resolve()
                result = {
                    "full_markdown": str(out_dir / f"{source.stem}.full.md"),
                    "manifest": str(manifest_path),
                    "report": str(report_path),
                }
                if query_path:
                    result["query_results"] = str(query_path)
                print(json.dumps(result, ensure_ascii=False, indent=2))
                return 0
            try:
                conversion = convert_legacy(source, out_dir)
                normalized_path = conversion.get("output_path")
                if not normalized_path:
                    raise RuntimeError("Legacy conversion did not return an output_path.")
                normalized = Path(normalized_path).resolve()
            except Exception as conversion_exc:
                conversion_failure = parse_conversion_error(conversion_exc)
                extraction = extract_legacy_text(source, out_dir)
                manifest_path = write_legacy_text_artifacts(source, out_dir, extraction, conversion_failure, args.mode)
                if args.query:
                    query_path = apply_query(manifest_path, args.query, args.query_limit)
                report_path = assemble_report(manifest_path)
                normalized = Path(extraction.get("output_path", source)).resolve()
                result = {
                    "full_markdown": str(out_dir / f"{source.stem}.full.md"),
                    "manifest": str(manifest_path),
                    "report": str(report_path),
                }
                if query_path:
                    result["query_results"] = str(query_path)
                print(json.dumps(result, ensure_ascii=False, indent=2))
                return 0
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
            media_ocr=args.media_ocr,
        )
        if args.query:
            query_path = apply_query(manifest_path, args.query, args.query_limit)
        report_path = assemble_report(manifest_path)
    except Exception as exc:
        print(str(exc), file=sys.stderr)
        return 1

    result = {
        "full_markdown": str(out_dir / f"{normalized.stem}.full.md"),
        "manifest": str(manifest_path),
        "report": str(report_path),
    }
    if query_path:
        result["query_results"] = str(query_path)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
