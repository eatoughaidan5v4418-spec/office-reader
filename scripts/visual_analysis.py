#!/usr/bin/env python3
"""Visual deep-read enrichment for office-reader manifests."""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import os
import shutil
import site
import subprocess
import sys
import time
import zipfile
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
VISUAL_MARKER = "<!-- office-reader-visual-deep-read -->"
VISUAL_FIELDS = {
    "page_index": None,
    "slide_index": None,
    "image_path": "",
    "ocr_text": "",
    "vision_summary": "",
    "diagram_summary": "",
    "confidence": "unknown",
    "backend": "",
    "duration_ms": 0,
    "cache_hit": False,
}


def add_skill_venv_to_path() -> None:
    venv_dir = SKILL_DIR / ".venv"
    candidates = [
        venv_dir / "Lib" / "site-packages",
        venv_dir / "lib" / f"python{sys.version_info.major}.{sys.version_info.minor}" / "site-packages",
    ]
    for candidate in candidates:
        if candidate.exists():
            site.addsitedir(str(candidate))


def local_tool_directories() -> list[Path]:
    directories: list[Path] = []
    raw = os.environ.get("OFFICE_READER_TOOL_PATHS", "")
    for item in raw.split(os.pathsep):
        if item.strip():
            directories.append(Path(item.strip()))
    roots = [
        SKILL_DIR / "tools",
        SKILL_DIR.parent,
        Path.cwd(),
        Path.cwd().parent,
    ]
    for root in roots:
        if not root.exists():
            continue
        for pattern in ("poppler*", "tesseract*", "ocr*", "tools"):
            directories.extend(path for path in root.glob(pattern) if path.is_dir())
    expanded: list[Path] = []
    for directory in directories:
        expanded.append(directory)
        expanded.append(directory / "bin")
        expanded.append(directory / "Library" / "bin")
    seen: set[str] = set()
    result: list[Path] = []
    for directory in expanded:
        try:
            resolved = directory.resolve()
        except OSError:
            continue
        key = str(resolved).lower()
        if key not in seen and resolved.exists():
            result.append(resolved)
            seen.add(key)
    return result


def add_local_tools_to_path() -> None:
    paths = [str(path) for path in local_tool_directories()]
    if paths:
        os.environ["PATH"] = os.pathsep.join([*paths, os.environ.get("PATH", "")])


add_skill_venv_to_path()
add_local_tools_to_path()


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def package_member_hash(package: zipfile.ZipFile, member: str) -> str:
    digest = hashlib.sha256()
    with package.open(member) as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def stable_hash(data: Any) -> str:
    encoded = json.dumps(data, ensure_ascii=False, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def unique_media_members(manifest: dict[str, Any]) -> list[str]:
    members: list[str] = []
    for finding in manifest.get("visual_findings", []) or []:
        for media in finding.get("media", []) or []:
            if isinstance(media, str):
                members.append(media.replace("\\", "/"))
        for rel in finding.get("relationships", []) or []:
            target = str(rel.get("target", "")).replace("\\", "/")
            if target:
                members.append(target)
        for obj in finding.get("objects", []) or []:
            target = str(obj.get("target", "")).replace("\\", "/")
            if target:
                members.append(target)
    seen: set[str] = set()
    result: list[str] = []
    for member in members:
        if "/media/" not in member:
            continue
        key = member.lower()
        if key not in seen:
            result.append(member)
            seen.add(key)
    return result


def media_output_name(member: str, digest: str) -> str:
    suffix = Path(member).suffix.lower() or ".bin"
    stem = Path(member).stem
    return f"{stem}-{digest[:12]}{suffix}"


def convert_emf_to_png(source: Path, cache_dir: Path, messages: list[str]) -> Path | None:
    digest = sha256_file(source)
    dest = cache_dir / f"{source.stem}-{digest[:12]}.png"
    if dest.exists():
        return dest
    try:
        from PIL import Image  # type: ignore

        with Image.open(source) as image:
            image.save(dest)
        return dest
    except Exception:
        pass
    if os.name != "nt":
        messages.append(f"EMF conversion skipped for {source.name}: Windows GDI+ is unavailable.")
        return None
    source_ps = str(source).replace("'", "''")
    dest_ps = str(dest).replace("'", "''")
    script = (
        "Add-Type -AssemblyName System.Drawing; "
        f"$mf=New-Object System.Drawing.Imaging.Metafile('{source_ps}'); "
        "try { "
        "$w=[Math]::Max(800,[int]($mf.Width*3)); $h=[Math]::Max(500,[int]($mf.Height*3)); "
        "$bmp=New-Object System.Drawing.Bitmap($w,$h); "
        "$g=[System.Drawing.Graphics]::FromImage($bmp); $g.Clear([System.Drawing.Color]::White); "
        "$g.DrawImage($mf,0,0,$w,$h); "
        f"$bmp.Save('{dest_ps}',[System.Drawing.Imaging.ImageFormat]::Png); "
        "$g.Dispose(); $bmp.Dispose(); "
        "} finally { $mf.Dispose() }"
    )
    try:
        proc = run_command(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script], timeout=30)
    except Exception as exc:
        messages.append(f"EMF conversion failed for {source.name}: {exc}")
        return None
    if proc.returncode == 0 and dest.exists():
        return dest
    detail = (proc.stderr or proc.stdout or "").strip()
    messages.append(f"EMF conversion failed for {source.name}: {detail}")
    return None


def extract_embedded_media(
    normalized_file: Path,
    manifest: dict[str, Any],
    out_dir: Path,
    cache_dir: Path,
    messages: list[str],
) -> list[dict[str, Any]]:
    members = unique_media_members(manifest)
    if not members or normalized_file.suffix.lower() not in {".docx", ".pptx"}:
        return []
    media_dir = out_dir / "embedded_media"
    media_dir.mkdir(parents=True, exist_ok=True)
    emf_cache = cache_dir / "emf_png"
    emf_cache.mkdir(parents=True, exist_ok=True)
    extracted: list[dict[str, Any]] = []
    with zipfile.ZipFile(normalized_file) as package:
        names = set(package.namelist())
        for member in members:
            if member not in names:
                messages.append(f"Embedded media member was referenced but not found: {member}")
                continue
            digest = package_member_hash(package, member)
            dest = media_dir / media_output_name(member, digest)
            cache_hit = dest.exists()
            if not cache_hit:
                with package.open(member) as source, dest.open("wb") as target:
                    shutil.copyfileobj(source, target)
            item: dict[str, Any] = {
                "member": member,
                "path": str(dest),
                "sha256": digest,
                "content_type": dest.suffix.lower().lstrip(".") or "binary",
                "cache_hit": cache_hit,
            }
            if dest.suffix.lower() == ".emf":
                png = convert_emf_to_png(dest, emf_cache, messages)
                if png:
                    item["preview_path"] = str(png)
                    item["preview_format"] = "png"
            extracted.append(item)
    return extracted


def ensure_visual_fields(finding: dict[str, Any]) -> dict[str, Any]:
    for key, value in VISUAL_FIELDS.items():
        finding.setdefault(key, value)
    return finding


def run_command(command: list[str], timeout: int = 120) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        check=False,
        timeout=timeout,
    )


def render_preview(normalized_file: Path, out_dir: Path, timeout_seconds: int, messages: list[str]) -> tuple[Path | None, dict[str, Any]]:
    preview_dir = out_dir / "preview"
    preview_dir.mkdir(parents=True, exist_ok=True)
    script = SCRIPT_DIR / "render_preview.ps1"
    command = [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-InputPath",
        str(normalized_file),
        "-OutputDir",
        str(preview_dir),
        "-TimeoutSeconds",
        str(timeout_seconds),
        "-ContinueAfterComFailure",
    ]
    try:
        proc = run_command(command, timeout=max(timeout_seconds + 20, 30))
    except subprocess.TimeoutExpired:
        messages.append(f"Preview rendering command timed out after {timeout_seconds} seconds.")
        return None, {"status": "failed", "messages": messages[-1:]}
    try:
        result = json.loads(proc.stdout)
    except json.JSONDecodeError:
        detail = (proc.stderr or proc.stdout or "").strip()
        messages.append(f"Preview rendering did not return JSON: {detail}")
        return None, {"status": "failed", "messages": [detail]}
    for message in result.get("messages", []):
        if message:
            messages.append(str(message))
    if proc.returncode != 0 or result.get("status") != "success":
        existing_pdf = preview_dir / f"{normalized_file.stem}.pdf"
        if existing_pdf.exists() and any("already exists" in str(item).lower() for item in result.get("messages", [])):
            messages.append(f"Reusing existing preview PDF without overwriting: {existing_pdf}")
            reuse_result = dict(result)
            reuse_result["status"] = "success"
            reuse_result["backend"] = "existing-preview"
            reuse_result["artifacts"] = [str(existing_pdf)]
            return existing_pdf, reuse_result
        return None, result
    artifacts = [Path(item) for item in result.get("artifacts", []) if item]
    pdf = next((item for item in artifacts if item.suffix.lower() == ".pdf" and item.exists()), None)
    return pdf, result


def render_pdf_pages(pdf_path: Path, out_dir: Path, messages: list[str], max_pages: int | None = None) -> list[Path]:
    image_dir = out_dir / "page_images"
    image_dir.mkdir(parents=True, exist_ok=True)
    try:
        import fitz  # type: ignore
    except Exception as exc:
        messages.append(f"PyMuPDF is unavailable, so preview PDF pages were not rasterized: {exc}")
        return []

    pages: list[Path] = []
    try:
        with fitz.open(str(pdf_path)) as document:
            page_total = len(document)
            limit = page_total if max_pages is None else min(page_total, max_pages)
            for index in range(limit):
                page = document.load_page(index)
                pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image_path = image_dir / f"page-{index + 1:03d}.png"
                pixmap.save(str(image_path))
                pages.append(image_path)
    except Exception as exc:
        messages.append(f"Preview PDF page rendering failed: {exc}")
    return pages


def tesseract_ocr(image_path: Path) -> tuple[str, str]:
    exe = shutil.which("tesseract")
    if not exe:
        return "", ""
    try:
        proc = run_command([exe, str(image_path), "stdout", "-l", "chi_sim+eng", "--psm", "6"], timeout=60)
    except Exception:
        return "", ""
    if proc.returncode == 0 and proc.stdout.strip():
        return proc.stdout.strip(), "tesseract"
    return "", ""


def rapidocr_text(image_path: Path) -> tuple[str, str]:
    try:
        from rapidocr import RapidOCR  # type: ignore
    except Exception:
        return "", ""
    try:
        engine = RapidOCR()
        result = engine(str(image_path))
    except Exception:
        return "", ""

    lines: list[str] = []
    # RapidOCR versions return either a tuple-like result or an object with txts.
    texts = getattr(result, "txts", None)
    if texts:
        lines.extend(str(item) for item in texts if str(item).strip())
    elif isinstance(result, tuple) and result:
        for item in result[0] or []:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                text = item[1][0] if isinstance(item[1], (list, tuple)) else item[1]
                if str(text).strip():
                    lines.append(str(text).strip())
    elif isinstance(result, list):
        for item in result:
            if isinstance(item, (list, tuple)) and len(item) >= 2:
                text = item[1][0] if isinstance(item[1], (list, tuple)) else item[1]
                if str(text).strip():
                    lines.append(str(text).strip())
    return "\n".join(lines).strip(), "rapidocr" if lines else ""


def local_ocr(image_path: Path, messages: list[str]) -> tuple[str, str]:
    text, backend = rapidocr_text(image_path)
    if text:
        return text, backend
    text, backend = tesseract_ocr(image_path)
    if text:
        return text, backend
    messages.append("Local OCR backend is unavailable or returned no text; run bootstrap_deps.ps1 to install RapidOCR/Tesseract.")
    return "", ""


def image_to_data_url(image_path: Path) -> str:
    mime = "image/png" if image_path.suffix.lower() == ".png" else "image/jpeg"
    payload = base64.b64encode(image_path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{payload}"


def parse_openai_text(response: Any) -> str:
    output_text = getattr(response, "output_text", None)
    if output_text:
        return str(output_text).strip()
    pieces: list[str] = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            text = getattr(content, "text", None)
            if text:
                pieces.append(str(text))
    return "\n".join(pieces).strip()


def openai_visual_summary(image_path: Path, ocr_text: str, messages: list[str], enabled: bool) -> tuple[str, str]:
    if not enabled:
        return "", ""
    if not os.environ.get("OPENAI_API_KEY"):
        messages.append("OpenAI vision was skipped because OPENAI_API_KEY is not set.")
        return "", ""
    try:
        from openai import OpenAI  # type: ignore
    except Exception as exc:
        messages.append(f"OpenAI Python package is unavailable: {exc}")
        return "", ""

    model = os.environ.get("OFFICE_READER_VISION_MODEL", "gpt-4o")
    prompt = (
        "You are deep-reading a rendered Word/PowerPoint page. Extract any visible text, "
        "summarize diagrams/charts/screenshots, identify key conclusions, and mention uncertainty. "
        "Return concise Markdown with sections: Visible text, Visual summary, Diagram/chart reading, Gaps. "
        f"Local OCR text, if any:\n{ocr_text or '(none)'}"
    )
    try:
        client = OpenAI()
        response = client.responses.create(
            model=model,
            input=[
                {
                    "role": "user",
                    "content": [
                        {"type": "input_text", "text": prompt},
                        {"type": "input_image", "image_url": image_to_data_url(image_path), "detail": "high"},
                    ],
                }
            ],
        )
        text = parse_openai_text(response)
        return text, f"openai:{model}" if text else ""
    except Exception as exc:
        messages.append(f"OpenAI vision failed for {image_path.name}: {exc}")
        return "", ""


def cache_path(cache_dir: Path, source_hash: str, mode: str, item_hash: str) -> Path:
    return cache_dir / f"{source_hash[:16]}-{mode}-{item_hash[:24]}.json"


def analyze_item(
    source_hash: str,
    mode: str,
    cache_dir: Path,
    item: dict[str, Any],
    messages: list[str],
    enable_openai: bool,
) -> dict[str, Any]:
    image_path = Path(item["image_path"]) if item.get("image_path") else None
    item_hash = sha256_file(image_path) if image_path and image_path.exists() else stable_hash(item)
    path = cache_path(cache_dir, source_hash, mode, item_hash)
    if path.exists():
        cached = read_json(path)
        cached["cache_hit"] = True
        return cached

    started = time.perf_counter()
    result = {key: item.get(key, value) for key, value in VISUAL_FIELDS.items()}
    result["reason"] = item.get("reason", "")
    result["requires_visual_review"] = item.get("requires_visual_review", True)
    backends: list[str] = []

    if image_path and image_path.exists():
        ocr_text, ocr_backend = local_ocr(image_path, messages)
        if ocr_text:
            result["ocr_text"] = ocr_text
            backends.append(ocr_backend)
        vision_text, vision_backend = openai_visual_summary(image_path, ocr_text, messages, enable_openai)
        if vision_text:
            result["vision_summary"] = vision_text
            result["diagram_summary"] = vision_text
            backends.append(vision_backend)
        result["image_path"] = str(image_path)
    else:
        media_count = item.get("media_count")
        media_hint = f"OOXML found {media_count} media item(s)." if media_count is not None else "OOXML found media/drawing references."
        result["vision_summary"] = media_hint + " Rendered image analysis was not available for this item."
        result["diagram_summary"] = "Potential diagram/screenshot/image content requires rendered-page or embedded-image review."
        backends.append("ooxml-media")

    result["backend"] = "+".join(item for item in backends if item) or "visual-unavailable"
    result["confidence"] = "medium" if result.get("ocr_text") or result.get("vision_summary") else "low"
    result["duration_ms"] = int((time.perf_counter() - started) * 1000)
    result["cache_hit"] = False
    write_json(path, result)
    return result


def page_index_from_image(path: Path) -> int:
    digits = "".join(ch for ch in path.stem if ch.isdigit())
    return int(digits) if digits else 0


def candidate_pages(manifest: dict[str, Any], pages: list[Path], mode: str) -> list[Path]:
    if mode == "complete":
        return pages
    max_pages = int(os.environ.get("OFFICE_READER_BALANCED_MAX_PAGES", "8"))
    doc_type = manifest.get("document_type")
    visual = manifest.get("visual_findings", [])
    if doc_type == "pptx":
        slide_indexes = {int(item["slide_index"]) for item in visual if item.get("requires_visual_review") and item.get("slide_index")}
        if slide_indexes:
            return [page for page in pages if page_index_from_image(page) in slide_indexes][:max_pages]
    if any(item.get("requires_visual_review") for item in visual):
        return pages[:max_pages]
    return []


def has_package_media(manifest: dict[str, Any]) -> bool:
    for item in manifest.get("visual_findings", []):
        for media in item.get("media", []) or []:
            if isinstance(media, str) and "/media/" in media.replace("\\", "/"):
                return True
    return False


def append_visual_markdown(manifest: dict[str, Any]) -> None:
    artifacts = manifest.get("artifacts", {})
    full_path = Path(artifacts.get("full_markdown", ""))
    if not full_path.exists():
        return
    base = full_path.read_text(encoding="utf-8")
    if VISUAL_MARKER in base:
        base = base.split(VISUAL_MARKER, 1)[0].rstrip() + "\n"
    lines = ["", VISUAL_MARKER, "## Visual Deep Read", ""]
    findings = [item for item in manifest.get("visual_findings", []) if item.get("ocr_text") or item.get("vision_summary")]
    if not findings:
        lines.append("No visual OCR or vision summary was produced.")
    for item in findings[:80]:
        location = f"page {item.get('page_index')}" if item.get("page_index") else f"slide {item.get('slide_index')}" if item.get("slide_index") else "visual item"
        lines.append(f"### {location}")
        if item.get("ocr_text"):
            lines.extend(["", "OCR text:", "", str(item["ocr_text"])])
        if item.get("vision_summary"):
            lines.extend(["", "Visual summary:", "", str(item["vision_summary"])])
        if item.get("diagram_summary") and item.get("diagram_summary") != item.get("vision_summary"):
            lines.extend(["", "Diagram summary:", "", str(item["diagram_summary"])])
        lines.append("")
    full_path.write_text(base.rstrip() + "\n" + "\n".join(lines).rstrip() + "\n", encoding="utf-8")


def clamp_score(value: float) -> int:
    return max(0, min(100, int(round(value))))


def has_verified_visual_read(finding: dict[str, Any]) -> bool:
    backend = str(finding.get("backend", "")).lower()
    if finding.get("ocr_text") and any(name in backend for name in ("rapidocr", "tesseract")):
        return True
    if finding.get("vision_summary") and "openai:" in backend:
        return True
    if finding.get("diagram_summary") and "openai:" in backend:
        return True
    return False


def compute_completeness_score(manifest: dict[str, Any]) -> dict[str, Any]:
    structure = manifest.get("structure", []) or []
    tables = manifest.get("tables", []) or []
    visual = manifest.get("visual_findings", []) or []
    visual_analysis = manifest.get("visual_analysis", {}) or {}
    signals: list[str] = []

    if structure:
        readable = sum(1 for item in structure if (item.get("text") or item.get("title") or "").strip())
        text_coverage = clamp_score((readable / len(structure)) * 100)
    else:
        text_coverage = 0
        signals.append("No readable body or slide structure was extracted.")

    if tables:
        readable_tables = sum(1 for table in tables if any(any(str(cell).strip() for cell in row) for row in table.get("rows", [])))
        table_coverage = clamp_score((readable_tables / len(tables)) * 100)
    else:
        table_coverage = 100

    required_visual = [item for item in visual if item.get("requires_visual_review")]
    verified_visual = [item for item in required_visual if has_verified_visual_read(item)]
    if required_visual:
        visual_coverage = clamp_score((len(verified_visual) / len(required_visual)) * 100)
    else:
        visual_coverage = 100
    unverified_visual_count = len(required_visual) - len(verified_visual)

    ocr_items = [
        item
        for item in required_visual
        if item.get("ocr_text") and any(name in str(item.get("backend", "")).lower() for name in ("rapidocr", "tesseract"))
    ]
    ocr_confidence = clamp_score((len(ocr_items) / len(required_visual)) * 100) if required_visual else 100
    openai_enabled = any("openai:" in str(item.get("backend", "")).lower() for item in visual)

    status = visual_analysis.get("status")
    if status in {"skipped", "partial"} and required_visual:
        signals.append(f"Visual analysis was {status}.")
    if unverified_visual_count:
        signals.append(f"{unverified_visual_count} visual item(s) still need OCR or vision confirmation.")
    if required_visual and not openai_enabled:
        signals.append("OpenAI vision was not used for visual interpretation.")

    overall = clamp_score((text_coverage * 0.35) + (table_coverage * 0.15) + (visual_coverage * 0.40) + (ocr_confidence * 0.10))
    return {
        "overall": overall,
        "text_coverage": text_coverage,
        "table_coverage": table_coverage,
        "visual_coverage": visual_coverage,
        "ocr_confidence": ocr_confidence,
        "openai_vision_enabled": openai_enabled,
        "unverified_visual_count": unverified_visual_count,
        "signals": signals,
    }


def enrich_manifest(
    manifest_path: Path,
    normalized_file: Path,
    out_dir: Path,
    mode: str,
    enable_openai: bool,
    timeout_seconds: int,
) -> dict[str, Any]:
    manifest = read_json(manifest_path)
    manifest["reading_mode"] = mode
    visual = manifest.setdefault("visual_findings", [])
    if not visual:
        visual.append({"requires_visual_review": False, "reason": "no visual findings recorded"})
    for item in visual:
        ensure_visual_fields(item)

    messages: list[str] = []
    analysis = {
        "status": "skipped",
        "mode": mode,
        "rendered_page_count": 0,
        "analyzed_item_count": 0,
        "cache_hits": 0,
        "embedded_media_count": 0,
        "messages": messages,
        "backends": [],
    }
    cache_dir = out_dir / ".office-reader-cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    embedded_media = extract_embedded_media(normalized_file, manifest, out_dir, cache_dir, messages)
    if embedded_media:
        manifest.setdefault("artifacts", {})["embedded_media_dir"] = str(out_dir / "embedded_media")
        manifest["embedded_media"] = embedded_media
        analysis["embedded_media_count"] = len(embedded_media)

    if mode == "fast":
        messages.append("Visual page rendering skipped in fast mode; OOXML media hints were preserved.")
        for item in visual:
            item["backend"] = item.get("backend") or ("ooxml-media" if item.get("requires_visual_review") else "ooxml")
            item["vision_summary"] = item.get("vision_summary") or item.get("reason", "")
        analysis["status"] = "skipped"
        manifest["visual_analysis"] = analysis
        manifest["completeness_score"] = compute_completeness_score(manifest)
        write_json(manifest_path, manifest)
        append_visual_markdown(manifest)
        return manifest

    source_hash = sha256_file(normalized_file)
    page_images: list[Path] = []
    preview_result: dict[str, Any] = {}
    should_render = mode == "complete" or has_package_media(manifest) or manifest.get("document_type") == "pptx"
    if should_render:
        pdf_path, preview_result = render_preview(normalized_file, out_dir, timeout_seconds, messages)
        if pdf_path:
            max_pages = None if mode == "complete" else int(os.environ.get("OFFICE_READER_BALANCED_MAX_PAGES", "8"))
            page_images = render_pdf_pages(pdf_path, out_dir, messages, max_pages=max_pages)
            manifest.setdefault("artifacts", {})["preview_pdf"] = str(pdf_path)
            manifest["artifacts"]["page_images"] = [str(path) for path in page_images]
        elif preview_result.get("messages"):
            messages.extend(str(item) for item in preview_result.get("messages", []) if str(item) not in messages)
    else:
        messages.append("Preview rendering skipped because OOXML referenced drawings/media but no packaged media files were present.")

    analysis["rendered_page_count"] = len(page_images)
    work_items: list[dict[str, Any]] = []
    for page in candidate_pages(manifest, page_images, mode):
        work_items.append(
            {
                "page_index": page_index_from_image(page),
                "slide_index": page_index_from_image(page) if manifest.get("document_type") == "pptx" else None,
                "image_path": str(page),
                "reason": "rendered page selected for visual deep read",
                "requires_visual_review": True,
            }
        )

    if not work_items:
        for item in visual:
            if item.get("requires_visual_review"):
                work_items.append(item)

    enriched: list[dict[str, Any]] = []
    for item in work_items:
        ensure_visual_fields(item)
        result = analyze_item(source_hash, mode, cache_dir, item, messages, enable_openai)
        if result.get("cache_hit"):
            analysis["cache_hits"] += 1
        enriched.append(result)

    if page_images and enriched:
        for original in visual:
            ensure_visual_fields(original)
            if original.get("requires_visual_review") and not original.get("vision_summary"):
                original["vision_summary"] = "Rendered page-level visual findings were added below."
                original["backend"] = original.get("backend") or "rendered-page-analysis"
        visual.extend(enriched)
    elif enriched:
        for index, result in enumerate(enriched):
            if index < len(visual):
                visual[index].update(result)
            else:
                visual.append(result)

    analysis["analyzed_item_count"] = len(enriched)
    analysis["backends"] = sorted({item.get("backend", "") for item in visual if item.get("backend")})
    if enriched and page_images:
        analysis["status"] = "completed"
    elif enriched:
        analysis["status"] = "partial"
    else:
        analysis["status"] = "skipped"
        if not messages:
            messages.append("No visual work items were selected.")
    manifest["visual_analysis"] = analysis
    manifest.setdefault("artifacts", {})["visual_cache_dir"] = str(cache_dir)
    manifest["completeness_score"] = compute_completeness_score(manifest)
    write_json(manifest_path, manifest)
    append_visual_markdown(manifest)
    return manifest


def main() -> int:
    parser = argparse.ArgumentParser(description="Enrich an office-reader manifest with OCR and visual findings.")
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--normalized-file", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--mode", choices=["fast", "balanced", "complete"], default="balanced")
    parser.add_argument("--no-openai-vision", action="store_true")
    parser.add_argument("--timeout-seconds", type=int, default=90)
    args = parser.parse_args()

    try:
        enrich_manifest(
            args.manifest.resolve(),
            args.normalized_file.resolve(),
            args.out_dir.resolve(),
            args.mode,
            enable_openai=not args.no_openai_vision,
            timeout_seconds=args.timeout_seconds,
        )
    except Exception as exc:
        print(f"Visual analysis failed: {exc}", file=sys.stderr)
        return 1
    print(str(args.manifest.resolve()))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
