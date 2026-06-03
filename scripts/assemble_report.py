#!/usr/bin/env python3
"""Assemble a structured reading report from an office-reader manifest."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def first_sentence(text: str, limit: int = 240) -> str:
    text = " ".join(str(text or "").split())
    if len(text) <= limit:
        return text
    return text[: limit - 1].rstrip() + "..."


def bullet(label: str, value: Any) -> str:
    return f"- {label}: {value}"


def visual_object_summary(item: dict[str, Any]) -> str:
    parts = [str(item.get("object_type", "object"))]
    if item.get("name"):
        parts.append(f"name={item.get('name')}")
    if item.get("alt_text"):
        parts.append(f"alt={first_sentence(item.get('alt_text'), 80)}")
    if item.get("prog_id"):
        parts.append(f"prog_id={item.get('prog_id')}")
    if item.get("relationship_id"):
        parts.append(f"rel={item.get('relationship_id')}")
    if item.get("target"):
        parts.append(f"target={item.get('target')}")
    if item.get("relationships"):
        roles = ", ".join(
            f"{rel.get('role', 'rel')}={rel.get('target') or rel.get('relationship_id')}"
            for rel in item.get("relationships", [])[:6]
        )
        if roles:
            parts.append(f"relationships=({roles})")
    if item.get("geometry"):
        geometry = item.get("geometry", {})
        parts.append(
            "geometry="
            + ",".join(str(geometry.get(key, "")) for key in ("x", "y", "cx", "cy"))
        )
    return "; ".join(parts)


def build_report(manifest: dict[str, Any]) -> str:
    source = manifest.get("source", {})
    name = source.get("name") or Path(source.get("path", "document")).name
    structure = manifest.get("structure", [])
    tables = manifest.get("tables", [])
    comments = manifest.get("comments", [])
    revisions = manifest.get("revisions", [])
    notes = manifest.get("notes", [])
    visual = manifest.get("visual_findings", [])
    embedded_media = manifest.get("embedded_media", [])
    visual_analysis = manifest.get("visual_analysis", {})
    conversion = manifest.get("conversion", {})
    completeness = manifest.get("completeness_score", {})
    query = manifest.get("query", {})

    lines: list[str] = [f"# Structured Reading Report: {name}", ""]
    lines.extend(
        [
            "## Summary",
            bullet("Document type", manifest.get("document_type", "unknown")),
            bullet("Slides/sections", len(structure)),
            bullet("Tables", len(tables)),
            bullet("Comments", len(comments)),
            bullet("Revisions", len(revisions)),
            bullet("Speaker notes", len(notes)),
            bullet("Conversion status", conversion.get("status", "unknown")),
            bullet("Reading mode", manifest.get("reading_mode", "balanced")),
            bullet("Visual analysis", visual_analysis.get("status", "unknown")),
            bullet("Completeness score", f"{completeness.get('overall', 'unknown')}/100" if completeness else "unknown"),
            "",
        ]
    )

    lines.extend(["## Read Completeness", ""])
    if completeness:
        lines.append(bullet("Text coverage", f"{completeness.get('text_coverage', 0)}/100"))
        lines.append(bullet("Table coverage", f"{completeness.get('table_coverage', 0)}/100"))
        lines.append(bullet("Visual coverage", f"{completeness.get('visual_coverage', 0)}/100"))
        lines.append(bullet("OCR confidence", f"{completeness.get('ocr_confidence', 0)}/100"))
        lines.append(bullet("OpenAI vision enabled", completeness.get("openai_vision_enabled", False)))
        lines.append(bullet("Unverified visual items", completeness.get("unverified_visual_count", 0)))
        for signal in completeness.get("signals", [])[:8]:
            lines.append(f"- Signal: {first_sentence(signal)}")
    else:
        lines.append("- No completeness score was recorded.")
    lines.append("")

    if query:
        lines.extend(["## Query Results", ""])
        lines.append(bullet("Query", query.get("query", "")))
        lines.append(bullet("Matches", query.get("total_matches", 0)))
        if query.get("truncated"):
            lines.append("- Results were truncated; open the query artifact for the full available set.")
        for match in query.get("matches", [])[:20]:
            location = match.get("location", {}) or {}
            location_bits = [f"{key}={value}" for key, value in location.items() if value not in (None, "")]
            where = f" ({', '.join(location_bits)})" if location_bits else ""
            lines.append(f"- {match.get('source_type', 'match')}{where}: {first_sentence(match.get('text', ''), 320)}")
        if not query.get("matches"):
            lines.append("- No matches found in extracted text, comments, revisions, notes, tables, or visual text fields.")
        lines.append("")

    lines.extend(["## Outline", ""])
    if structure:
        for item in structure[:80]:
            label = "Slide" if item.get("type") == "slide" else item.get("type", "section").title()
            index = item.get("index", "")
            title = item.get("title") or item.get("text") or "Untitled"
            lines.append(f"- {label} {index}: {first_sentence(title)}")
    else:
        lines.append("- No outline entries extracted.")
    lines.append("")

    lines.extend(["## Table Index", ""])
    if tables:
        for table in tables:
            location = f"slide {table.get('slide_index')}" if table.get("slide_index") else "document"
            row_count = len(table.get("rows", []))
            preview = table.get("rows", [[]])[0] if table.get("rows") else []
            lines.append(f"- Table {table.get('index')}: {location}, {row_count} rows, columns: {', '.join(preview)}")
    else:
        lines.append("- No tables extracted.")
    lines.append("")

    lines.extend(["## Comments And Revisions", ""])
    if comments:
        for comment in comments[:80]:
            location = f"slide {comment.get('slide_index')}" if comment.get("slide_index") else f"id {comment.get('id', '')}".strip()
            anchor = f" on '{first_sentence(comment.get('anchor_text', ''), 120)}'" if comment.get("anchor_text") else ""
            lines.append(f"- Comment {location}{anchor}: {first_sentence(comment.get('text', ''))}")
    else:
        lines.append("- No comments extracted.")
    if revisions:
        for revision in revisions[:80]:
            lines.append(f"- {revision.get('type', 'revision').title()}: {first_sentence(revision.get('text', ''))}")
    else:
        lines.append("- No tracked insertions/deletions extracted.")
    lines.append("")

    lines.extend(["## Speaker Notes", ""])
    if notes:
        for note in notes[:80]:
            lines.append(f"- Slide {note.get('slide_index')}: {first_sentence(note.get('text', ''))}")
    else:
        lines.append("- No speaker notes extracted.")
    lines.append("")

    lines.extend(["## Visual Findings", ""])
    if visual:
        for finding in visual:
            status = "requires visual review" if finding.get("requires_visual_review") else "no visual review flag"
            location_parts = []
            if finding.get("page_index"):
                location_parts.append(f"page {finding.get('page_index')}")
            if finding.get("slide_index"):
                location_parts.append(f"slide {finding.get('slide_index')}")
            location = f" ({', '.join(location_parts)})" if location_parts else ""
            backend = f", backend: {finding.get('backend')}" if finding.get("backend") else ""
            lines.append(f"- {status}{location}: {finding.get('reason', 'no reason recorded')}{backend}")
            for rel in finding.get("relationships", [])[:30]:
                detail = []
                if rel.get("target"):
                    detail.append(f"target={rel.get('target')}")
                if rel.get("caption"):
                    detail.append(f"caption={first_sentence(rel.get('caption'), 120)}")
                if rel.get("nearest_heading"):
                    detail.append(f"section={first_sentence(rel.get('nearest_heading'), 120)}")
                if rel.get("nearby_text_before"):
                    detail.append(f"before={first_sentence(rel.get('nearby_text_before'), 120)}")
                if rel.get("nearby_text_after"):
                    detail.append(f"after={first_sentence(rel.get('nearby_text_after'), 120)}")
                if detail:
                    lines.append(f"  Media context: {'; '.join(detail)}")
            for item in finding.get("objects", [])[:20]:
                lines.append(f"  Object: {visual_object_summary(item)}")
    else:
        lines.append("- No visual findings recorded.")
    lines.append("")

    lines.extend(["## Embedded Media", ""])
    if embedded_media:
        for item in embedded_media[:80]:
            preview = f", preview={item.get('preview_path')}" if item.get("preview_path") else ""
            cache = "cache hit" if item.get("cache_hit") else "extracted"
            contexts = item.get("contexts", []) or []
            label = str(item.get("label", "")).strip()
            if contexts:
                first = contexts[0]
                label = label or first.get("caption") or first.get("alt_text") or first.get("title") or first.get("name") or first.get("nearest_heading") or ""
            label_text = f", label={first_sentence(label, 120)}" if label else ""
            lines.append(f"- {item.get('member')}: {item.get('content_type')} {cache}{label_text}, path={item.get('path')}{preview}")
            if item.get("ocr_text"):
                backend = f" via {item.get('ocr_backend')}" if item.get("ocr_backend") else ""
                lines.append(f"  Media OCR{backend}: {first_sentence(item.get('ocr_text'), 320)}")
    else:
        lines.append("- No embedded media files were extracted.")
    if manifest.get("artifacts", {}).get("media_contact_sheet"):
        lines.append(f"- Contact sheet: {manifest['artifacts']['media_contact_sheet']}")
    if manifest.get("artifacts", {}).get("media_summary"):
        lines.append(f"- Media summary: {manifest['artifacts']['media_summary']}")
    lines.append("")

    lines.extend(["## Visual Deep Read", ""])
    deep_items = [item for item in visual if item.get("ocr_text") or item.get("vision_summary") or item.get("diagram_summary")]
    if deep_items:
        for item in deep_items[:80]:
            location = f"page {item.get('page_index')}" if item.get("page_index") else f"slide {item.get('slide_index')}" if item.get("slide_index") else "visual item"
            backend = item.get("backend") or "unknown backend"
            lines.append(f"- {location} via {backend}, confidence {item.get('confidence', 'unknown')}:")
            if item.get("ocr_text"):
                lines.append(f"  OCR: {first_sentence(item.get('ocr_text', ''), 320)}")
            if item.get("vision_summary"):
                lines.append(f"  Visual: {first_sentence(item.get('vision_summary', ''), 420)}")
            if item.get("diagram_summary") and item.get("diagram_summary") != item.get("vision_summary"):
                lines.append(f"  Diagram: {first_sentence(item.get('diagram_summary', ''), 420)}")
    else:
        lines.append("- No OCR or vision summary was produced.")
    if visual_analysis:
        lines.append(f"- Rendered pages/slides: {visual_analysis.get('rendered_page_count', 0)}")
        lines.append(f"- Analyzed visual items: {visual_analysis.get('analyzed_item_count', 0)}")
        lines.append(f"- Cache hits: {visual_analysis.get('cache_hits', 0)}")
    lines.append("")

    lines.extend(["## Remaining Unverified Visual Gaps", ""])
    gaps = []
    if any(item.get("requires_visual_review") and not (item.get("ocr_text") or item.get("vision_summary")) for item in visual):
        gaps.append("Some media/drawing items were detected but not visually interpreted.")
    if visual_analysis.get("status") in {"skipped", "partial"}:
        gaps.append("Visual analysis was not complete; use --mode complete and install OCR/vision dependencies for deeper coverage.")
    for message in visual_analysis.get("messages", [])[:10]:
        if message:
            gaps.append(str(message))
    if not gaps:
        gaps.append("No unverified visual gaps were recorded by the visual pipeline.")
    lines.extend(f"- {gap}" for gap in gaps)
    lines.append("")

    lines.extend(["## Risks And Gaps", ""])
    risks = []
    if any(item.get("requires_visual_review") and not (item.get("ocr_text") or item.get("vision_summary")) for item in visual):
        risks.append("Embedded media or drawings may contain text or chart evidence that XML extraction cannot fully read.")
    if conversion.get("required") and conversion.get("status") != "success":
        risks.append("Legacy format conversion did not complete successfully.")
    if not structure:
        risks.append("No readable body/slide structure was extracted.")
    if not risks:
        risks.append("No blocking extraction risks detected by the manifest.")
    lines.extend(f"- {risk}" for risk in risks)
    lines.append("")

    artifacts = manifest.get("artifacts", {})
    lines.extend(["## Artifacts", ""])
    for key, value in artifacts.items():
        lines.append(f"- {key}: {value}")
    if not artifacts:
        lines.append("- No artifacts recorded.")
    lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="Build a structured report from an office-reader manifest.")
    parser.add_argument("manifest", type=Path)
    parser.add_argument("--out", type=Path, default=None)
    args = parser.parse_args()

    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    report = build_report(manifest)
    out = args.out or args.manifest.with_name(args.manifest.name.replace(".manifest.json", ".report.md"))
    out.write_text(report, encoding="utf-8")
    print(str(out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
