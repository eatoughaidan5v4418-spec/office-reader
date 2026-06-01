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


def build_report(manifest: dict[str, Any]) -> str:
    source = manifest.get("source", {})
    name = source.get("name") or Path(source.get("path", "document")).name
    structure = manifest.get("structure", [])
    tables = manifest.get("tables", [])
    comments = manifest.get("comments", [])
    revisions = manifest.get("revisions", [])
    notes = manifest.get("notes", [])
    visual = manifest.get("visual_findings", [])
    visual_analysis = manifest.get("visual_analysis", {})
    conversion = manifest.get("conversion", {})

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
            "",
        ]
    )

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
            lines.append(f"- Comment {location}: {first_sentence(comment.get('text', ''))}")
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
    else:
        lines.append("- No visual findings recorded.")
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
