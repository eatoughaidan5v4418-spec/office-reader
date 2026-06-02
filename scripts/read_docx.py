#!/usr/bin/env python3
"""Extract DOCX structure, comments, revisions, tables, media hints, and Markdown."""

from __future__ import annotations

import argparse
import re
import sys
import zipfile
from pathlib import Path
from typing import Any

from common_ooxml import (
    NS,
    R_EMBED,
    W_ID,
    W_VAL,
    collect_plain_text,
    core_properties,
    default_manifest,
    markdown_table,
    md_heading,
    media_members,
    paragraph_text_with_revisions,
    qn,
    read_relationships,
    read_xml,
    table_to_rows,
    write_json,
)


def heading_level(paragraph) -> int | None:
    style = paragraph.find("w:pPr/w:pStyle", NS)
    if style is None:
        return None
    value = style.attrib.get(W_VAL, "")
    match = re.search(r"Heading\s*([1-6])|Heading([1-6])", value, re.IGNORECASE)
    if match:
        return int(match.group(1) or match.group(2))
    return None


def extract_comments(package: zipfile.ZipFile) -> list[dict[str, Any]]:
    root = read_xml(package, "word/comments.xml")
    if root is None:
        return []
    comments = []
    for node in root.findall("w:comment", NS):
        text = collect_plain_text(node)
        if text:
            comments.append(
                {
                    "id": node.attrib.get(W_ID, ""),
                    "author": node.attrib.get(qn("w", "author"), ""),
                    "date": node.attrib.get(qn("w", "date"), ""),
                    "text": text,
                }
            )
    return comments


def extract_comment_refs(paragraph) -> list[str]:
    ids = []
    for ref in paragraph.iter(qn("w", "commentReference")):
        ref_id = ref.attrib.get(W_ID)
        if ref_id is not None:
            ids.append(ref_id)
    return ids


def table_cell_location(table_index: int, row_index: int, cell_index: int) -> dict[str, int]:
    return {"table_index": table_index, "row_index": row_index, "cell_index": cell_index}


def extract_docx(source: Path, out_dir: Path) -> tuple[dict[str, Any], str]:
    manifest = default_manifest(source, "docx")
    markdown: list[str] = []

    with zipfile.ZipFile(source) as package:
        manifest["metadata"] = core_properties(package)
        manifest["comments"] = extract_comments(package)
        comments_by_id = {item["id"]: item for item in manifest["comments"]}
        rels = read_relationships(package, "word/_rels/document.xml.rels")
        root = read_xml(package, "word/document.xml")
        if root is None:
            raise ValueError("word/document.xml is missing from DOCX package")
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("word/document.xml has no word body")

        paragraph_index = 0
        table_index = 0
        media_refs: list[dict[str, str]] = []

        for child in list(body):
            tag = child.tag
            if tag == qn("w", "p"):
                text, revisions = paragraph_text_with_revisions(child)
                if not text and not list(child.iter(qn("w", "drawing"))):
                    continue
                paragraph_index += 1
                level = heading_level(child)
                refs = extract_comment_refs(child)
                item = {
                    "type": "heading" if level else "paragraph",
                    "index": paragraph_index,
                    "level": level,
                    "text": text,
                    "comment_ids": refs,
                }
                manifest["structure"].append(item)
                for revision in revisions:
                    revision["paragraph_index"] = paragraph_index
                    manifest["revisions"].append(revision)

                if level:
                    markdown.append(md_heading(text, level))
                elif text:
                    markdown.append(text)
                for ref_id in refs:
                    comment = comments_by_id.get(ref_id)
                    if comment:
                        markdown.append(f"> Comment {ref_id}: {comment['text']}")

                for blip in child.iter(qn("a", "blip")):
                    rel_id = blip.attrib.get(R_EMBED)
                    rel = rels.get(rel_id or "")
                    target = rel.get("target", "") if rel else ""
                    if target:
                        media_refs.append({"relationship_id": rel_id or "", "target": target})

            elif tag == qn("w", "tbl"):
                table_index += 1
                rows = table_to_rows(child, "w:tc", "w:tr")
                table_entry = {"index": table_index, "rows": rows}
                manifest["tables"].append(table_entry)
                markdown.append(f"Table {table_index}")
                markdown.append(markdown_table(rows))
                for row_index, row in enumerate(child.findall("w:tr", NS), start=1):
                    for cell_index, cell in enumerate(row.findall("w:tc", NS), start=1):
                        location = table_cell_location(table_index, row_index, cell_index)
                        for paragraph in cell.findall(".//w:p", NS):
                            _text, revisions = paragraph_text_with_revisions(paragraph)
                            for revision in revisions:
                                revision.update(location)
                                manifest["revisions"].append(revision)
                            for ref_id in extract_comment_refs(paragraph):
                                comment = comments_by_id.get(ref_id)
                                if comment:
                                    comment.update(location)
                            for blip in paragraph.iter(qn("a", "blip")):
                                rel_id = blip.attrib.get(R_EMBED)
                                rel = rels.get(rel_id or "")
                                target = rel.get("target", "") if rel else ""
                                if target:
                                    media_refs.append({"relationship_id": rel_id or "", "target": target, **location})

        package_media = media_members(package, "word/media/")
        if package_media or media_refs:
            manifest["visual_findings"].append(
                {
                    "requires_visual_review": True,
                    "reason": "document contains embedded media or drawing references",
                    "media_count": len(package_media) or len(media_refs),
                    "media": package_media,
                    "relationships": media_refs,
                }
            )
        else:
            manifest["visual_findings"].append(
                {
                    "requires_visual_review": False,
                    "reason": "no embedded media detected in OOXML package",
                    "media_count": 0,
                    "media": [],
                    "relationships": [],
                }
            )

    title = manifest.get("metadata", {}).get("title") or source.stem
    full_md = "\n\n".join([md_heading(title, 1), *markdown]).strip() + "\n"
    full_path = out_dir / f"{source.stem}.full.md"
    manifest_path = out_dir / f"{source.stem}.manifest.json"
    manifest["artifacts"] = {"full_markdown": str(full_path), "manifest": str(manifest_path)}
    return manifest, full_md


def main() -> int:
    parser = argparse.ArgumentParser(description="Read a DOCX into Markdown and a structured manifest.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out-dir", type=Path, default=None)
    args = parser.parse_args()

    source = args.source.resolve()
    if source.suffix.lower() != ".docx":
        print(f"read_docx.py expects a .docx file, got {source.suffix}", file=sys.stderr)
        return 2
    out_dir = (args.out_dir or source.parent).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        manifest, full_md = extract_docx(source, out_dir)
    except Exception as exc:
        print(f"Failed to read DOCX: {exc}", file=sys.stderr)
        return 1

    full_path = out_dir / f"{source.stem}.full.md"
    manifest_path = out_dir / f"{source.stem}.manifest.json"
    full_path.write_text(full_md, encoding="utf-8")
    write_json(manifest_path, manifest)
    print(str(manifest_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
