#!/usr/bin/env python3
"""Extract PPTX slides, notes, comments, tables, media hints, and Markdown."""

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
    atomic_write_text,
    collect_plain_text,
    core_properties,
    default_manifest,
    markdown_table,
    md_heading,
    media_members,
    qn,
    read_optional_xml,
    read_relationships,
    relationship_part_path,
    select_artifact_output_dir,
    table_to_rows,
    write_json,
)


def slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def presentation_slide_paths(package: zipfile.ZipFile, warnings: list[str]) -> list[str]:
    root = read_optional_xml(package, "ppt/presentation.xml", warnings)
    rels = read_relationships(package, "ppt/_rels/presentation.xml.rels", warnings)
    package_members = set(package.namelist())
    fallback = sorted(
        [name for name in package_members if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
        key=slide_number,
    )
    if root is None:
        if fallback:
            warnings.append("ppt/presentation.xml was missing or unreadable; used sorted slide members as a fallback.")
        else:
            raise ValueError("ppt/presentation.xml is missing or unreadable and no fallback slide XML parts were found")
        return fallback
    paths: list[str] = []
    for sld_id in root.findall(".//p:sldId", NS):
        rel_id = sld_id.attrib.get(qn("r", "id"))
        rel = rels.get(rel_id or "")
        if rel and rel.get("target"):
            part = relationship_part_path("ppt/presentation.xml", rel, warnings)
            if part and part in package_members:
                paths.append(part)
            elif part:
                warnings.append(f"Referenced slide part was not found in PPTX package: {part}")
    if not paths and fallback:
        warnings.append("Presentation slide relationships were missing, unreadable, or invalid; used sorted slide members as a fallback.")
        return fallback
    return paths


def slide_title(texts: list[str], index: int) -> str:
    for text in texts:
        if text:
            return text
    return f"Slide {index}"


def slide_text_paragraphs(root) -> list[str]:
    texts: list[str] = []
    for paragraph in root.findall(".//p:txBody/a:p", NS):
        text = collect_plain_text(paragraph)
        if text:
            texts.append(text)
    return texts


def notes_body_text(root) -> str:
    body_texts: list[str] = []
    for shape in root.findall(".//p:sp", NS):
        placeholder = shape.find("./p:nvSpPr/p:nvPr/p:ph", NS)
        if placeholder is not None and placeholder.attrib.get("type") == "body":
            text = collect_plain_text(shape)
            if text:
                body_texts.append(text)
    return "\n".join(body_texts) if body_texts else collect_plain_text(root)


def extract_notes(
    package: zipfile.ZipFile,
    slide_path: str,
    slide_index: int,
    rels: dict[str, dict[str, str]],
    warnings: list[str],
) -> list[dict[str, Any]]:
    notes: list[dict[str, Any]] = []
    for rel in rels.values():
        if "notesSlide" not in rel.get("type", ""):
            continue
        part = relationship_part_path(slide_path, rel, warnings)
        if not part:
            continue
        root = read_optional_xml(package, part, warnings)
        text = notes_body_text(root) if root is not None else ""
        if text:
            notes.append({"slide_index": slide_index, "part": part, "text": text})
    return notes


def comment_relationship_parts(
    slide_path: str,
    rels: dict[str, dict[str, str]],
    warnings: list[str],
) -> list[str]:
    parts: list[str] = []
    for rel in rels.values():
        if "comment" not in rel.get("type", "").lower():
            continue
        part = relationship_part_path(slide_path, rel, warnings)
        if part:
            parts.append(part)
    return parts


def extract_comment_authors(package: zipfile.ZipFile, warnings: list[str]) -> dict[str, str]:
    root = read_optional_xml(package, "ppt/commentAuthors.xml", warnings)
    if root is None:
        return {}
    return {
        node.attrib["id"]: node.attrib.get("name", "")
        for node in root.findall("p:cmAuthor", NS)
        if node.attrib.get("id")
    }


def comment_node_text(node) -> str:
    parts = [child.text for child in node.iter() if local_name_fallback(child.tag) == "text" and child.text]
    return "".join(parts).strip() or collect_plain_text(node)


def extract_comment_parts(
    package: zipfile.ZipFile,
    comment_parts: list[str],
    slide_index: int | None,
    warnings: list[str],
    authors: dict[str, str],
) -> list[dict[str, Any]]:
    comments: list[dict[str, Any]] = []
    for part in comment_parts:
        root = read_optional_xml(package, part, warnings)
        if root is None:
            continue
        for node in root.iter():
            if node.tag in {qn("p", "cm"), qn("p", "cmAuthor")}:
                text = comment_node_text(node)
                if text:
                    author_id = node.attrib.get("authorId", "")
                    comments.append(
                        {
                            "slide_index": slide_index,
                            "part": part,
                            "text": text,
                            "date": node.attrib.get("dt", ""),
                            "author_id": author_id,
                            "author": authors.get(author_id, ""),
                        }
                    )
            elif local_name_fallback(node.tag) == "text" and node.text:
                comments.append({"slide_index": slide_index, "part": part, "text": node.text.strip(), "date": "", "author_id": "", "author": ""})
    seen = set()
    unique = []
    for item in comments:
        key = (item["slide_index"], item["part"], item["text"])
        if key not in seen:
            unique.append(item)
            seen.add(key)
    return unique


def local_name_fallback(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def slide_relationships_member(slide_path: str) -> str:
    return f"{Path(slide_path).parent.as_posix()}/_rels/{Path(slide_path).name}.rels"


def extract_pptx(source: Path, out_dir: Path) -> tuple[dict[str, Any], str]:
    manifest = default_manifest(source, "pptx")
    markdown: list[str] = []

    with zipfile.ZipFile(source) as package:
        manifest["metadata"] = core_properties(package, manifest["warnings"])
        slide_paths = presentation_slide_paths(package, manifest["warnings"])
        media = media_members(package, "ppt/media/")
        comment_authors = extract_comment_authors(package, manifest["warnings"])
        related_comment_parts: set[str] = set()

        for index, slide_path in enumerate(slide_paths, start=1):
            rels = read_relationships(package, slide_relationships_member(slide_path), manifest["warnings"])
            comment_parts = comment_relationship_parts(slide_path, rels, manifest["warnings"])
            related_comment_parts.update(comment_parts)
            root = read_optional_xml(package, slide_path, manifest["warnings"])
            if root is None:
                continue
            texts = slide_text_paragraphs(root)
            title = slide_title(texts, index)
            body_text = "\n".join(texts)
            notes = extract_notes(package, slide_path, index, rels, manifest["warnings"])
            comments = extract_comment_parts(package, comment_parts, index, manifest["warnings"], comment_authors)
            manifest["notes"].extend(notes)
            manifest["comments"].extend(comments)

            manifest["structure"].append(
                {
                    "type": "slide",
                    "index": index,
                    "part": slide_path,
                    "title": title,
                    "text": body_text,
                }
            )

            markdown.append(f"## Slide {index}: {title}")
            if len(texts) > 1:
                markdown.extend(texts[1:])
            elif texts:
                markdown.append(texts[0])

            slide_table_index = 0
            for table in root.iter(qn("a", "tbl")):
                slide_table_index += 1
                rows = table_to_rows(table, "a:tc", "a:tr")
                table_entry = {"slide_index": index, "index": len(manifest["tables"]) + 1, "rows": rows}
                manifest["tables"].append(table_entry)
                markdown.append(f"Table {table_entry['index']}")
                markdown.append(markdown_table(rows))

            for note in notes:
                markdown.append(f"Speaker notes: {note['text']}")
            for comment in comments:
                markdown.append(f"> Comment: {comment['text']}")

            slide_media = []
            for blip in root.iter(qn("a", "blip")):
                rel_id = blip.attrib.get(R_EMBED)
                rel = rels.get(rel_id or "")
                if rel:
                    slide_media.append({"relationship_id": rel_id or "", "target": rel.get("target", "")})
            if slide_media:
                manifest["visual_findings"].append(
                    {
                        "slide_index": index,
                        "requires_visual_review": True,
                        "reason": "slide contains embedded media references",
                        "media": slide_media,
                    }
                )

        if slide_paths and not manifest["structure"]:
            raise ValueError("PPTX package contains no readable slide XML parts")

        package_comment_parts = {
            name for name in package.namelist() if name.startswith("ppt/comments/") and name.endswith(".xml")
        }
        orphan_comments = extract_comment_parts(
            package,
            sorted(package_comment_parts - related_comment_parts),
            None,
            manifest["warnings"],
            comment_authors,
        )
        if orphan_comments:
            manifest["comments"].extend(orphan_comments)
            markdown.append("## Unattributed comments")
            for comment in orphan_comments:
                markdown.append(f"> Comment: {comment['text']}")

        if media and not manifest["visual_findings"]:
            manifest["visual_findings"].append(
                {
                    "requires_visual_review": True,
                    "reason": "presentation contains package media",
                    "media_count": len(media),
                    "media": media,
                }
            )
        elif not manifest["visual_findings"]:
            manifest["visual_findings"].append(
                {
                    "requires_visual_review": False,
                    "reason": "no embedded media detected in OOXML package",
                    "media_count": 0,
                    "media": [],
                }
            )

    title = manifest.get("metadata", {}).get("title") or source.stem
    full_md = "\n\n".join([md_heading(title, 1), *markdown]).strip() + "\n"
    full_path = out_dir / f"{source.stem}.full.md"
    manifest_path = out_dir / f"{source.stem}.manifest.json"
    manifest["artifacts"] = {"full_markdown": str(full_path), "manifest": str(manifest_path)}
    return manifest, full_md


def main() -> int:
    parser = argparse.ArgumentParser(description="Read a PPTX into Markdown and a structured manifest.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--out-dir", type=Path, default=None)
    args = parser.parse_args()

    source = args.source.resolve()
    if source.suffix.lower() != ".pptx":
        print(f"read_pptx.py expects a .pptx file, got {source.suffix}", file=sys.stderr)
        return 2
    out_dir = (args.out_dir or source.parent).resolve()
    out_dir = select_artifact_output_dir(source, out_dir, "pptx")

    try:
        manifest, full_md = extract_pptx(source, out_dir)
    except Exception as exc:
        print(f"Failed to read PPTX: {exc}", file=sys.stderr)
        return 1

    full_path = out_dir / f"{source.stem}.full.md"
    manifest_path = out_dir / f"{source.stem}.manifest.json"
    atomic_write_text(full_path, full_md)
    write_json(manifest_path, manifest)
    print(str(manifest_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
