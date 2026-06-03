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
    resolve_part_path,
    table_to_rows,
    write_json,
)


TEXTBOX_SKIP_NAMES = {"txbxContent"}


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


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def iter_without_subtrees(element, skip_subtree_names: set[str]):
    if local_name(element.tag) in skip_subtree_names:
        return
    yield element
    for child in list(element):
        yield from iter_without_subtrees(child, skip_subtree_names)


def extract_comment_refs(paragraph, skip_subtree_names: set[str] | None = None) -> list[str]:
    ids = []
    nodes = iter_without_subtrees(paragraph, skip_subtree_names or set())
    for ref in nodes:
        if ref.tag != qn("w", "commentReference"):
            continue
        ref_id = ref.attrib.get(W_ID)
        if ref_id is not None:
            ids.append(ref_id)
    return ids


def comment_anchor_texts(paragraph, skip_subtree_names: set[str] | None = None) -> dict[str, str]:
    anchors: dict[str, list[str]] = {}
    active: list[str] = []
    for node in iter_without_subtrees(paragraph, skip_subtree_names or set()):
        if node.tag == qn("w", "commentRangeStart"):
            ref_id = node.attrib.get(W_ID)
            if ref_id is not None:
                active.append(ref_id)
                anchors.setdefault(ref_id, [])
            continue
        if node.tag == qn("w", "commentRangeEnd"):
            ref_id = node.attrib.get(W_ID)
            if ref_id in active:
                active.remove(ref_id)
            continue
        if local_name(node.tag) in {"t", "delText"} and node.text and active:
            for ref_id in active:
                anchors.setdefault(ref_id, []).append(node.text)
    return {ref_id: "".join(parts).strip() for ref_id, parts in anchors.items() if "".join(parts).strip()}


def iter_blips(element, skip_subtree_names: set[str] | None = None):
    for node in iter_without_subtrees(element, skip_subtree_names or set()):
        if node.tag == qn("a", "blip"):
            yield node


def table_cell_location(table_index: int, row_index: int, cell_index: int) -> dict[str, int]:
    return {"table_index": table_index, "row_index": row_index, "cell_index": cell_index}


def part_relationships_member(part_path: str) -> str:
    part = Path(part_path)
    return f"{part.parent.as_posix()}/_rels/{part.name}.rels"


def resolved_relationship_target(part_path: str, rels: dict[str, dict[str, str]], rel_id: str | None) -> str:
    rel = rels.get(rel_id or "")
    target = rel.get("target", "") if rel else ""
    return resolve_part_path(part_path, target) if target else ""


def is_caption_text(text: str) -> bool:
    return bool(re.match(r"^\s*(图|表|Figure|Fig\.|Table)\s*[\d一二三四五六七八九十IVXivx\-\.]*", text or ""))


def nearest_structure_text(structure: list[dict[str, Any]], paragraph_index: int, direction: int) -> str:
    ordered = sorted(
        (item for item in structure if item.get("index") and str(item.get("text", "")).strip()),
        key=lambda item: int(item["index"]),
    )
    if direction < 0:
        candidates = [item for item in ordered if int(item["index"]) < paragraph_index]
        return str(candidates[-1]["text"]) if candidates else ""
    candidates = [item for item in ordered if int(item["index"]) > paragraph_index]
    return str(candidates[0]["text"]) if candidates else ""


def nearest_heading_text(structure: list[dict[str, Any]], paragraph_index: int) -> str:
    headings = [
        item
        for item in structure
        if item.get("type") == "heading" and item.get("index") and int(item["index"]) < paragraph_index
    ]
    return str(headings[-1].get("text", "")) if headings else ""


def add_media_context(media_refs: list[dict[str, Any]], structure: list[dict[str, Any]]) -> list[dict[str, Any]]:
    enriched: list[dict[str, Any]] = []
    by_index = {int(item["index"]): item for item in structure if item.get("index")}
    for ref in media_refs:
        item = dict(ref)
        paragraph_text = str(item.get("paragraph_text", ""))
        if paragraph_text and is_caption_text(paragraph_text):
            item.setdefault("caption", paragraph_text)
        paragraph_index = item.get("paragraph_index")
        if isinstance(paragraph_index, int):
            current = str(by_index.get(paragraph_index, {}).get("text", ""))
            before = nearest_structure_text(structure, paragraph_index, -1)
            after = nearest_structure_text(structure, paragraph_index, 1)
            item["nearby_text_before"] = before
            item["nearby_text_after"] = after
            item["nearest_heading"] = nearest_heading_text(structure, paragraph_index)
            for candidate in (current, after, before):
                if is_caption_text(candidate):
                    item["caption"] = candidate
                    break
        enriched.append(item)
    return enriched


def block_children(container) -> list[Any]:
    if container is None:
        return []
    if container.tag in {
        qn("w", "body"),
        qn("w", "hdr"),
        qn("w", "ftr"),
        qn("w", "footnote"),
        qn("w", "endnote"),
    }:
        return list(container)
    body = container.find("w:body", NS)
    return list(body) if body is not None else list(container)


def related_parts(package: zipfile.ZipFile, source_part: str, rels: dict[str, dict[str, str]], part_type: str) -> list[tuple[str, str, str]]:
    matches: list[tuple[str, str, str]] = []
    needle = f"/{part_type}"
    for rel_id, rel in rels.items():
        if needle not in rel.get("type", "") or not rel.get("target"):
            continue
        part = resolve_part_path(source_part, rel["target"])
        if part in package.namelist():
            matches.append((part_type, part, rel_id))
    return matches


def fixed_note_parts(package: zipfile.ZipFile) -> list[tuple[str, str, str]]:
    parts = []
    for part_type, part in (("footnote", "word/footnotes.xml"), ("endnote", "word/endnotes.xml")):
        if part in package.namelist():
            parts.append((part_type, part, ""))
    return parts


def textbox_blocks(container) -> list[Any]:
    blocks = []
    for node in container.iter():
        if node.tag.rsplit("}", 1)[-1] != "txbxContent":
            continue
        blocks.extend(block_children(node))
    return blocks


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

        def process_part(part_type: str, part_path: str, container, part_rels: dict[str, dict[str, str]], label: str = "") -> None:
            nonlocal paragraph_index, table_index
            added_part_heading = False
            def base_location(extra_location: dict[str, str] | None = None) -> dict[str, str]:
                location = {"part_type": part_type, "part": part_path}
                if extra_location:
                    location.update(extra_location)
                return location

            def ensure_part_heading() -> None:
                nonlocal added_part_heading
                if part_type != "document" and not added_part_heading:
                    markdown.append(md_heading(label or part_type.title(), 2))
                    added_part_heading = True

            def process_child(child, extra_location: dict[str, str] | None = None) -> None:
                nonlocal paragraph_index, table_index
                skip_names = set() if (extra_location or {}).get("container") == "textbox" else TEXTBOX_SKIP_NAMES
                tag = child.tag
                if tag == qn("w", "p"):
                    text, revisions = paragraph_text_with_revisions(child, skip_subtree_names=skip_names)
                    if not text and not list(child.iter(qn("w", "drawing"))):
                        pass
                    else:
                        ensure_part_heading()
                        paragraph_index += 1
                        level = heading_level(child) if part_type == "document" else None
                        refs = extract_comment_refs(child, skip_subtree_names=skip_names)
                        anchors = comment_anchor_texts(child, skip_subtree_names=skip_names)
                        location = base_location(extra_location)
                        item = {
                            "type": "heading" if level else "paragraph",
                            "index": paragraph_index,
                            "level": level,
                            "text": text,
                            "comment_ids": refs,
                            **location,
                        }
                        manifest["structure"].append(item)
                        for revision in revisions:
                            revision.update({"paragraph_index": paragraph_index, **location})
                            manifest["revisions"].append(revision)

                        if level:
                            markdown.append(md_heading(text, level))
                        elif text:
                            markdown.append(text)
                        for ref_id in refs:
                            comment = comments_by_id.get(ref_id)
                            if comment:
                                comment.update({"paragraph_index": paragraph_index, **location})
                                if anchors.get(ref_id):
                                    comment["anchor_text"] = anchors[ref_id]
                                markdown.append(f"> Comment {ref_id}: {comment['text']}")

                        for blip in iter_blips(child, skip_subtree_names=skip_names):
                            rel_id = blip.attrib.get(R_EMBED)
                            target = resolved_relationship_target(part_path, part_rels, rel_id)
                            if target:
                                media_refs.append(
                                    {
                                        "relationship_id": rel_id or "",
                                        "target": target,
                                        "paragraph_index": paragraph_index,
                                        "paragraph_text": text,
                                        **location,
                                    }
                                )
                    for textbox_child in textbox_blocks(child):
                        textbox_location = {**(extra_location or {}), "container": "textbox"}
                        process_child(textbox_child, textbox_location)

                elif tag == qn("w", "tbl"):
                    ensure_part_heading()
                    table_index += 1
                    rows = table_to_rows(child, "w:tc", "w:tr")
                    part_location = base_location(extra_location)
                    table_entry = {"index": table_index, "rows": rows, **part_location}
                    manifest["tables"].append(table_entry)
                    markdown.append(f"Table {table_index}")
                    markdown.append(markdown_table(rows))
                    for row_index, row in enumerate(child.findall("w:tr", NS), start=1):
                        for cell_index, cell in enumerate(row.findall("w:tc", NS), start=1):
                            location = {**table_cell_location(table_index, row_index, cell_index), **part_location}
                            for paragraph in cell.findall(".//w:p", NS):
                                _text, revisions = paragraph_text_with_revisions(paragraph, skip_subtree_names=TEXTBOX_SKIP_NAMES)
                                for revision in revisions:
                                    revision.update(location)
                                    manifest["revisions"].append(revision)
                                anchors = comment_anchor_texts(paragraph, skip_subtree_names=TEXTBOX_SKIP_NAMES)
                                for ref_id in extract_comment_refs(paragraph, skip_subtree_names=TEXTBOX_SKIP_NAMES):
                                    comment = comments_by_id.get(ref_id)
                                    if comment:
                                        comment.update(location)
                                        if anchors.get(ref_id):
                                            comment["anchor_text"] = anchors[ref_id]
                                for blip in iter_blips(paragraph, skip_subtree_names=TEXTBOX_SKIP_NAMES):
                                    rel_id = blip.attrib.get(R_EMBED)
                                    target = resolved_relationship_target(part_path, part_rels, rel_id)
                                    if target:
                                        media_refs.append(
                                            {
                                                "relationship_id": rel_id or "",
                                                "target": target,
                                                "paragraph_text": _text,
                                                **location,
                                            }
                                        )
                                for textbox_child in textbox_blocks(paragraph):
                                    textbox_location = {**location, "container": "textbox"}
                                    process_child(textbox_child, textbox_location)
                elif tag == qn("w", "sdt"):
                    sdt_content = child.find("w:sdtContent", NS)
                    if sdt_content is not None:
                        nested_location = {**(extra_location or {}), "container": "content_control"}
                        for nested_child in block_children(sdt_content):
                            process_child(nested_child, nested_location)

            for child in block_children(container):
                process_child(child)

        process_part("document", "word/document.xml", body, rels)

        supplemental_parts = []
        supplemental_parts.extend(related_parts(package, "word/document.xml", rels, "header"))
        supplemental_parts.extend(related_parts(package, "word/document.xml", rels, "footer"))
        supplemental_parts.extend(fixed_note_parts(package))
        for part_type, part_path, rel_id in supplemental_parts:
            part_root = read_xml(package, part_path)
            if part_root is None:
                continue
            part_rels = read_relationships(package, part_relationships_member(part_path))
            if part_type in {"footnote", "endnote"}:
                for note in part_root.findall(f"w:{part_type}", NS):
                    note_id = note.attrib.get(W_ID, "")
                    if note_id in {"-1", "0"}:
                        continue
                    process_part(part_type, part_path, note, part_rels, f"{part_type.title()} {note_id}".strip())
            else:
                process_part(part_type, part_path, part_root, part_rels, f"{part_type.title()} {rel_id}".strip())

        package_media = media_members(package, "word/media/")
        media_refs = add_media_context(media_refs, manifest["structure"])
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
