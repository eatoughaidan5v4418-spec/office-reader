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
    collect_plain_text,
    core_properties,
    default_manifest,
    markdown_table,
    md_heading,
    media_members,
    qn,
    read_relationships,
    read_xml,
    resolve_part_path,
    table_to_rows,
    write_json,
)


R_LINK = qn("r", "link")
R_ID = qn("r", "id")
DIAGRAM_REL_ATTRS = {
    qn("r", "dm"): "data_model",
    qn("r", "lo"): "layout",
    qn("r", "qs"): "quick_style",
    qn("r", "cs"): "colors",
}


def slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def presentation_slide_paths(package: zipfile.ZipFile) -> list[str]:
    root = read_xml(package, "ppt/presentation.xml")
    rels = read_relationships(package, "ppt/_rels/presentation.xml.rels")
    if root is None:
        return sorted(
            [name for name in package.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=slide_number,
        )
    paths: list[str] = []
    for sld_id in root.findall(".//p:sldId", NS):
        rel_id = sld_id.attrib.get(qn("r", "id"))
        rel = rels.get(rel_id or "")
        if rel and rel.get("target"):
            paths.append(resolve_part_path("ppt/presentation.xml", rel["target"]))
    return paths


def slide_title(texts: list[str], index: int) -> str:
    for text in texts:
        if text:
            return text
    return f"Slide {index}"


def extract_notes(package: zipfile.ZipFile, slide_path: str, slide_index: int, rels: dict[str, dict[str, str]]) -> list[dict[str, Any]]:
    notes: list[dict[str, Any]] = []
    for rel in rels.values():
        if "notesSlide" not in rel.get("type", ""):
            continue
        part = resolve_part_path(slide_path, rel["target"])
        root = read_xml(package, part)
        text = collect_plain_text(root)
        if text:
            notes.append({"slide_index": slide_index, "part": part, "text": text})
    return notes


def comment_authors(package: zipfile.ZipFile) -> dict[str, dict[str, str]]:
    root = read_xml(package, "ppt/commentAuthors.xml")
    if root is None:
        return {}
    authors: dict[str, dict[str, str]] = {}
    for node in root.iter(qn("p", "cmAuthor")):
        author_id = node.attrib.get("id")
        if not author_id:
            continue
        authors[author_id] = {
            "author": node.attrib.get("name", ""),
            "initials": node.attrib.get("initials", ""),
        }
    return authors


def extract_comments(
    package: zipfile.ZipFile,
    slide_path: str,
    slide_index: int,
    rels: dict[str, dict[str, str]],
    authors: dict[str, dict[str, str]] | None = None,
) -> list[dict[str, Any]]:
    comments: list[dict[str, Any]] = []
    authors = authors or {}
    comment_parts = []
    for rel in rels.values():
        if "comment" in rel.get("type", "").lower():
            comment_parts.append(resolve_part_path(slide_path, rel["target"]))
    if not comment_parts:
        comment_parts = sorted(
            [name for name in package.namelist() if name.startswith("ppt/comments/") and name.endswith(".xml")]
        )
    for part in comment_parts:
        root = read_xml(package, part)
        if root is None:
            continue
        for node in root.iter():
            if node.tag in {qn("p", "cm"), qn("p", "cmAuthor")}:
                text = collect_plain_text(node)
                if not text:
                    text_node = node.find("p:text", NS)
                    text = text_node.text.strip() if text_node is not None and text_node.text else ""
                if text:
                    author_id = node.attrib.get("authorId", "")
                    author = authors.get(author_id, {})
                    comment = {
                        "slide_index": slide_index,
                        "part": part,
                        "text": text,
                        "date": node.attrib.get("dt", ""),
                    }
                    if author_id:
                        comment["author_id"] = author_id
                    if author.get("author"):
                        comment["author"] = author["author"]
                    if author.get("initials"):
                        comment["initials"] = author["initials"]
                    comments.append(comment)
            elif local_name_fallback(node.tag) == "text" and node.text:
                comments.append({"slide_index": slide_index, "part": part, "text": node.text.strip(), "date": ""})
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


def rel_target(slide_path: str, rel: dict[str, str]) -> str:
    target = rel.get("target", "")
    return resolve_part_path(slide_path, target) if target else ""


def relationship_ref(slide_path: str, rels: dict[str, dict[str, str]], rel_id: str) -> dict[str, str]:
    rel = rels.get(rel_id, {})
    return {
        "relationship_id": rel_id,
        "relationship_type": rel.get("type", ""),
        "target": rel_target(slide_path, rel),
        "target_mode": rel.get("target_mode", ""),
    }


def non_visual_props(node) -> dict[str, str]:
    props = node.find(".//p:cNvPr", NS)
    if props is None:
        return {"name": "", "alt_text": "", "title": ""}
    return {
        "name": props.attrib.get("name", ""),
        "alt_text": props.attrib.get("descr", ""),
        "title": props.attrib.get("title", ""),
    }


def transform_geometry(node) -> dict[str, int]:
    xfrm = node.find(".//a:xfrm", NS) or node.find(".//p:xfrm", NS)
    if xfrm is None:
        return {}
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    geometry: dict[str, int] = {}
    if off is not None:
        for key in ("x", "y"):
            if key in off.attrib:
                geometry[key] = int(off.attrib[key])
    if ext is not None:
        for key in ("cx", "cy"):
            if key in ext.attrib:
                geometry[key] = int(ext.attrib[key])
    return geometry


def visual_object_base(node, object_type: str) -> dict[str, Any]:
    base = {"object_type": object_type, **non_visual_props(node)}
    geometry = transform_geometry(node)
    if geometry:
        base["geometry"] = geometry
    return base


def append_visual_object(objects: list[dict[str, Any]], item: dict[str, Any]) -> None:
    key = (
        item.get("slide_index"),
        item.get("object_type"),
        item.get("relationship_id"),
        item.get("target"),
        item.get("name"),
        item.get("prog_id"),
    )
    for existing in objects:
        existing_key = (
            existing.get("slide_index"),
            existing.get("object_type"),
            existing.get("relationship_id"),
            existing.get("target"),
            existing.get("name"),
            existing.get("prog_id"),
        )
        if existing_key == key:
            return
    objects.append(item)


def relationship_id_from(node) -> str:
    for attr in (R_EMBED, R_LINK, R_ID):
        value = node.attrib.get(attr, "")
        if value:
            return value
    return ""


def diagram_relationships(node, slide_path: str, rels: dict[str, dict[str, str]]) -> list[dict[str, str]]:
    refs: list[dict[str, str]] = []
    for rel_ids in node.iter(qn("dgm", "relIds")):
        for attr, role in DIAGRAM_REL_ATTRS.items():
            rel_id = rel_ids.attrib.get(attr, "")
            if rel_id:
                refs.append({"role": role, **relationship_ref(slide_path, rels, rel_id)})
    return refs


def extract_media_file_objects(node, slide_path: str, slide_index: int, rels: dict[str, dict[str, str]]) -> list[dict[str, Any]]:
    media_objects: list[dict[str, Any]] = []
    for child in node.iter():
        tag_name = local_name_fallback(child.tag)
        if tag_name not in {"videoFile", "audioFile", "wavAudioFile"}:
            continue
        object_type = "video" if tag_name == "videoFile" else "audio"
        rel_id = relationship_id_from(child)
        item = visual_object_base(node, object_type)
        item.update({"slide_index": slide_index, **relationship_ref(slide_path, rels, rel_id)})
        media_objects.append(item)
    return media_objects


def extract_visual_objects(slide_root, slide_path: str, slide_index: int, rels: dict[str, dict[str, str]]) -> list[dict[str, Any]]:
    objects: list[dict[str, Any]] = []
    for pic in slide_root.iter(qn("p", "pic")):
        item = visual_object_base(pic, "image")
        blip = next(pic.iter(qn("a", "blip")), None)
        if blip is not None:
            rel_id = blip.attrib.get(R_EMBED, "")
            item.update(relationship_ref(slide_path, rels, rel_id))
        item["slide_index"] = slide_index
        append_visual_object(objects, item)
        for media_item in extract_media_file_objects(pic, slide_path, slide_index, rels):
            append_visual_object(objects, media_item)

    for frame in slide_root.iter(qn("p", "graphicFrame")):
        for chart in frame.iter(qn("c", "chart")):
            rel_id = chart.attrib.get(R_ID, "")
            item = visual_object_base(frame, "chart")
            item.update({"slide_index": slide_index, **relationship_ref(slide_path, rels, rel_id)})
            append_visual_object(objects, item)

        has_diagram_uri = any(
            "diagram" in graphic_data.attrib.get("uri", "").lower()
            for graphic_data in frame.iter(qn("a", "graphicData"))
        )
        diagram_rels = diagram_relationships(frame, slide_path, rels)
        if has_diagram_uri or diagram_rels:
            item = visual_object_base(frame, "smartart")
            item.update({"slide_index": slide_index, "relationships": diagram_rels})
            append_visual_object(objects, item)

        for ole in frame.iter(qn("p", "oleObj")):
            rel_id = relationship_id_from(ole)
            item = visual_object_base(frame, "ole")
            item.update(
                {
                    "slide_index": slide_index,
                    "prog_id": ole.attrib.get("progId", ""),
                    **relationship_ref(slide_path, rels, rel_id),
                }
            )
            append_visual_object(objects, item)

        for media_item in extract_media_file_objects(frame, slide_path, slide_index, rels):
            append_visual_object(objects, media_item)

    for shape in slide_root.iter(qn("p", "sp")):
        for media_item in extract_media_file_objects(shape, slide_path, slide_index, rels):
            append_visual_object(objects, media_item)

    return objects


def extract_pptx(source: Path, out_dir: Path) -> tuple[dict[str, Any], str]:
    manifest = default_manifest(source, "pptx")
    markdown: list[str] = []

    with zipfile.ZipFile(source) as package:
        manifest["metadata"] = core_properties(package)
        slide_paths = presentation_slide_paths(package)
        media = media_members(package, "ppt/media/")
        authors = comment_authors(package)

        for index, slide_path in enumerate(slide_paths, start=1):
            root = read_xml(package, slide_path)
            if root is None:
                continue
            texts = []
            for node in root.iter(qn("a", "t")):
                if node.text and node.text.strip():
                    texts.append(node.text.strip())
            title = slide_title(texts, index)
            body_text = "\n".join(texts)
            rels = read_relationships(package, slide_relationships_member(slide_path))
            notes = extract_notes(package, slide_path, index, rels)
            comments = extract_comments(package, slide_path, index, rels, authors)
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
            visual_objects = extract_visual_objects(root, slide_path, index, rels)
            for blip in root.iter(qn("a", "blip")):
                rel_id = blip.attrib.get(R_EMBED)
                rel = rels.get(rel_id or "")
                if rel:
                    slide_media.append({"relationship_id": rel_id or "", "target": rel.get("target", "")})
            if visual_objects or slide_media:
                manifest["visual_findings"].append(
                    {
                        "slide_index": index,
                        "requires_visual_review": True,
                        "reason": "slide contains visual objects or embedded media references",
                        "media": slide_media,
                        "objects": visual_objects,
                    }
                )

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
    out_dir.mkdir(parents=True, exist_ok=True)

    try:
        manifest, full_md = extract_pptx(source, out_dir)
    except Exception as exc:
        print(f"Failed to read PPTX: {exc}", file=sys.stderr)
        return 1

    full_path = out_dir / f"{source.stem}.full.md"
    manifest_path = out_dir / f"{source.stem}.manifest.json"
    full_path.write_text(full_md, encoding="utf-8")
    write_json(manifest_path, manifest)
    print(str(manifest_path))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
