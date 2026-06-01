#!/usr/bin/env python3
"""Shared OOXML helpers for the office-reader skill."""

from __future__ import annotations

import html
import json
import posixpath
import zipfile
from pathlib import Path
from typing import Any, Iterable
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


W_REL_ID = f"{{{NS['r']}}}id"
R_EMBED = f"{{{NS['r']}}}embed"
W_VAL = f"{{{NS['w']}}}val"
W_ID = f"{{{NS['w']}}}id"
W_AUTHOR = f"{{{NS['w']}}}author"
W_DATE = f"{{{NS['w']}}}date"


def qn(prefix: str, name: str) -> str:
    return f"{{{NS[prefix]}}}{name}"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def read_xml(package: zipfile.ZipFile, member: str) -> ET.Element | None:
    try:
        with package.open(member) as stream:
            return ET.fromstring(stream.read())
    except KeyError:
        return None


def read_relationships(package: zipfile.ZipFile, member: str) -> dict[str, dict[str, str]]:
    root = read_xml(package, member)
    if root is None:
        return {}
    relationships: dict[str, dict[str, str]] = {}
    for rel in root.findall("rel:Relationship", NS):
        rel_id = rel.attrib.get("Id")
        if not rel_id:
            continue
        relationships[rel_id] = {
            "id": rel_id,
            "type": rel.attrib.get("Type", ""),
            "target": rel.attrib.get("Target", ""),
            "target_mode": rel.attrib.get("TargetMode", ""),
        }
    return relationships


def resolve_part_path(base_part: str, target: str) -> str:
    base_dir = posixpath.dirname(base_part)
    return posixpath.normpath(posixpath.join(base_dir, target)).lstrip("/")


def collect_text(element: ET.Element | None, include_deleted: bool = True) -> str:
    if element is None:
        return ""
    parts: list[str] = []
    for node in element.iter():
        name = local_name(node.tag)
        if name == "t" and node.text:
            parts.append(node.text)
        elif include_deleted and name == "delText" and node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def collect_plain_text(element: ET.Element | None) -> str:
    if element is None:
        return ""
    parts: list[str] = []
    for node in element.iter():
        if local_name(node.tag) in {"t", "delText"} and node.text:
            parts.append(node.text)
    return "".join(parts).strip()


def paragraph_text_with_revisions(paragraph: ET.Element) -> tuple[str, list[dict[str, Any]]]:
    parts: list[str] = []
    revisions: list[dict[str, Any]] = []

    def walk(node: ET.Element, mode: str | None = None, meta: dict[str, str] | None = None) -> None:
        name = local_name(node.tag)
        if name == "ins":
            next_meta = {
                "author": node.attrib.get(W_AUTHOR, ""),
                "date": node.attrib.get(W_DATE, ""),
            }
            chunk = collect_text(node, include_deleted=False)
            if chunk:
                revisions.append({"type": "insertion", "text": chunk, **next_meta})
                parts.append("{+" + chunk + "+}")
            return
        if name == "del":
            next_meta = {
                "author": node.attrib.get(W_AUTHOR, ""),
                "date": node.attrib.get(W_DATE, ""),
            }
            chunk = collect_text(node, include_deleted=True)
            if chunk:
                revisions.append({"type": "deletion", "text": chunk, **next_meta})
                parts.append("{-" + chunk + "-}")
            return
        if name == "t" and node.text:
            parts.append(node.text)
        elif name == "delText" and node.text and mode == "delete":
            parts.append(node.text)
        for child in list(node):
            walk(child, mode, meta)

    walk(paragraph)
    return "".join(parts).strip(), revisions


def table_to_rows(table: ET.Element, cell_xpath: str, row_xpath: str) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.findall(row_xpath, NS):
        cells = [collect_plain_text(cell) for cell in row.findall(cell_xpath, NS)]
        if any(cells):
            rows.append(cells)
    return rows


def markdown_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    padded = [row + [""] * (width - len(row)) for row in rows]
    header = padded[0]
    separator = ["---"] * width
    body = padded[1:]

    def line(row: Iterable[str]) -> str:
        return "| " + " | ".join(str(cell).replace("\n", " ").strip() for cell in row) + " |"

    return "\n".join([line(header), line(separator), *[line(row) for row in body]])


def core_properties(package: zipfile.ZipFile) -> dict[str, str]:
    root = read_xml(package, "docProps/core.xml")
    if root is None:
        return {}
    fields = {
        "title": "dc:title",
        "creator": "dc:creator",
        "subject": "dc:subject",
        "description": "dc:description",
        "created": "dcterms:created",
        "modified": "dcterms:modified",
    }
    metadata: dict[str, str] = {}
    for key, xpath in fields.items():
        if ":" in xpath and xpath.split(":", 1)[0] not in NS:
            continue
        node = root.find(xpath, NS)
        if node is not None and node.text:
            metadata[key] = node.text.strip()
    return metadata


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def md_heading(text: str, level: int = 1) -> str:
    safe = text.strip() or "Untitled"
    level = max(1, min(level, 6))
    return f"{'#' * level} {safe}"


def escape_md(text: str) -> str:
    return html.unescape(text).strip()


def media_members(package: zipfile.ZipFile, prefix: str) -> list[str]:
    return sorted(name for name in package.namelist() if name.startswith(prefix))


def default_manifest(source: Path, document_type: str, normalized_file: Path | None = None) -> dict[str, Any]:
    normalized_file = normalized_file or source
    return {
        "source": {"path": str(source), "name": source.name},
        "normalized_file": {"path": str(normalized_file), "extension": normalized_file.suffix.lower()},
        "conversion": {"required": False, "backend": None, "status": "not_required", "messages": []},
        "document_type": document_type,
        "metadata": {},
        "structure": [],
        "tables": [],
        "comments": [],
        "revisions": [],
        "notes": [],
        "visual_findings": [],
        "artifacts": {},
    }
