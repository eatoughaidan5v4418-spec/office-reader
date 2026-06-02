import base64
import json
import importlib.util
import subprocess
import sys
import tempfile
import time
import types
import unittest
import zipfile
import os
import re
from pathlib import Path
from unittest.mock import patch


SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = SKILL_DIR / "scripts"
sys.path.insert(0, str(SCRIPTS_DIR))
VISUAL_ANALYSIS_SPEC = importlib.util.spec_from_file_location("visual_analysis", SCRIPTS_DIR / "visual_analysis.py")
visual_analysis = importlib.util.module_from_spec(VISUAL_ANALYSIS_SPEC)
assert VISUAL_ANALYSIS_SPEC.loader
VISUAL_ANALYSIS_SPEC.loader.exec_module(visual_analysis)
COMMON_OOXML_SPEC = importlib.util.spec_from_file_location("common_ooxml", SCRIPTS_DIR / "common_ooxml.py")
common_ooxml = importlib.util.module_from_spec(COMMON_OOXML_SPEC)
assert COMMON_OOXML_SPEC.loader
COMMON_OOXML_SPEC.loader.exec_module(common_ooxml)
SMOKE_READER_SPEC = importlib.util.spec_from_file_location("smoke_office_reader", SCRIPTS_DIR / "smoke_office_reader.py")
smoke_office_reader = importlib.util.module_from_spec(SMOKE_READER_SPEC)
assert SMOKE_READER_SPEC.loader
SMOKE_READER_SPEC.loader.exec_module(smoke_office_reader)
READ_OFFICE_SPEC = importlib.util.spec_from_file_location("read_office", SCRIPTS_DIR / "read_office.py")
read_office = importlib.util.module_from_spec(READ_OFFICE_SPEC)
assert READ_OFFICE_SPEC.loader
READ_OFFICE_SPEC.loader.exec_module(read_office)


def write_zip(path, files):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as package:
        for name, content in files.items():
            package.writestr(name, content)


def make_docx(path):
    files = {
        "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Executive Summary</w:t></w:r></w:p>
    <w:p>
      <w:r><w:t>Revenue grew </w:t></w:r>
      <w:ins w:author="Analyst" w:date="2026-01-02T00:00:00Z"><w:r><w:t>quickly</w:t></w:r></w:ins>
      <w:r><w:t> while costs </w:t></w:r>
      <w:del w:author="Analyst" w:date="2026-01-03T00:00:00Z"><w:r><w:delText>fell</w:delText></w:r></w:del>
      <w:r><w:t> stabilized.</w:t></w:r>
      <w:r><w:commentReference w:id="0"/></w:r>
    </w:p>
    <w:tbl>
      <w:tr><w:tc><w:p><w:r><w:t>Metric</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
      <w:tr><w:tc><w:p><w:r><w:t>ARR</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$10M</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
    <w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImage1"/></w:drawing></w:r></w:p>
  </w:body>
</w:document>""",
        "word/comments.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer" w:date="2026-01-04T00:00:00Z">
    <w:p><w:r><w:t>Please verify the ARR source.</w:t></w:r></w:p>
  </w:comment>
</w:comments>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>""",
        "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Board Memo</dc:title>
  <dc:creator>Finance Team</dc:creator>
</cp:coreProperties>""",
    }
    write_zip(path, files)


def replace_zip_member(path, member, content):
    replacement = path.with_suffix(path.suffix + ".replacement")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(replacement, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            if item.filename != member:
                target.writestr(item, source.read(item.filename))
        target.writestr(member, content)
    os.replace(replacement, path)


def remove_zip_member(path, member):
    replacement = path.with_suffix(path.suffix + ".replacement")
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(replacement, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            if item.filename != member:
                target.writestr(item, source.read(item.filename))
    os.replace(replacement, path)


def make_pptx(path):
    files = {
        "ppt/presentation.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>""",
        "ppt/_rels/presentation.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>""",
        "ppt/slides/slide1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:sp><p:txBody><a:p><a:r><a:t>Launch Readiness</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:txBody><a:p><a:r><a:t>Three </a:t></a:r><a:r><a:t>blockers remain.</a:t></a:r></a:p></p:txBody></p:sp>
    <p:graphicFrame><a:graphic><a:graphicData><a:tbl>
      <a:tr><a:tc><a:txBody><a:p><a:r><a:t>Owner</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>Status</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
      <a:tr><a:tc><a:txBody><a:p><a:r><a:t>Ops</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>At risk</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
    </a:tbl></a:graphicData></a:graphic></p:graphicFrame>
    <p:pic><p:blipFill><a:blip r:embed="rIdImage1"/></p:blipFill></p:pic>
  </p:spTree></p:cSld>
</p:sld>""",
        "ppt/slides/_rels/slide1.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNotes1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rIdComments1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
</Relationships>""",
        "ppt/notesSlides/notesSlide1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="1" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>Ask support to confirm staffing.</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Number Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum"/></p:nvPr></p:nvSpPr><p:txBody><a:p><a:r><a:t>1</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>""",
        "ppt/comments/comment1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cm authorId="0" dt="2026-02-01T00:00:00Z"><p:text>Clarify launch date.</p:text></p:cm>
</p:cmLst>""",
        "ppt/commentAuthors.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cmAuthor id="0" name="Reviewer One" initials="RO" lastIdx="1" clrIdx="0"/>
</p:cmAuthorLst>""",
        "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Launch Deck</dc:title>
</cp:coreProperties>""",
    }
    write_zip(path, files)


def make_pptx_with_orphan_comment(path):
    files = {
        "ppt/presentation.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>""",
        "ppt/_rels/presentation.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>""",
        "ppt/slides/slide1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>""",
        "ppt/slides/slide2.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>""",
        "ppt/comments/comment1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2026-02-01T00:00:00Z"><p:text>Orphan package comment.</p:text></p:cm>
</p:cmLst>""",
    }
    write_zip(path, files)


def make_pptx_with_late_relationship_comment(path):
    files = {
        "ppt/presentation.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>""",
        "ppt/_rels/presentation.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>""",
        "ppt/slides/slide1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>First slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>""",
        "ppt/slides/slide2.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Second slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:sld>""",
        "ppt/slides/_rels/slide2.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>
</Relationships>""",
        "ppt/comments/comment1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cm authorId="0" dt="2026-02-01T00:00:00Z"><p:text>Second slide comment.</p:text></p:cm>
</p:cmLst>""",
    }
    write_zip(path, files)


class OfficeReaderTests(unittest.TestCase):
    def run_script(self, script_name, *args):
        script = SCRIPTS_DIR / script_name
        proc = subprocess.run(
            [sys.executable, str(script), *map(str, args)],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        return proc

    def test_markdown_table_escapes_pipe_and_backslash_cells(self):
        rendered = common_ooxml.markdown_table([["Region", "Path"], ["North | South", r"C:\reports"]])

        self.assertIn(r"| North \| South | C:\\reports |", rendered)

    def test_atomic_write_text_replaces_destination_without_temp_residue(self):
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "artifact.md"
            target.write_text("old", encoding="utf-8")

            common_ooxml.atomic_write_text(target, "new")

            self.assertEqual(target.read_text(encoding="utf-8"), "new")
            self.assertFalse(list(target.parent.glob(f".{target.name}.*.tmp")))

    def test_word_text_collectors_preserve_tabs_and_line_breaks(self):
        paragraph = common_ooxml.ET.fromstring(
            f"""<w:p xmlns:w="{common_ooxml.NS['w']}">
  <w:r><w:t>Alpha</w:t><w:tab/><w:t>Beta</w:t><w:br/><w:t>Gamma</w:t><w:cr/><w:t>Delta</w:t></w:r>
</w:p>"""
        )

        plain = common_ooxml.collect_plain_text(paragraph)
        revised, revisions = common_ooxml.paragraph_text_with_revisions(paragraph)

        self.assertEqual(plain, "Alpha\tBeta\nGamma\nDelta")
        self.assertEqual(revised, "Alpha\tBeta\nGamma\nDelta")
        self.assertEqual(revisions, [])

    def test_resolve_part_path_rejects_package_escape_and_external_targets(self):
        self.assertEqual(common_ooxml.resolve_part_path("ppt/slides/slide1.xml", "../comments/comment1.xml"), "ppt/comments/comment1.xml")
        self.assertEqual(common_ooxml.resolve_part_path("ppt/slides/slide1.xml", "/ppt/comments/comment1.xml"), "ppt/comments/comment1.xml")
        self.assertEqual(common_ooxml.resolve_part_path("ppt/slides/slide1.xml", "../../../outside.xml"), "")
        self.assertEqual(common_ooxml.resolve_part_path("ppt/slides/slide1.xml", "https://example.com/comment.xml"), "")
        self.assertEqual(common_ooxml.resolve_part_path("ppt/slides/slide1.xml", "//server/share/comment.xml"), "")

    def test_docx_reader_outputs_markdown_and_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            make_docx(source)

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            full_md = (out_dir / "board-memo.full.md").read_text(encoding="utf-8")
            manifest = json.loads((out_dir / "board-memo.manifest.json").read_text(encoding="utf-8"))
            self.assertIn("# Executive Summary", full_md)
            self.assertIn("{+quickly+}", full_md)
            self.assertIn("{-fell-}", full_md)
            self.assertIn("| Metric | Value |", full_md)
            self.assertEqual(manifest["document_type"], "docx")
            self.assertEqual(manifest["comments"][0]["text"], "Please verify the ARR source.")
            self.assertEqual(manifest["revisions"][0]["type"], "insertion")
            self.assertEqual(manifest["tables"][0]["rows"][1], ["ARR", "$10M"])
            self.assertTrue(manifest["visual_findings"][0]["requires_visual_review"])

    def test_docx_reader_preserves_body_when_optional_comments_xml_is_malformed(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "malformed-comments.docx"
            out_dir = tmp_path / "out"
            make_docx(source)
            replace_zip_member(source, "word/comments.xml", "<w:comments")

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            full_md = (out_dir / "malformed-comments.full.md").read_text(encoding="utf-8")
            manifest = json.loads((out_dir / "malformed-comments.manifest.json").read_text(encoding="utf-8"))
            self.assertIn("# Executive Summary", full_md)
            self.assertEqual(manifest["comments"], [])
            self.assertIn("word/comments.xml", manifest["warnings"][0])

    def test_docx_reader_preserves_unknown_same_name_artifact(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            existing = out_dir / "board-memo.full.md"
            existing.write_text("user-owned markdown", encoding="utf-8")
            make_docx(source)

            proc = self.run_script("read_docx.py", source, "--out-dir", out_dir)

            manifest_path = Path(proc.stdout.strip())
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            generated_markdown = Path(manifest["artifacts"]["full_markdown"])
            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned markdown")
            self.assertNotEqual(generated_markdown, existing)
            self.assertTrue(generated_markdown.exists())
            self.assertNotEqual(manifest_path.parent, out_dir)

    def test_docx_reader_does_not_trust_incomplete_same_source_manifest_for_overwrite(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            existing_markdown = out_dir / "board-memo.full.md"
            existing_manifest = out_dir / "board-memo.manifest.json"
            existing_markdown.write_text("user-owned markdown", encoding="utf-8")
            make_docx(source)
            existing_manifest.write_text(
                json.dumps({"source": {"path": str(source)}, "document_type": "docx", "artifacts": {}}),
                encoding="utf-8",
            )

            proc = self.run_script("read_docx.py", source, "--out-dir", out_dir)

            generated_manifest = Path(proc.stdout.strip())
            self.assertEqual(existing_markdown.read_text(encoding="utf-8"), "user-owned markdown")
            self.assertNotEqual(generated_manifest, existing_manifest)
            self.assertNotEqual(generated_manifest.parent, out_dir)

    def test_docx_reader_ignores_external_media_relationship_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "external-media.docx"
            out_dir = tmp_path / "out"
            make_docx(source)
            replace_zip_member(
                source,
                "word/_rels/document.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png" TargetMode="External"/>
</Relationships>""",
            )

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "external-media.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["visual_findings"][0]["relationships"], [])
            self.assertTrue(any("rejected external ooxml relationship target" in warning.lower() for warning in manifest["warnings"]))

    def test_pptx_reader_extracts_slides_notes_comments_and_tables(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "launch-deck.pptx"
            out_dir = tmp_path / "out"
            make_pptx(source)

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            full_md = (out_dir / "launch-deck.full.md").read_text(encoding="utf-8")
            manifest = json.loads((out_dir / "launch-deck.manifest.json").read_text(encoding="utf-8"))
            self.assertIn("## Slide 1: Launch Readiness", full_md)
            self.assertIn("Three blockers remain.", full_md)
            self.assertIn("Ask support to confirm staffing.", full_md)
            self.assertEqual(manifest["document_type"], "pptx")
            self.assertEqual(manifest["structure"][0]["title"], "Launch Readiness")
            self.assertNotIn("Owner", manifest["structure"][0]["text"])
            self.assertEqual(manifest["notes"][0]["text"], "Ask support to confirm staffing.")
            self.assertEqual(manifest["comments"][0]["text"], "Clarify launch date.")
            self.assertEqual(manifest["comments"][0]["author"], "Reviewer One")
            self.assertEqual(manifest["tables"][0]["rows"][1], ["Ops", "At risk"])

    def test_pptx_reader_falls_back_to_sorted_slides_when_presentation_relationships_are_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "missing-presentation-rels.pptx"
            out_dir = tmp_path / "out"
            make_pptx(source)
            remove_zip_member(source, "ppt/_rels/presentation.xml.rels")

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "missing-presentation-rels.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["structure"]), 1)
            self.assertEqual(manifest["structure"][0]["title"], "Launch Readiness")
            self.assertTrue(any("slide relationships" in warning.lower() for warning in manifest["warnings"]))

    def test_pptx_reader_ignores_unsafe_slide_relationship_and_uses_fallback(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "unsafe-slide-relationship.pptx"
            out_dir = tmp_path / "out"
            make_pptx(source)
            replace_zip_member(
                source,
                "ppt/_rels/presentation.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../../../outside.xml"/>
</Relationships>""",
            )

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "unsafe-slide-relationship.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["structure"]), 1)
            self.assertEqual(manifest["structure"][0]["title"], "Launch Readiness")
            self.assertTrue(any("rejected unsafe relationship target" in warning.lower() for warning in manifest["warnings"]))

    def test_pptx_reader_uses_fallback_when_relationship_slide_member_is_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "missing-related-slide.pptx"
            out_dir = tmp_path / "out"
            make_pptx(source)
            replace_zip_member(
                source,
                "ppt/_rels/presentation.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/missing.xml"/>
</Relationships>""",
            )

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "missing-related-slide.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["structure"]), 1)
            self.assertEqual(manifest["structure"][0]["title"], "Launch Readiness")
            self.assertTrue(any("referenced slide part" in warning.lower() for warning in manifest["warnings"]))

    def test_pptx_reader_rejects_package_without_presentation_or_fallback_slides(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "empty-package.pptx"
            out_dir = tmp_path / "out"
            write_zip(source, {})

            proc = subprocess.run(
                [sys.executable, str(SCRIPTS_DIR / "read_pptx.py"), str(source), "--out-dir", str(out_dir)],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )

            self.assertEqual(proc.returncode, 1)
            self.assertIn("ppt/presentation.xml", proc.stderr)
            self.assertIn("no fallback slide XML parts", proc.stderr)

    def test_pptx_reader_rejects_package_when_all_slide_xml_is_malformed(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "malformed-slides.pptx"
            out_dir = tmp_path / "out"
            make_pptx(source)
            replace_zip_member(source, "ppt/slides/slide1.xml", "<p:sld>")

            proc = subprocess.run(
                [sys.executable, str(SCRIPTS_DIR / "read_pptx.py"), str(source), "--out-dir", str(out_dir)],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )

            self.assertEqual(proc.returncode, 1)
            self.assertIn("no readable slide XML parts", proc.stderr)

    def test_pptx_reader_does_not_duplicate_orphan_package_comments_across_slides(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "orphan-comment.pptx"
            out_dir = tmp_path / "out"
            make_pptx_with_orphan_comment(source)

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "orphan-comment.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["structure"]), 2)
            self.assertEqual(len(manifest["comments"]), 1)
            self.assertIsNone(manifest["comments"][0]["slide_index"])
            self.assertEqual(manifest["comments"][0]["text"], "Orphan package comment.")

    def test_pptx_reader_does_not_treat_later_relationship_comments_as_orphans(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "late-relationship-comment.pptx"
            out_dir = tmp_path / "out"
            make_pptx_with_late_relationship_comment(source)

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "late-relationship-comment.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["comments"]), 1)
            self.assertEqual(manifest["comments"][0]["slide_index"], 2)
            self.assertEqual(manifest["comments"][0]["text"], "Second slide comment.")

    def test_pptx_reader_ignores_external_comment_relationship_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "external-comment-relationship.pptx"
            out_dir = tmp_path / "out"
            make_pptx(source)
            replace_zip_member(
                source,
                "ppt/slides/_rels/slide1.xml.rels",
                """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdComments1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml" TargetMode="External"/>
</Relationships>""",
            )

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "external-comment-relationship.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(len(manifest["comments"]), 1)
            self.assertIsNone(manifest["comments"][0]["slide_index"])
            self.assertTrue(any("rejected external ooxml relationship target" in warning.lower() for warning in manifest["warnings"]))

    def test_report_assembler_uses_manifest_counts_and_outline(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            manifest_path = tmp_path / "sample.manifest.json"
            report_path = tmp_path / "sample.report.md"
            manifest_path.write_text(
                json.dumps(
                    {
                        "source": {"path": "sample.pptx", "name": "sample.pptx"},
                        "normalized_file": {"path": "sample.pptx", "extension": ".pptx"},
                        "conversion": {"required": False, "backend": None, "status": "not_required"},
                        "document_type": "pptx",
                        "structure": [{"type": "slide", "index": 1, "title": "Launch Readiness", "text": "Three blockers remain."}],
                        "tables": [{"index": 1, "rows": [["Owner", "Status"]]}],
                        "comments": [{"text": "Clarify launch date."}],
                        "revisions": [],
                        "notes": [{"slide_index": 1, "text": "Ask support to confirm staffing."}],
                        "warnings": ["Skipped malformed optional OOXML part ppt/notesSlides/notesSlide1.xml: invalid token"],
                        "visual_findings": [{"requires_visual_review": True, "reason": "slide has media"}],
                        "artifacts": {"full_markdown": "sample.full.md"},
                    }
                ),
                encoding="utf-8",
            )

            self.run_script("assemble_report.py", manifest_path, "--out", report_path)

            report = report_path.read_text(encoding="utf-8")
            self.assertIn("# Structured Reading Report: sample.pptx", report)
            self.assertIn("- Slides/sections: 1", report)
            self.assertIn("- Comments: 1", report)
            self.assertIn("Launch Readiness", report)
            self.assertIn("Comment unattributed: Clarify launch date.", report)
            self.assertIn("- Extraction warnings: 1", report)
            self.assertIn("Skipped malformed optional OOXML part ppt/notesSlides/notesSlide1.xml", report)

    def test_backend_discovery_emits_json_shape(self):
        script = SCRIPTS_DIR / "discover_office_backends.ps1"
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script), "-InputExtension", ".doc", "-Format", "json"],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        data = json.loads(proc.stdout)
        self.assertEqual(data["input_extension"], ".doc")
        self.assertIn("backends", data)
        self.assertEqual([item["name"] for item in data["priority_order"]], ["office-com", "wps", "libreoffice"])

    def test_unified_reader_dispatches_and_builds_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            make_docx(source)

            self.run_script("read_office.py", source, "--out-dir", out_dir)

            self.assertTrue((out_dir / "board-memo.full.md").exists())
            self.assertTrue((out_dir / "board-memo.manifest.json").exists())
            report = (out_dir / "board-memo.report.md").read_text(encoding="utf-8")
            self.assertIn("# Structured Reading Report: board-memo.docx", report)
            self.assertIn("- Comments: 1", report)

    def test_unified_reader_reports_run_specific_markdown_path_after_artifact_collision(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            existing = out_dir / "board-memo.full.md"
            existing.write_text("user-owned markdown", encoding="utf-8")
            make_docx(source)

            proc = self.run_script("read_office.py", source, "--out-dir", out_dir, "--mode", "fast", "--no-openai-vision")

            result = json.loads(proc.stdout)
            manifest = json.loads(Path(result["manifest"]).read_text(encoding="utf-8"))
            generated_markdown = Path(manifest["artifacts"]["full_markdown"])
            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned markdown")
            self.assertEqual(Path(result["full_markdown"]), generated_markdown)
            self.assertNotEqual(generated_markdown, existing)
            self.assertTrue(generated_markdown.exists())

    def test_unified_reader_preserves_unknown_same_name_report_artifact(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            existing = out_dir / "board-memo.report.md"
            existing.write_text("user-owned report", encoding="utf-8")
            make_docx(source)

            proc = self.run_script("read_office.py", source, "--out-dir", out_dir, "--mode", "fast", "--no-openai-vision")

            result = json.loads(proc.stdout)
            generated_report = Path(result["report"])
            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned report")
            self.assertNotEqual(generated_report, existing)
            self.assertTrue(generated_report.exists())

    def test_unified_reader_rejects_reader_manifest_outside_output_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "board-memo.docx"
            normalized.write_bytes(b"document")
            out_dir = tmp_path / "out"
            outside_manifest = tmp_path / "outside.manifest.json"
            outside_manifest.write_text("{}", encoding="utf-8")
            completed = subprocess.CompletedProcess(["python"], 0, str(outside_manifest), "")

            with patch.object(read_office, "run_command", return_value=completed):
                with self.assertRaises(RuntimeError):
                    read_office.run_reader(normalized, out_dir)

    def test_unified_reader_rejects_unreadable_or_non_object_reader_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "board-memo.docx"
            normalized.write_bytes(b"document")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            manifest = out_dir / "board-memo.manifest.json"

            for content in ("not-json", "[]"):
                with self.subTest(content=content):
                    manifest.write_text(content, encoding="utf-8")
                    completed = subprocess.CompletedProcess(["python"], 0, str(manifest), "")
                    with patch.object(read_office, "run_command", return_value=completed):
                        with self.assertRaises(RuntimeError):
                            read_office.run_reader(normalized, out_dir)

    def test_unified_reader_rejects_reader_manifest_with_external_markdown_artifact(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "board-memo.docx"
            normalized.write_bytes(b"document")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            manifest = out_dir / "board-memo.manifest.json"
            outside_markdown = tmp_path / "outside.full.md"
            outside_markdown.write_text("outside", encoding="utf-8")
            manifest.write_text(
                json.dumps({"artifacts": {"full_markdown": str(outside_markdown), "manifest": str(manifest)}}),
                encoding="utf-8",
            )
            completed = subprocess.CompletedProcess(["python"], 0, str(manifest), "")

            with patch.object(read_office, "run_command", return_value=completed):
                with self.assertRaises(RuntimeError):
                    read_office.run_reader(normalized, out_dir)

    def test_unified_reader_rejects_visual_analysis_redirected_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "board-memo.docx"
            normalized.write_bytes(b"document")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            manifest = out_dir / "board-memo.manifest.json"
            manifest.write_text("{}", encoding="utf-8")
            redirected = out_dir / "redirected.manifest.json"
            redirected.write_text("{}", encoding="utf-8")
            completed = subprocess.CompletedProcess(["python"], 0, str(redirected), "")

            with patch.object(read_office, "run_command", return_value=completed):
                with self.assertRaises(RuntimeError):
                    read_office.run_visual_analysis(manifest, normalized, out_dir, "balanced", False, 90)

    def test_unified_reader_rejects_visual_analysis_corrupted_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "board-memo.docx"
            normalized.write_bytes(b"document")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            manifest = out_dir / "board-memo.manifest.json"
            markdown = out_dir / "board-memo.full.md"
            markdown.write_text("markdown", encoding="utf-8")
            manifest.write_text(
                json.dumps({"artifacts": {"full_markdown": str(markdown), "manifest": str(manifest)}}),
                encoding="utf-8",
            )

            def corrupt_manifest(command):
                manifest.write_text("[]", encoding="utf-8")
                return subprocess.CompletedProcess(command, 0, str(manifest), "")

            with patch.object(read_office, "run_command", side_effect=corrupt_manifest):
                with self.assertRaises(RuntimeError):
                    read_office.run_visual_analysis(manifest, normalized, out_dir, "balanced", False, 90)

    def test_unified_reader_rejects_report_assembler_redirected_markdown(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            manifest = tmp_path / "board-memo.manifest.json"
            manifest.write_text("{}", encoding="utf-8")
            redirected = tmp_path / "redirected.report.md"
            redirected.write_text("report", encoding="utf-8")
            completed = subprocess.CompletedProcess(["python"], 0, str(redirected), "")

            with patch.object(read_office, "run_command", return_value=completed):
                with self.assertRaises(RuntimeError):
                    read_office.assemble_report(manifest)

    def test_unified_reader_accepts_modes_and_enriches_visual_manifest(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            make_docx(source)

            self.run_script("read_office.py", source, "--out-dir", out_dir, "--mode", "balanced", "--no-openai-vision")

            manifest = json.loads((out_dir / "board-memo.manifest.json").read_text(encoding="utf-8"))
            self.assertEqual(manifest["reading_mode"], "balanced")
            self.assertIn("visual_analysis", manifest)
            self.assertEqual(manifest["visual_analysis"]["mode"], "balanced")
            self.assertIn(manifest["visual_analysis"]["status"], {"completed", "skipped", "partial"})
            self.assertTrue(manifest["visual_findings"])
            finding = manifest["visual_findings"][0]
            self.assertIn("ocr_text", finding)
            self.assertIn("vision_summary", finding)
            self.assertIn("diagram_summary", finding)
            self.assertIn("backend", finding)
            self.assertIn("cache_hit", finding)

    def test_visual_cache_marks_second_run_cache_hit(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            make_docx(source)

            self.run_script("read_office.py", source, "--out-dir", out_dir, "--mode", "balanced", "--no-openai-vision")
            self.run_script("read_office.py", source, "--out-dir", out_dir, "--mode", "balanced", "--no-openai-vision")

            manifest = json.loads((out_dir / "board-memo.manifest.json").read_text(encoding="utf-8"))
            self.assertTrue(any(item.get("cache_hit") for item in manifest["visual_findings"]))

    def test_visual_cache_separates_local_only_from_openai_enabled_runs(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            image = tmp_path / "page.png"
            image.write_bytes(b"fake image bytes")
            cache_dir = tmp_path / "cache"
            cache_dir.mkdir()
            item = {"image_path": str(image), "requires_visual_review": True}
            messages = []
            with (
                patch.object(visual_analysis, "local_ocr", return_value=("OCR text", "rapidocr")),
                patch.object(
                    visual_analysis,
                    "openai_visual_summary",
                    side_effect=lambda image_path, ocr_text, output_messages, enabled: (
                        ("Vision text", "openai:test") if enabled else ("", "")
                    ),
                ),
            ):
                local_only = visual_analysis.analyze_item("source", "complete", cache_dir, item, messages, False)
                cloud_first = visual_analysis.analyze_item("source", "complete", cache_dir, item, messages, True)
                cloud_second = visual_analysis.analyze_item("source", "complete", cache_dir, item, messages, True)
            self.assertFalse(local_only["cache_hit"])
            self.assertEqual(local_only["vision_summary"], "")
            self.assertFalse(cloud_first["cache_hit"])
            self.assertEqual(cloud_first["vision_summary"], "Vision text")
            self.assertTrue(cloud_second["cache_hit"])

    def test_visual_cache_recomputes_when_cached_json_is_malformed(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            image = tmp_path / "page.png"
            image.write_bytes(b"fake image bytes")
            cache_dir = tmp_path / "cache"
            cache_dir.mkdir()
            item = {"image_path": str(image), "requires_visual_review": True}
            cache = visual_analysis.cache_path(
                cache_dir,
                "source",
                "complete",
                "local-only",
                visual_analysis.sha256_file(image),
            )
            cache.write_text("{broken", encoding="utf-8")
            messages = []

            with (
                patch.object(visual_analysis, "local_ocr", return_value=("OCR text", "rapidocr")),
                patch.object(visual_analysis, "openai_visual_summary", return_value=("", "")),
            ):
                result = visual_analysis.analyze_item("source", "complete", cache_dir, item, messages, False)

            self.assertFalse(result["cache_hit"])
            self.assertEqual(result["ocr_text"], "OCR text")
            self.assertTrue(any("ignored unreadable visual cache" in message.lower() for message in messages))
            self.assertEqual(json.loads(cache.read_text(encoding="utf-8"))["ocr_text"], "OCR text")

    def test_visual_cache_hit_refreshes_current_rendered_page_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            first_image = tmp_path / "render-one" / "page-001.png"
            second_image = tmp_path / "render-two" / "page-001.png"
            first_image.parent.mkdir()
            second_image.parent.mkdir()
            first_image.write_bytes(b"same image bytes")
            second_image.write_bytes(b"same image bytes")
            cache_dir = tmp_path / "cache"
            cache_dir.mkdir()
            messages = []

            with (
                patch.object(visual_analysis, "local_ocr", return_value=("OCR text", "rapidocr")),
                patch.object(visual_analysis, "openai_visual_summary", return_value=("", "")),
            ):
                visual_analysis.analyze_item(
                    "source",
                    "complete",
                    cache_dir,
                    {"image_path": str(first_image), "page_index": 1, "origin": "rendered-page"},
                    messages,
                    False,
                )
                cached = visual_analysis.analyze_item(
                    "source",
                    "complete",
                    cache_dir,
                    {"image_path": str(second_image), "page_index": 2, "origin": "rendered-page"},
                    messages,
                    False,
                )

            self.assertTrue(cached["cache_hit"])
            self.assertEqual(cached["image_path"], str(second_image))
            self.assertEqual(cached["page_index"], 2)
            self.assertEqual(cached["origin"], "rendered-page")

    def test_visual_enrichment_replaces_prior_rendered_page_findings_on_rerun(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "sample.pptx"
            normalized.write_bytes(b"normalized presentation")
            full_md = tmp_path / "sample.full.md"
            full_md.write_text("# Sample\n", encoding="utf-8")
            manifest_path = tmp_path / "sample.manifest.json"
            manifest_path.write_text(
                json.dumps(
                    {
                        "document_type": "pptx",
                        "visual_findings": [{"slide_index": 1, "requires_visual_review": True, "reason": "slide contains media"}],
                        "artifacts": {"full_markdown": str(full_md)},
                    }
                ),
                encoding="utf-8",
            )
            image = tmp_path / "page-001.png"
            image.write_bytes(b"image")

            def fake_analyze(source_hash, mode, cache_dir, item, messages, enable_openai):
                return {
                    **visual_analysis.VISUAL_FIELDS,
                    **item,
                    "ocr_text": "Visible text",
                    "backend": "mock-ocr",
                    "confidence": "medium",
                }

            with (
                patch.object(visual_analysis, "render_preview", return_value=(tmp_path / "preview.pdf", {"status": "success"})),
                patch.object(visual_analysis, "render_pdf_pages", return_value=[image]),
                patch.object(visual_analysis, "analyze_item", side_effect=fake_analyze),
            ):
                visual_analysis.enrich_manifest(manifest_path, normalized, tmp_path, "complete", False, 1)
                visual_analysis.enrich_manifest(manifest_path, normalized, tmp_path, "complete", False, 1)

            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            rendered = [item for item in manifest["visual_findings"] if item.get("image_path")]
            self.assertEqual(len(rendered), 1)

    def test_render_pdf_pages_preserves_existing_page_image(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pdf = tmp_path / "preview.pdf"
            pdf.write_bytes(b"pdf placeholder")
            existing_dir = tmp_path / "page_images"
            existing_dir.mkdir()
            existing = existing_dir / "page-001.png"
            existing.write_text("user-owned image", encoding="utf-8")

            class FakePixmap:
                def save(self, path):
                    Path(path).write_text("rendered image", encoding="utf-8")

            class FakePage:
                def get_pixmap(self, matrix, alpha):
                    return FakePixmap()

            class FakeDocument:
                def __enter__(self):
                    return self

                def __exit__(self, exc_type, exc, traceback):
                    return False

                def __len__(self):
                    return 1

                def load_page(self, index):
                    return FakePage()

            fake_fitz = types.SimpleNamespace(open=lambda path: FakeDocument(), Matrix=lambda x, y: (x, y))
            with patch.dict(sys.modules, {"fitz": fake_fitz}):
                rendered = visual_analysis.render_pdf_pages(pdf, tmp_path, [])

            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned image")
            self.assertEqual(len(rendered), 1)
            self.assertNotEqual(rendered[0], existing)
            self.assertTrue(rendered[0].exists())

    def test_visual_preview_rejects_pdf_outside_preview_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "sample.docx"
            normalized.write_bytes(b"document")
            outside_pdf = tmp_path / "outside.pdf"
            outside_pdf.write_bytes(b"pdf")
            messages = []
            completed = subprocess.CompletedProcess(
                ["powershell"],
                0,
                json.dumps({"status": "success", "artifacts": [str(outside_pdf)], "messages": []}),
                "",
            )

            with patch.object(visual_analysis, "run_command", return_value=completed):
                pdf, result = visual_analysis.render_preview(normalized, tmp_path / "out", 1, messages)

            self.assertIsNone(pdf)
            self.assertEqual(result["status"], "failed")
            self.assertTrue(any("outside preview directory" in message.lower() for message in messages))

    def test_visual_preview_reports_non_object_json_as_failed_result(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "sample.docx"
            normalized.write_bytes(b"document")
            messages = []
            completed = subprocess.CompletedProcess(["powershell"], 0, "[]", "")

            with patch.object(visual_analysis, "run_command", return_value=completed):
                pdf, result = visual_analysis.render_preview(normalized, tmp_path / "out", 1, messages)

            self.assertIsNone(pdf)
            self.assertEqual(result["status"], "failed")
            self.assertTrue(any("json object" in message.lower() for message in messages))

    def test_visual_preview_reports_malformed_json_fields_as_failed_result(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "sample.docx"
            normalized.write_bytes(b"document")
            messages = []
            completed = subprocess.CompletedProcess(
                ["powershell"],
                0,
                json.dumps({"status": "success", "messages": None, "artifacts": None}),
                "",
            )

            with patch.object(visual_analysis, "run_command", return_value=completed):
                pdf, result = visual_analysis.render_preview(normalized, tmp_path / "out", 1, messages)

            self.assertIsNone(pdf)
            self.assertEqual(result["status"], "failed")
            self.assertTrue(any("list fields" in message.lower() for message in messages))

    def test_visual_preview_reports_non_string_artifact_as_failed_result(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            normalized = tmp_path / "sample.docx"
            normalized.write_bytes(b"document")
            messages = []
            completed = subprocess.CompletedProcess(
                ["powershell"],
                0,
                json.dumps({"status": "success", "messages": [], "artifacts": [123]}),
                "",
            )

            with patch.object(visual_analysis, "run_command", return_value=completed):
                pdf, result = visual_analysis.render_preview(normalized, tmp_path / "out", 1, messages)

            self.assertIsNone(pdf)
            self.assertEqual(result["status"], "failed")
            self.assertTrue(any("artifact paths" in message.lower() for message in messages))

    def test_render_pdf_pages_removes_empty_run_directory_after_failure(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            pdf = tmp_path / "broken-preview.pdf"
            pdf.write_bytes(b"not a pdf")
            messages = []
            fake_fitz = types.SimpleNamespace(open=lambda path: (_ for _ in ()).throw(ValueError("broken pdf")))

            with patch.dict(sys.modules, {"fitz": fake_fitz}):
                rendered = visual_analysis.render_pdf_pages(pdf, tmp_path, messages)

            self.assertEqual(rendered, [])
            self.assertTrue(any("page rendering failed" in message.lower() for message in messages))
            self.assertEqual(list((tmp_path / "page_images").glob("render-*")), [])

    def test_local_ocr_distinguishes_blank_output_from_missing_backends(self):
        messages = []
        with (
            patch.object(visual_analysis, "rapidocr_text", return_value=("", "")),
            patch.object(visual_analysis, "tesseract_ocr", return_value=("", "")),
            patch.object(visual_analysis.shutil, "which", return_value="C:/tools/tesseract.exe"),
        ):
            text, backend = visual_analysis.local_ocr(Path("blank-page.png"), messages)
        self.assertEqual((text, backend), ("", ""))
        self.assertTrue(any("returned no text" in message for message in messages))
        self.assertFalse(any("backend is unavailable" in message for message in messages))

    def test_fast_mode_does_not_render_pages(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            make_docx(source)

            self.run_script("read_office.py", source, "--out-dir", out_dir, "--mode", "fast", "--no-openai-vision")

            manifest = json.loads((out_dir / "board-memo.manifest.json").read_text(encoding="utf-8"))
            visual = manifest["visual_analysis"]
            self.assertEqual(visual["mode"], "fast")
            self.assertEqual(visual["rendered_page_count"], 0)
            self.assertIn("fast mode", " ".join(visual["messages"]).lower())

    def test_visual_enrichment_does_not_write_markdown_outside_output_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            manifest_path = out_dir / "sample.manifest.json"
            normalized = tmp_path / "sample.docx"
            normalized.write_bytes(b"normalized")
            outside_markdown = tmp_path / "user-owned.md"
            outside_markdown.write_text("# User owned\n", encoding="utf-8")
            manifest_path.write_text(
                json.dumps(
                    {
                        "document_type": "docx",
                        "visual_findings": [{"requires_visual_review": True, "reason": "media"}],
                        "artifacts": {"full_markdown": str(outside_markdown)},
                    }
                ),
                encoding="utf-8",
            )

            visual_analysis.enrich_manifest(manifest_path, normalized, out_dir, "fast", False, 1)

            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(outside_markdown.read_text(encoding="utf-8"), "# User owned\n")
            self.assertTrue(any("outside output directory" in message.lower() for message in manifest["visual_analysis"]["messages"]))

    def test_report_includes_visual_deep_read_sections(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            manifest_path = tmp_path / "sample.manifest.json"
            report_path = tmp_path / "sample.report.md"
            manifest_path.write_text(
                json.dumps(
                    {
                        "source": {"path": "sample.docx", "name": "sample.docx"},
                        "normalized_file": {"path": "sample.docx", "extension": ".docx"},
                        "conversion": {"required": False, "backend": None, "status": "not_required"},
                        "document_type": "docx",
                        "reading_mode": "balanced",
                        "structure": [{"type": "paragraph", "index": 1, "text": "Body text"}],
                        "tables": [],
                        "comments": [],
                        "revisions": [],
                        "notes": [],
                        "visual_analysis": {
                            "status": "completed",
                            "mode": "balanced",
                            "rendered_page_count": 1,
                            "analyzed_item_count": 1,
                            "cache_hits": 0,
                            "messages": [],
                        },
                        "visual_findings": [
                            {
                                "requires_visual_review": True,
                                "reason": "page has embedded image",
                                "page_index": 1,
                                "ocr_text": "图中写着 HT32",
                                "vision_summary": "系统框图展示传感器到 MCU 的链路。",
                                "diagram_summary": "压力传感器 -> HT32 -> 蓝牙。",
                                "backend": "rapidocr+openai",
                                "confidence": "medium",
                                "cache_hit": False,
                            }
                        ],
                        "artifacts": {"full_markdown": "sample.full.md"},
                    }
                ),
                encoding="utf-8",
            )

            self.run_script("assemble_report.py", manifest_path, "--out", report_path)

            report = report_path.read_text(encoding="utf-8")
            self.assertIn("## Visual Deep Read", report)
            self.assertIn("图中写着 HT32", report)
            self.assertIn("系统框图展示传感器到 MCU 的链路。", report)
            self.assertIn("## Remaining Unverified Visual Gaps", report)

    def test_report_assembler_reports_damaged_manifest_without_traceback(self):
        with tempfile.TemporaryDirectory() as tmp:
            manifest_path = Path(tmp) / "damaged.manifest.json"
            manifest_path.write_text("[]", encoding="utf-8")

            proc = subprocess.run(
                [sys.executable, str(SCRIPTS_DIR / "assemble_report.py"), str(manifest_path)],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
            )

            self.assertEqual(proc.returncode, 1)
            self.assertIn("Failed to assemble report", proc.stderr)
            self.assertNotIn("Traceback", proc.stderr)

    def test_bootstrap_deps_dry_run_reports_install_plan(self):
        script = SCRIPTS_DIR / "bootstrap_deps.ps1"
        proc = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-DryRun",
                "-IncludeSystemTools",
            ],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
            timeout=30,
        )
        self.assertEqual(proc.returncode, 0, proc.stderr)
        data = json.loads(proc.stdout)
        self.assertEqual(data["status"], "dry_run")
        self.assertIn("python_packages", data)
        self.assertIn("system_tools", data)
        self.assertTrue(any(pkg["name"] == "rapidocr" for pkg in data["python_packages"]))
        self.assertTrue(any(tool["name"] == "LibreOffice" for tool in data["system_tools"]))

    def test_bootstrap_deps_refreshes_system_tool_status_after_install_attempts(self):
        script = (SCRIPTS_DIR / "bootstrap_deps.ps1").read_text(encoding="utf-8")

        self.assertIn("function Get-SystemToolPlan", script)
        self.assertGreaterEqual(script.count("$toolPlan = @(Get-SystemToolPlan)"), 2)
        self.assertIn("install_exit_code", script)
        self.assertIn("install_result", script)
        self.assertIn('Install failed for $($tool.name)', script)

    def test_bootstrap_deps_fails_when_required_python_packages_remain_missing(self):
        script = (SCRIPTS_DIR / "bootstrap_deps.ps1").read_text(encoding="utf-8")

        self.assertIn("$missingRequiredPackages", script)
        self.assertIn('if ($missingRequiredPackages.Count -gt 0) { "failed" } else { "completed" }', script)
        self.assertIn("Required Python packages are still missing after bootstrap", script)
        self.assertIn("exit 1", script)

    def test_unified_reader_bootstrap_forwards_system_tool_install_request(self):
        completed = subprocess.CompletedProcess(["powershell"], 0, json.dumps({"status": "completed"}), "")
        with patch.object(read_office, "run_command", return_value=completed) as run:
            read_office.bootstrap_deps(include_system_tools=True, install_system_tools=True)

        command = run.call_args.args[0]
        self.assertIn("-IncludeSystemTools", command)
        self.assertIn("-InstallSystemTools", command)

    def test_unified_reader_bootstrap_rejects_failed_or_malformed_success_output(self):
        cases = [
            subprocess.CompletedProcess(["powershell"], 0, json.dumps({"status": "failed", "messages": ["missing"]}), ""),
            subprocess.CompletedProcess(["powershell"], 0, "not-json", ""),
            subprocess.CompletedProcess(["powershell"], 0, "[]", ""),
        ]
        for completed in cases:
            with self.subTest(stdout=completed.stdout):
                with patch.object(read_office, "run_command", return_value=completed):
                    with self.assertRaises(RuntimeError):
                        read_office.bootstrap_deps()

    def test_unified_reader_legacy_conversion_rejects_unusable_success_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "legacy.doc"
            source.write_bytes(b"legacy")
            outside_normalized = tmp_path / "outside.docx"
            outside_normalized.write_bytes(b"outside")
            cases = [
                subprocess.CompletedProcess(["powershell"], 0, json.dumps({"status": "failed", "messages": ["failed"]}), ""),
                subprocess.CompletedProcess(["powershell"], 0, "not-json", ""),
                subprocess.CompletedProcess(
                    ["powershell"],
                    0,
                    json.dumps({"status": "success", "output_path": str(tmp_path / "missing.docx")}),
                    "",
                ),
                subprocess.CompletedProcess(
                    ["powershell"],
                    0,
                    json.dumps({"status": "success", "output_path": str(outside_normalized)}),
                    "",
                ),
                subprocess.CompletedProcess(["powershell"], 0, json.dumps({"status": "success", "output_path": 123}), ""),
            ]

            for completed in cases:
                with self.subTest(stdout=completed.stdout):
                    with patch.object(read_office, "run_command", return_value=completed):
                        with self.assertRaises(RuntimeError):
                            read_office.convert_legacy(source, tmp_path / "out", 1)

    def test_unified_reader_rejects_system_tool_install_without_dependency_install_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "board-memo.docx"
            make_docx(source)

            proc = subprocess.run(
                [sys.executable, str(SCRIPTS_DIR / "read_office.py"), str(source), "--install-system-tools"],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
            )

            self.assertEqual(proc.returncode, 2)
            self.assertIn("requires --install-missing-deps", proc.stderr)

    def test_libreoffice_environment_override_is_consistent_across_diagnostics(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            fake_soffice = tmp_path / "soffice.cmd"
            fake_soffice.write_text("@echo off\r\nexit /b 0\r\n", encoding="utf-8")
            env = os.environ.copy()
            env["SOFFICE_PATH"] = str(fake_soffice)

            discovery = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(SCRIPTS_DIR / "discover_office_backends.ps1"),
                    "-InputExtension",
                    ".doc",
                    "-Format",
                    "json",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
                env=env,
            )
            self.assertEqual(discovery.returncode, 0, discovery.stderr)
            discovery_data = json.loads(discovery.stdout)
            self.assertEqual(discovery_data["backends"]["libreoffice"]["executable"], str(fake_soffice))

            bootstrap = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(SCRIPTS_DIR / "bootstrap_deps.ps1"),
                    "-DryRun",
                    "-IncludeSystemTools",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
                env=env,
            )
            self.assertEqual(bootstrap.returncode, 0, bootstrap.stderr)
            bootstrap_data = json.loads(bootstrap.stdout)
            libreoffice = next(item for item in bootstrap_data["system_tools"] if item["name"] == "LibreOffice")
            self.assertTrue(libreoffice["available"])
            self.assertEqual(libreoffice["source"], str(fake_soffice))

    def test_legacy_conversion_times_out_com_worker_then_falls_back_to_libreoffice(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "legacy.doc"
            source.write_text("legacy placeholder", encoding="utf-8")
            fake_soffice = tmp_path / "soffice.cmd"
            fake_soffice.write_text(
                "@echo off\r\n"
                "set OUTDIR=\r\n"
                ":args\r\n"
                "if \"%~1\"==\"\" goto doneargs\r\n"
                "if \"%~1\"==\"--outdir\" (\r\n"
                "  set OUTDIR=%~2\r\n"
                "  shift\r\n"
                ")\r\n"
                "shift\r\n"
                "goto args\r\n"
                ":doneargs\r\n"
                "echo converted> \"%OUTDIR%\\legacy.docx\"\r\n"
                "exit /b 0\r\n",
                encoding="utf-8",
            )
            env = os.environ.copy()
            env["SOFFICE_PATH"] = str(fake_soffice)
            env["OFFICE_READER_COM_WORKER_DELAY_SECONDS"] = "6"
            worker_log = tmp_path / "worker-pids.log"
            env["OFFICE_READER_WORKER_PID_LOG"] = str(worker_log)

            started = time.perf_counter()
            proc = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(SCRIPTS_DIR / "convert_legacy_office.ps1"),
                    "-InputPath",
                    str(source),
                    "-OutputDir",
                    str(tmp_path / "out"),
                    "-TimeoutSeconds",
                    "3",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=20,
                env=env,
            )
            elapsed = time.perf_counter() - started
            self.assertEqual(proc.returncode, 0, proc.stderr)
            self.assertLess(elapsed, 15)
            data = json.loads(proc.stdout)
            self.assertEqual(data["status"], "success")
            self.assertEqual(data["backend"], "libreoffice")
            self.assertTrue(any("timed out" in message.lower() for message in data["messages"]))
            self.assertTrue((tmp_path / "out" / "legacy.docx").exists())
            self.assertFalse(list((tmp_path / "out").glob("legacy-worker-*")))
            self.assertFalse(list((tmp_path / "out").glob(".lo_profile_*")))
            worker_pids = {
                backend: int(pid)
                for backend, pid in (line.split(",", 1) for line in worker_log.read_text(encoding="utf-8").splitlines())
            }
            office_worker = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-Command",
                    f"if (Get-Process -Id {worker_pids['office-com']} -ErrorAction SilentlyContinue) {{ exit 1 }}",
                ],
                check=False,
            )
            self.assertEqual(office_worker.returncode, 0)

    def test_legacy_conversion_preserves_existing_normalized_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "legacy.doc"
            source.write_text("legacy placeholder", encoding="utf-8")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            existing = out_dir / "legacy.docx"
            existing.write_text("user-owned normalized document", encoding="utf-8")
            fake_soffice = tmp_path / "soffice.cmd"
            fake_soffice.write_text(
                "@echo off\r\n"
                "set OUTDIR=\r\n"
                ":args\r\n"
                "if \"%~1\"==\"\" goto doneargs\r\n"
                "if \"%~1\"==\"--outdir\" (\r\n"
                "  set OUTDIR=%~2\r\n"
                "  shift\r\n"
                ")\r\n"
                "shift\r\n"
                "goto args\r\n"
                ":doneargs\r\n"
                "echo converted> \"%OUTDIR%\\legacy.docx\"\r\n"
                "exit /b 0\r\n",
                encoding="utf-8",
            )
            env = os.environ.copy()
            env["SOFFICE_PATH"] = str(fake_soffice)

            proc = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(SCRIPTS_DIR / "convert_legacy_office.ps1"),
                    "-InputPath",
                    str(source),
                    "-OutputDir",
                    str(out_dir),
                    "-PreferredBackend",
                    "libreoffice",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
                env=env,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            data = json.loads(proc.stdout)
            converted = Path(data["output_path"])
            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned normalized document")
            self.assertNotEqual(converted, existing)
            self.assertTrue(converted.exists())

    def test_unified_reader_and_smoke_harness_expose_legacy_timeout(self):
        reader_help = subprocess.run(
            [sys.executable, str(SCRIPTS_DIR / "read_office.py"), "--help"],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        self.assertEqual(reader_help.returncode, 0, reader_help.stderr)
        self.assertIn("--legacy-timeout-seconds", reader_help.stdout)

        smoke_script = SCRIPTS_DIR / "smoke_office_reader.py"
        self.assertTrue(smoke_script.exists())
        smoke_help = subprocess.run(
            [sys.executable, str(smoke_script), "--help"],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        self.assertEqual(smoke_help.returncode, 0, smoke_help.stderr)
        self.assertIn("--doc", smoke_help.stdout)
        self.assertIn("--docx", smoke_help.stdout)
        self.assertIn("--ppt", smoke_help.stdout)
        self.assertIn("--pptx", smoke_help.stdout)
        self.assertIn("--derive-ppt-from-pptx", smoke_help.stdout)
        self.assertIn("--visual-timeout-seconds", smoke_help.stdout)

    def test_timeout_arguments_reject_non_positive_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "sample.docx"
            source.write_bytes(b"placeholder")
            cases = [
                [sys.executable, str(SCRIPTS_DIR / "read_office.py"), str(source), "--legacy-timeout-seconds", "0"],
                [sys.executable, str(SCRIPTS_DIR / "visual_analysis.py"), "manifest.json", "--normalized-file", str(source), "--out-dir", tmp, "--timeout-seconds", "0"],
                [sys.executable, str(SCRIPTS_DIR / "smoke_office_reader.py"), "--command-timeout-seconds", "0", "--skip-complete"],
                [sys.executable, str(SCRIPTS_DIR / "smoke_office_reader.py"), "--visual-timeout-seconds", "0", "--skip-complete"],
            ]
            for command in cases:
                with self.subTest(command=command):
                    proc = subprocess.run(
                        command,
                        text=True,
                        encoding="utf-8",
                        errors="replace",
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                        check=False,
                        timeout=30,
                    )
                    self.assertEqual(proc.returncode, 2)
                    self.assertIn("positive integer", proc.stderr)

        for script_name in ("convert_legacy_office.ps1", "render_preview.ps1"):
            script = (SCRIPTS_DIR / script_name).read_text(encoding="utf-8")
            self.assertIn("[ValidateRange(1, 86400)]", script)

    def test_smoke_harness_summarizes_modern_fast_runs(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            docx = tmp_path / "board-memo.docx"
            pptx = tmp_path / "标题幻灯片.pptx"
            make_docx(docx)
            make_pptx(pptx)
            proc = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPTS_DIR / "smoke_office_reader.py"),
                    "--docx",
                    str(docx),
                    "--pptx",
                    str(pptx),
                    "--out-dir",
                    str(tmp_path / "smoke"),
                    "--skip-complete",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            data = json.loads(proc.stdout)
            self.assertEqual(data["status"], "success")
            self.assertEqual(len(data["runs"]), 2)
            self.assertTrue(all(item["mode"] == "fast" for item in data["runs"]))
            self.assertTrue(all(item["manifest_summary"]["document_type"] in {"docx", "pptx"} for item in data["runs"]))
            self.assertEqual(data["runs"][1]["source"], str(pptx))

    def test_smoke_run_read_forwards_visual_timeout_to_unified_reader(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "sample.docx"
            source.write_bytes(b"source")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            manifest = out_dir / "sample.manifest.json"
            manifest.write_text(json.dumps({"visual_analysis": {}}), encoding="utf-8")

            def fake_run(command, timeout):
                self.assertIn("--visual-timeout-seconds", command)
                index = command.index("--visual-timeout-seconds")
                self.assertEqual(command[index + 1], "123")
                return subprocess.CompletedProcess(command, 0, json.dumps({"manifest": str(manifest)}), "")

            with patch.object(smoke_office_reader, "run_command", side_effect=fake_run):
                result = smoke_office_reader.run_read(source, out_dir, "complete", 45, 123, 300)

            self.assertEqual(result["status"], "success")

    def test_smoke_run_read_reports_malformed_reader_output_as_failed_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "sample.docx"
            source.write_bytes(b"source")
            cases = [
                subprocess.CompletedProcess(["python"], 0, "not-json", ""),
                subprocess.CompletedProcess(["python"], 0, json.dumps({}), ""),
            ]

            for completed in cases:
                with self.subTest(stdout=completed.stdout):
                    with patch.object(smoke_office_reader, "run_command", return_value=completed):
                        result = smoke_office_reader.run_read(source, tmp_path / "out", "fast", 1, 1, 1)

                    self.assertEqual(result["status"], "failed")
                    self.assertTrue(result["messages"])

    def test_smoke_run_read_reports_unusable_manifest_as_failed_run(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "sample.docx"
            source.write_bytes(b"source")
            out_dir = tmp_path / "out"
            out_dir.mkdir()
            outside_manifest = tmp_path / "outside.manifest.json"
            outside_manifest.write_text("{}", encoding="utf-8")
            non_object_manifest = out_dir / "non-object.manifest.json"
            non_object_manifest.write_text("[]", encoding="utf-8")

            for manifest in (outside_manifest, non_object_manifest):
                with self.subTest(manifest=manifest):
                    completed = subprocess.CompletedProcess(["python"], 0, json.dumps({"manifest": str(manifest)}), "")
                    with patch.object(smoke_office_reader, "run_command", return_value=completed):
                        result = smoke_office_reader.run_read(source, out_dir, "fast", 1, 1, 1)

                    self.assertEqual(result["status"], "failed")
                    self.assertTrue(result["messages"])

    def test_smoke_run_read_timeout_cleans_new_office_backend_processes(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "sample.docx"
            source.write_bytes(b"source")

            with (
                patch.object(smoke_office_reader, "office_backend_pids", return_value={10}),
                patch.object(smoke_office_reader, "stop_new_office_backend_processes") as stop,
                patch.object(
                    smoke_office_reader,
                    "run_command",
                    side_effect=subprocess.TimeoutExpired(["python"], timeout=1),
                ),
            ):
                result = smoke_office_reader.run_read(source, Path(tmp) / "out", "fast", 1, 1, 1)

            self.assertEqual(result["status"], "failed")
            stop.assert_called_once_with({10})

    def test_smoke_derive_ppt_preserves_existing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "deck.pptx"
            source.write_text("pptx placeholder", encoding="utf-8")
            out_dir = tmp_path / "derived"
            out_dir.mkdir()
            existing = out_dir / "deck.ppt"
            existing.write_text("user-owned derived fixture", encoding="utf-8")

            def fake_run(command, timeout):
                script = base64.b64decode(command[-1]).decode("utf-16le")
                match = re.search(r"\$presentation\.SaveAs\('([^']+)', 1\)", script)
                self.assertIsNotNone(match)
                target = Path(match.group(1))
                target.write_text("derived fixture", encoding="utf-8")
                return subprocess.CompletedProcess(command, 0, "", "")

            with patch.object(smoke_office_reader, "run_command", side_effect=fake_run):
                derived = smoke_office_reader.derive_ppt(source, out_dir, timeout=1)

            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned derived fixture")
            self.assertNotEqual(derived, existing)
            self.assertTrue(derived.exists())

    def test_smoke_derive_ppt_timeout_cleans_new_powerpoint_automation(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "deck.pptx"
            source.write_text("pptx placeholder", encoding="utf-8")
            before = {101}

            with (
                patch.object(smoke_office_reader, "powerpoint_automation_pids", return_value=before),
                patch.object(smoke_office_reader, "stop_new_powerpoint_automation_processes") as cleanup,
                patch.object(smoke_office_reader, "run_command", side_effect=subprocess.TimeoutExpired(["powershell"], 1)),
            ):
                with self.assertRaises(subprocess.TimeoutExpired):
                    smoke_office_reader.derive_ppt(source, tmp_path / "derived", timeout=1)

            cleanup.assert_called_once_with(before)

    def test_smoke_derive_ppt_timeout_removes_partial_generated_target(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "deck.pptx"
            source.write_text("pptx placeholder", encoding="utf-8")
            out_dir = tmp_path / "derived"

            def fake_run(command, timeout):
                script = base64.b64decode(command[-1]).decode("utf-16le")
                match = re.search(r"\$presentation\.SaveAs\('([^']+)', 1\)", script)
                self.assertIsNotNone(match)
                target = Path(match.group(1))
                target.write_text("partial derived fixture", encoding="utf-8")
                raise subprocess.TimeoutExpired(command, timeout)

            with (
                patch.object(smoke_office_reader, "powerpoint_automation_pids", return_value=set()),
                patch.object(smoke_office_reader, "stop_new_powerpoint_automation_processes"),
                patch.object(smoke_office_reader, "run_command", side_effect=fake_run),
            ):
                with self.assertRaises(subprocess.TimeoutExpired):
                    smoke_office_reader.derive_ppt(source, out_dir, timeout=1)

            self.assertEqual(list(out_dir.glob("*.ppt")), [])

    def test_render_preview_failure_for_missing_file_is_structured_json(self):
        script = SCRIPTS_DIR / "render_preview.ps1"
        proc = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script),
                "-InputPath",
                str(Path("C:/definitely/missing.docx")),
                "-OutputDir",
                str(Path(tempfile.gettempdir())),
            ],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        self.assertNotEqual(proc.returncode, 0)
        data = json.loads(proc.stdout)
        self.assertEqual(data["status"], "failed")
        self.assertIn("does not exist", data["messages"][0])

    def test_render_preview_failure_preserves_unicode_path_in_json(self):
        missing = Path(tempfile.gettempdir()) / "不存在.docx"
        proc = subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(SCRIPTS_DIR / "render_preview.ps1"),
                "-InputPath",
                str(missing),
                "-OutputDir",
                tempfile.gettempdir(),
            ],
            text=True,
            encoding="utf-8",
            errors="strict",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
        self.assertNotEqual(proc.returncode, 0)
        data = json.loads(proc.stdout)
        self.assertIn(str(missing), data["messages"][0])

    def test_render_preview_accepts_timeout_parameter(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            make_docx(source)
            script = SCRIPTS_DIR / "render_preview.ps1"

            proc = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script),
                    "-InputPath",
                    str(source),
                    "-OutputDir",
                    str(out_dir),
                    "-TimeoutSeconds",
                    "1",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
            )
            self.assertIn(proc.returncode, {0, 1}, proc.stderr)
            data = json.loads(proc.stdout)
            self.assertIn(data["status"], {"success", "failed"})
            self.assertIn("messages", data)

    def test_render_preview_timeout_cleanup_is_limited_to_worker_process_tree(self):
        script_text = (SCRIPTS_DIR / "render_preview.ps1").read_text(encoding="utf-8")

        self.assertIn("function Stop-ProcessTree", script_text)
        self.assertGreaterEqual(script_text.count("Stop-ProcessTree -RootProcessId $process.Id"), 2)
        self.assertGreaterEqual(script_text.count("$automationBefore = @(Get-AutomationProcessIds)"), 2)
        self.assertGreaterEqual(script_text.count("Stop-NewAutomationProcesses -Before $automationBefore"), 2)
        self.assertNotIn("Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='POWERPNT.EXE'\"", script_text)

    def test_render_preview_removes_partial_pdf_before_fallback(self):
        script_text = (SCRIPTS_DIR / "render_preview.ps1").read_text(encoding="utf-8")

        self.assertIn("function Remove-PartialPreview", script_text)
        fallback_index = script_text.index("$soffice = Resolve-LibreOfficeExecutable")
        cleanup_index = script_text.rfind("Remove-PartialPreview -Path $pdfPath", 0, fallback_index)
        self.assertGreater(cleanup_index, 0)
        self.assertGreaterEqual(script_text.count("Remove-PartialPreview -Path $pdfPath"), 2)

    def test_legacy_timeout_cleanup_handles_reparented_office_automation(self):
        script_text = (SCRIPTS_DIR / "convert_legacy_office.ps1").read_text(encoding="utf-8")

        self.assertIn("$automationBefore = @(Get-AutomationProcessIds)", script_text)
        self.assertIn("Stop-NewAutomationProcesses -Before $automationBefore", script_text)

    def test_legacy_conversion_removes_partial_normalized_output_after_failed_worker(self):
        script_text = (SCRIPTS_DIR / "convert_legacy_office.ps1").read_text(encoding="utf-8")

        self.assertIn("function Remove-PartialNormalizedOutput", script_text)
        self.assertGreaterEqual(script_text.count("Remove-PartialNormalizedOutput -Path $output"), 2)
        failure_cleanup = script_text.index("Remove-PartialNormalizedOutput -Path $output")
        failure_messages = script_text.index("foreach ($message in @($result.messages))")
        self.assertLess(failure_cleanup, failure_messages)

    def test_render_preview_com_worker_failure_can_continue_to_later_backends(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            fake_soffice = tmp_path / "soffice.cmd"
            make_docx(source)
            fake_soffice.write_text(
                "@echo off\r\n"
                "set OUTDIR=\r\n"
                ":args\r\n"
                "if \"%~1\"==\"\" goto doneargs\r\n"
                "if \"%~1\"==\"--outdir\" (\r\n"
                "  set OUTDIR=%~2\r\n"
                "  shift\r\n"
                ")\r\n"
                "shift\r\n"
                "goto args\r\n"
                ":doneargs\r\n"
                "if not defined OUTDIR exit /b 2\r\n"
                "echo fake pdf> \"%OUTDIR%\\board-memo.pdf\"\r\n"
                "exit /b 0\r\n",
                encoding="utf-8",
            )
            script = SCRIPTS_DIR / "render_preview.ps1"
            env = os.environ.copy()
            env["SOFFICE_PATH"] = str(fake_soffice)

            proc = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(script),
                    "-InputPath",
                    str(source),
                    "-OutputDir",
                    str(out_dir),
                    "-TimeoutSeconds",
                    "1",
                    "-ContinueAfterComFailure",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
                env=env,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            data = json.loads(proc.stdout)
            self.assertEqual(data["status"], "success")
            self.assertEqual(data["backend"], "libreoffice")
            self.assertIn("messages", data)
            self.assertTrue((out_dir / "board-memo.pdf").exists())
            self.assertFalse(list(out_dir.glob("preview-worker-*")))
            self.assertFalse(list(out_dir.glob(".lo_profile_*")))

    def test_render_preview_preserves_existing_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            out_dir.mkdir()
            existing = out_dir / "board-memo.pdf"
            existing.write_text("user-owned preview", encoding="utf-8")
            fake_soffice = tmp_path / "soffice.cmd"
            make_docx(source)
            fake_soffice.write_text(
                "@echo off\r\n"
                "set OUTDIR=\r\n"
                ":args\r\n"
                "if \"%~1\"==\"\" goto doneargs\r\n"
                "if \"%~1\"==\"--outdir\" (\r\n"
                "  set OUTDIR=%~2\r\n"
                "  shift\r\n"
                ")\r\n"
                "shift\r\n"
                "goto args\r\n"
                ":doneargs\r\n"
                "echo fake pdf> \"%OUTDIR%\\board-memo.pdf\"\r\n"
                "exit /b 0\r\n",
                encoding="utf-8",
            )
            env = os.environ.copy()
            env["SOFFICE_PATH"] = str(fake_soffice)

            proc = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    str(SCRIPTS_DIR / "render_preview.ps1"),
                    "-InputPath",
                    str(source),
                    "-OutputDir",
                    str(out_dir),
                    "-TimeoutSeconds",
                    "1",
                    "-ContinueAfterComFailure",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
                env=env,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr)
            data = json.loads(proc.stdout)
            artifact = Path(data["artifacts"][0])
            self.assertEqual(existing.read_text(encoding="utf-8"), "user-owned preview")
            self.assertNotEqual(artifact, existing)
            self.assertTrue(artifact.exists())


if __name__ == "__main__":
    unittest.main()
