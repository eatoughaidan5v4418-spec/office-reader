import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
import os
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = SKILL_DIR / "scripts"


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
      <w:tr>
        <w:tc><w:p><w:r><w:t>ARR</w:t></w:r></w:p></w:tc>
        <w:tc><w:p>
          <w:r><w:t>$10M</w:t></w:r>
          <w:ins w:author="Table Reviewer" w:date="2026-01-05T00:00:00Z"><w:r><w:t> verified</w:t></w:r></w:ins>
          <w:del w:author="Table Reviewer" w:date="2026-01-06T00:00:00Z"><w:r><w:delText> draft</w:delText></w:r></w:del>
          <w:r><w:commentReference w:id="1"/></w:r>
          <w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImage2"/></w:drawing></w:r>
        </w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:drawing><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" r:embed="rIdImage1"/></w:drawing></w:r></w:p>
  </w:body>
</w:document>""",
        "word/comments.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="0" w:author="Reviewer" w:date="2026-01-04T00:00:00Z">
    <w:p><w:r><w:t>Please verify the ARR source.</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="1" w:author="Table Reviewer" w:date="2026-01-05T00:00:00Z">
    <w:p><w:r><w:t>Table value needs audit trail.</w:t></w:r></w:p>
  </w:comment>
</w:comments>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
  <Relationship Id="rIdImage2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.png"/>
</Relationships>""",
        "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Board Memo</dc:title>
  <dc:creator>Finance Team</dc:creator>
</cp:coreProperties>""",
    }
    write_zip(path, files)


def make_docx_with_supplemental_parts(path):
    files = {
        "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body ok</w:t></w:r></w:p>
    <w:p><w:r><w:footnoteReference w:id="2"/></w:r><w:r><w:endnoteReference w:id="3"/></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader1"/>
      <w:footerReference w:type="default" r:id="rIdFooter1"/>
    </w:sectPr>
  </w:body>
</w:document>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>""",
        "word/header1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:r><w:t>CONFIDENTIAL HEADER</w:t></w:r>
    <w:r><w:drawing><a:blip r:embed="rIdHeaderImage"/></w:drawing></w:r>
  </w:p>
</w:hdr>""",
        "word/_rels/header1.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeaderImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/header.png"/>
</Relationships>""",
        "word/footer1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>v2 footer</w:t></w:r></w:p>
</w:ftr>""",
        "word/footnotes.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="-1"><w:p><w:r><w:t>separator</w:t></w:r></w:p></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>Footnote evidence</w:t></w:r></w:p></w:footnote>
</w:footnotes>""",
        "word/endnotes.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="3"><w:p><w:r><w:t>Endnote source</w:t></w:r></w:p></w:endnote>
</w:endnotes>""",
        "word/media/header.png": b"fake image bytes",
    }
    write_zip(path, files)


def make_docx_with_content_control(path):
    files = {
        "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:sdt>
      <w:sdtPr><w:tag w:val="approval"/></w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r><w:t>Controlled approval text </w:t></w:r>
          <w:ins w:author="Approver" w:date="2026-03-01T00:00:00Z"><w:r><w:t>accepted</w:t></w:r></w:ins>
          <w:r><w:commentReference w:id="2"/></w:r>
          <w:r><w:drawing><a:blip r:embed="rIdControlledImage"/></w:drawing></w:r>
        </w:p>
        <w:tbl>
          <w:tr><w:tc><w:p><w:r><w:t>Field</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>
          <w:tr><w:tc><w:p><w:r><w:t>Decision</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Go</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>
      </w:sdtContent>
    </w:sdt>
  </w:body>
</w:document>""",
        "word/comments.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="2" w:author="Template Reviewer" w:date="2026-03-02T00:00:00Z">
    <w:p><w:r><w:t>Controlled text needs signoff.</w:t></w:r></w:p>
  </w:comment>
</w:comments>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdControlledImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/controlled.png"/>
</Relationships>""",
        "word/media/controlled.png": b"fake image bytes",
    }
    write_zip(path, files)


def make_docx_with_textbox(path):
    files = {
        "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r><w:t>Outer paragraph only.</w:t></w:r>
      <w:r><w:drawing><wps:txbx><w:txbxContent>
        <w:p>
          <w:r><w:t>Textbox risk note </w:t></w:r>
          <w:ins w:author="Shape Reviewer" w:date="2026-04-01T00:00:00Z"><w:r><w:t>approved</w:t></w:r></w:ins>
          <w:r><w:commentReference w:id="4"/></w:r>
          <w:r><w:drawing><a:blip r:embed="rIdTextboxImage"/></w:drawing></w:r>
        </w:p>
      </w:txbxContent></wps:txbx></w:drawing></w:r>
    </w:p>
  </w:body>
</w:document>""",
        "word/comments.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="4" w:author="Shape Reviewer" w:date="2026-04-02T00:00:00Z">
    <w:p><w:r><w:t>Textbox needs visual review.</w:t></w:r></w:p>
  </w:comment>
</w:comments>""",
        "word/_rels/document.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTextboxImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/textbox.png"/>
</Relationships>""",
        "word/media/textbox.png": b"fake image bytes",
    }
    write_zip(path, files)


def make_docx_with_comment_range(path):
    files = {
        "word/document.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Please verify </w:t></w:r>
      <w:commentRangeStart w:id="5"/>
      <w:r><w:t>ARR source</w:t></w:r>
      <w:commentRangeEnd w:id="5"/>
      <w:r><w:t> before release.</w:t></w:r>
      <w:r><w:commentReference w:id="5"/></w:r>
    </w:p>
  </w:body>
</w:document>""",
        "word/comments.xml": """<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="5" w:author="Reviewer" w:date="2026-05-01T00:00:00Z">
    <w:p><w:r><w:t>Add source citation.</w:t></w:r></w:p>
  </w:comment>
</w:comments>""",
    }
    write_zip(path, files)


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
    <p:sp><p:txBody><a:p><a:r><a:t>Three blockers remain.</a:t></a:r></a:p></p:txBody></p:sp>
    <p:graphicFrame><a:graphic><a:graphicData><a:tbl>
      <a:tr><a:tc><a:txBody><a:p><a:r><a:t>Owner</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>Status</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
      <a:tr><a:tc><a:txBody><a:p><a:r><a:t>Ops</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:p><a:r><a:t>At risk</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
    </a:tbl></a:graphicData></a:graphic></p:graphicFrame>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="7" name="Hero image" descr="Revenue dashboard screenshot"/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImage1"/></p:blipFill>
      <p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="1828800"/></a:xfrm></p:spPr>
    </p:pic>
    <p:graphicFrame>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart1"/>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
  </p:spTree></p:cSld>
</p:sld>""",
        "ppt/slides/_rels/slide1.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdNotes1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
  <Relationship Id="rIdComments1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>
  <Relationship Id="rIdImage1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
  <Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>""",
        "ppt/notesSlides/notesSlide1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Ask support to confirm staffing.</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld>
</p:notes>""",
        "ppt/comments/comment1.xml": """<?xml version="1.0" encoding="UTF-8"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
         xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cm authorId="0" dt="2026-02-01T00:00:00Z"><p:text>Clarify launch date.</p:text></p:cm>
</p:cmLst>""",
        "docProps/core.xml": """<?xml version="1.0" encoding="UTF-8"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Launch Deck</dc:title>
</cp:coreProperties>""",
    }
    write_zip(path, files)


def make_pptx_with_complex_visual_objects(path):
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
       xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:sp><p:txBody><a:p><a:r><a:t>Complex Objects</a:t></a:r></a:p></p:txBody></p:sp>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="4" name="Process SmartArt" descr="Approval workflow"/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
        <dgm:relIds r:dm="rIdDgmData" r:lo="rIdDgmLayout" r:qs="rIdDgmQuickStyle" r:cs="rIdDgmColors"/>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
    <p:graphicFrame>
      <p:nvGraphicFramePr><p:cNvPr id="5" name="Embedded workbook"/></p:nvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
        <p:oleObj r:id="rIdOle1" progId="Excel.Sheet.12" name="Workbook data"/>
      </a:graphicData></a:graphic>
    </p:graphicFrame>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="6" name="Video poster"/><p:nvPr><a:videoFile r:link="rIdVideo1"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdPoster"/></p:blipFill>
    </p:pic>
    <p:pic>
      <p:nvPicPr><p:cNvPr id="7" name="Audio icon"/><p:nvPr><a:wavAudioFile r:embed="rIdAudio1"/></p:nvPr></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdAudioIcon"/></p:blipFill>
    </p:pic>
  </p:spTree></p:cSld>
</p:sld>""",
        "ppt/slides/_rels/slide1.xml.rels": """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDgmData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
  <Relationship Id="rIdDgmLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>
  <Relationship Id="rIdDgmQuickStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
  <Relationship Id="rIdDgmColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/colors1.xml"/>
  <Relationship Id="rIdOle1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
  <Relationship Id="rIdVideo1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/movie1.mp4" TargetMode="External"/>
  <Relationship Id="rIdAudio1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="../media/audio1.wav"/>
  <Relationship Id="rIdPoster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/poster.png"/>
  <Relationship Id="rIdAudioIcon" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/audio.png"/>
</Relationships>""",
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
            self.assertEqual(manifest["tables"][0]["rows"][1], ["ARR", "$10M verified draft"])
            table_comment = next(item for item in manifest["comments"] if item["id"] == "1")
            self.assertEqual(table_comment["table_index"], 1)
            self.assertEqual(table_comment["row_index"], 2)
            self.assertEqual(table_comment["cell_index"], 2)
            table_revision = next(item for item in manifest["revisions"] if item["text"] == "verified")
            self.assertEqual(table_revision["table_index"], 1)
            self.assertEqual(table_revision["row_index"], 2)
            self.assertEqual(table_revision["cell_index"], 2)
            relationships = manifest["visual_findings"][0]["relationships"]
            table_media = next(item for item in relationships if item["relationship_id"] == "rIdImage2")
            self.assertEqual(table_media["table_index"], 1)
            self.assertEqual(table_media["row_index"], 2)
            self.assertEqual(table_media["cell_index"], 2)
            self.assertTrue(manifest["visual_findings"][0]["requires_visual_review"])

    def test_docx_reader_extracts_headers_footers_footnotes_and_endnotes(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "supplemental.docx"
            out_dir = tmp_path / "out"
            make_docx_with_supplemental_parts(source)

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            full_md = (out_dir / "supplemental.full.md").read_text(encoding="utf-8")
            manifest = json.loads((out_dir / "supplemental.manifest.json").read_text(encoding="utf-8"))
            self.assertIn("Body ok", full_md)
            self.assertIn("CONFIDENTIAL HEADER", full_md)
            self.assertIn("v2 footer", full_md)
            self.assertIn("Footnote evidence", full_md)
            self.assertIn("Endnote source", full_md)

            by_text = {item["text"]: item for item in manifest["structure"]}
            self.assertEqual(by_text["CONFIDENTIAL HEADER"]["part_type"], "header")
            self.assertEqual(by_text["CONFIDENTIAL HEADER"]["part"], "word/header1.xml")
            self.assertEqual(by_text["v2 footer"]["part_type"], "footer")
            self.assertEqual(by_text["Footnote evidence"]["part_type"], "footnote")
            self.assertEqual(by_text["Endnote source"]["part_type"], "endnote")

            relationships = manifest["visual_findings"][0]["relationships"]
            header_media = next(item for item in relationships if item["relationship_id"] == "rIdHeaderImage")
            self.assertEqual(header_media["target"], "word/media/header.png")
            self.assertEqual(header_media["part_type"], "header")
            self.assertEqual(header_media["part"], "word/header1.xml")

    def test_docx_reader_extracts_block_level_content_controls(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "content-control.docx"
            out_dir = tmp_path / "out"
            make_docx_with_content_control(source)

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            full_md = (out_dir / "content-control.full.md").read_text(encoding="utf-8")
            manifest = json.loads((out_dir / "content-control.manifest.json").read_text(encoding="utf-8"))
            self.assertIn("Controlled approval text", full_md)
            self.assertIn("| Field | Value |", full_md)

            controlled = next(item for item in manifest["structure"] if "Controlled approval text" in item["text"])
            self.assertEqual(controlled["container"], "content_control")
            self.assertEqual(controlled["part_type"], "document")

            table = next(item for item in manifest["tables"] if item["rows"][1] == ["Decision", "Go"])
            self.assertEqual(table["container"], "content_control")

            revision = next(item for item in manifest["revisions"] if item["text"] == "accepted")
            self.assertEqual(revision["container"], "content_control")

            comment = next(item for item in manifest["comments"] if item["id"] == "2")
            self.assertEqual(comment["container"], "content_control")
            self.assertEqual(comment["text"], "Controlled text needs signoff.")

            relationships = manifest["visual_findings"][0]["relationships"]
            controlled_media = next(item for item in relationships if item["relationship_id"] == "rIdControlledImage")
            self.assertEqual(controlled_media["container"], "content_control")
            self.assertEqual(controlled_media["target"], "word/media/controlled.png")

    def test_docx_reader_extracts_textbox_content_with_origin(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "textbox.docx"
            out_dir = tmp_path / "out"
            make_docx_with_textbox(source)

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            full_md = (out_dir / "textbox.full.md").read_text(encoding="utf-8")
            manifest = json.loads((out_dir / "textbox.manifest.json").read_text(encoding="utf-8"))
            self.assertIn("Outer paragraph only.", full_md)
            self.assertIn("Textbox risk note", full_md)

            outer = next(item for item in manifest["structure"] if item["text"] == "Outer paragraph only.")
            self.assertNotIn("container", outer)
            self.assertNotIn("Textbox risk note", outer["text"])

            textbox = next(item for item in manifest["structure"] if "Textbox risk note" in item["text"])
            self.assertEqual(textbox["container"], "textbox")
            self.assertIn("{+approved+}", textbox["text"])

            revision = next(item for item in manifest["revisions"] if item["text"] == "approved")
            self.assertEqual(revision["container"], "textbox")

            comment = next(item for item in manifest["comments"] if item["id"] == "4")
            self.assertEqual(comment["container"], "textbox")
            self.assertEqual(comment["text"], "Textbox needs visual review.")

            relationships = manifest["visual_findings"][0]["relationships"]
            textbox_media = next(item for item in relationships if item["relationship_id"] == "rIdTextboxImage")
            self.assertEqual(textbox_media["container"], "textbox")
            self.assertEqual(textbox_media["target"], "word/media/textbox.png")

    def test_docx_reader_extracts_comment_range_anchor_text(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "comment-range.docx"
            out_dir = tmp_path / "out"
            make_docx_with_comment_range(source)

            self.run_script("read_docx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "comment-range.manifest.json").read_text(encoding="utf-8"))
            comment = next(item for item in manifest["comments"] if item["id"] == "5")
            self.assertEqual(comment["text"], "Add source citation.")
            self.assertEqual(comment["anchor_text"], "ARR source")
            self.assertEqual(comment["paragraph_index"], 1)

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
            self.assertEqual(manifest["notes"][0]["text"], "Ask support to confirm staffing.")
            self.assertEqual(manifest["comments"][0]["text"], "Clarify launch date.")
            self.assertEqual(manifest["tables"][0]["rows"][1], ["Ops", "At risk"])
            finding = manifest["visual_findings"][0]
            objects = finding["objects"]
            image = next(item for item in objects if item["object_type"] == "image")
            self.assertEqual(image["name"], "Hero image")
            self.assertEqual(image["alt_text"], "Revenue dashboard screenshot")
            self.assertEqual(image["relationship_id"], "rIdImage1")
            self.assertEqual(image["target"], "ppt/media/image1.png")
            self.assertEqual(image["geometry"], {"x": 914400, "y": 1828800, "cx": 3657600, "cy": 1828800})
            chart = next(item for item in objects if item["object_type"] == "chart")
            self.assertEqual(chart["relationship_id"], "rIdChart1")
            self.assertEqual(chart["target"], "ppt/charts/chart1.xml")

    def test_pptx_reader_inventories_complex_visual_objects(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "complex-objects.pptx"
            out_dir = tmp_path / "out"
            make_pptx_with_complex_visual_objects(source)

            self.run_script("read_pptx.py", source, "--out-dir", out_dir)

            manifest = json.loads((out_dir / "complex-objects.manifest.json").read_text(encoding="utf-8"))
            finding = manifest["visual_findings"][0]
            self.assertTrue(finding["requires_visual_review"])
            objects = finding["objects"]
            object_types = {item["object_type"] for item in objects}
            self.assertTrue({"smartart", "ole", "video", "audio"}.issubset(object_types))

            smartart = next(item for item in objects if item["object_type"] == "smartart")
            self.assertEqual(smartart["name"], "Process SmartArt")
            self.assertEqual(smartart["alt_text"], "Approval workflow")
            self.assertEqual(smartart["geometry"], {"x": 1000, "y": 2000, "cx": 3000, "cy": 4000})
            self.assertEqual(
                {rel["role"]: rel["target"] for rel in smartart["relationships"]},
                {
                    "data_model": "ppt/diagrams/data1.xml",
                    "layout": "ppt/diagrams/layout1.xml",
                    "quick_style": "ppt/diagrams/quickStyle1.xml",
                    "colors": "ppt/diagrams/colors1.xml",
                },
            )

            ole = next(item for item in objects if item["object_type"] == "ole")
            self.assertEqual(ole["name"], "Embedded workbook")
            self.assertEqual(ole["prog_id"], "Excel.Sheet.12")
            self.assertEqual(ole["relationship_id"], "rIdOle1")
            self.assertEqual(ole["target"], "ppt/embeddings/oleObject1.bin")

            video = next(item for item in objects if item["object_type"] == "video")
            self.assertEqual(video["relationship_id"], "rIdVideo1")
            self.assertEqual(video["target"], "ppt/media/movie1.mp4")
            self.assertEqual(video["target_mode"], "External")

            audio = next(item for item in objects if item["object_type"] == "audio")
            self.assertEqual(audio["relationship_id"], "rIdAudio1")
            self.assertEqual(audio["target"], "ppt/media/audio1.wav")

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
                        "comments": [{"text": "Clarify launch date.", "anchor_text": "launch date"}],
                        "revisions": [],
                        "notes": [{"slide_index": 1, "text": "Ask support to confirm staffing."}],
                        "visual_findings": [
                            {
                                "requires_visual_review": True,
                                "reason": "slide has media",
                                "slide_index": 1,
                                "objects": [
                                    {
                                        "object_type": "smartart",
                                        "name": "Process SmartArt",
                                        "relationships": [
                                            {"role": "data_model", "target": "ppt/diagrams/data1.xml"}
                                        ],
                                    },
                                    {
                                        "object_type": "ole",
                                        "name": "Embedded workbook",
                                        "prog_id": "Excel.Sheet.12",
                                        "relationship_id": "rIdOle1",
                                        "target": "ppt/embeddings/oleObject1.bin",
                                    },
                                ],
                            }
                        ],
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
            self.assertIn("Clarify launch date.", report)
            self.assertIn("on 'launch date'", report)
            self.assertIn("Object: smartart", report)
            self.assertIn("data_model=ppt/diagrams/data1.xml", report)
            self.assertIn("prog_id=Excel.Sheet.12", report)

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
            self.assertIn("- Comments: 2", report)

    def test_unified_reader_stdout_json_is_utf8_safe_for_chinese_paths(self):
        with tempfile.TemporaryDirectory(prefix="office-reader-") as tmp:
            tmp_path = Path(tmp) / "\u4e2d\u6587 path smoke"
            tmp_path.mkdir()
            source = tmp_path / "\u4e2d\u6587 \u6837\u672c.docx"
            out_dir = tmp_path / "\u8f93\u51fa docx"
            make_docx(source)

            script = SCRIPTS_DIR / "read_office.py"
            proc = subprocess.run(
                [
                    sys.executable,
                    str(script),
                    str(source),
                    "--out-dir",
                    str(out_dir),
                    "--mode",
                    "fast",
                    "--no-openai-vision",
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )
            self.assertEqual(proc.returncode, 0, proc.stderr.decode(errors="replace"))
            stdout = proc.stdout.decode("utf-8")
            data = json.loads(stdout)
            self.assertTrue(Path(data["manifest"]).exists())
            self.assertTrue(Path(data["report"]).exists())

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

    def test_visual_analysis_reuses_existing_preview_pdf_without_overwriting(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "out"
            make_docx(source)

            self.run_script("read_docx.py", source, "--out-dir", out_dir)
            preview_dir = out_dir / "preview"
            preview_dir.mkdir()
            preview_pdf = preview_dir / "board-memo.pdf"
            preview_pdf.write_bytes(
                b"%PDF-1.4\n"
                b"1 0 obj<<>>endobj\n"
                b"2 0 obj<< /Type /Catalog /Pages 3 0 R >>endobj\n"
                b"3 0 obj<< /Type /Pages /Count 0 >>endobj\n"
                b"trailer<< /Root 2 0 R >>\n%%EOF\n"
            )
            before = preview_pdf.stat().st_mtime_ns

            self.run_script(
                "visual_analysis.py",
                out_dir / "board-memo.manifest.json",
                "--normalized-file",
                source,
                "--out-dir",
                out_dir,
                "--mode",
                "complete",
                "--no-openai-vision",
                "--timeout-seconds",
                "1",
            )

            manifest = json.loads((out_dir / "board-memo.manifest.json").read_text(encoding="utf-8"))
            messages = " ".join(manifest["visual_analysis"]["messages"]).lower()
            self.assertIn("reusing existing preview pdf", messages)
            self.assertEqual(manifest["artifacts"]["preview_pdf"], str(preview_pdf))
            self.assertEqual(preview_pdf.stat().st_mtime_ns, before)

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
            score = manifest["completeness_score"]
            self.assertLess(score["overall"], 80)
            self.assertEqual(score["visual_coverage"], 0)
            self.assertGreaterEqual(score["unverified_visual_count"], 1)
            self.assertFalse(score["openai_vision_enabled"])

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

    def test_report_includes_completeness_score_when_manifest_has_it(self):
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
                        "tables": [{"index": 1, "rows": [["Metric", "Value"]]}],
                        "comments": [],
                        "revisions": [],
                        "notes": [],
                        "visual_analysis": {"status": "partial", "mode": "balanced", "messages": []},
                        "visual_findings": [{"requires_visual_review": True, "vision_summary": ""}],
                        "completeness_score": {
                            "overall": 72,
                            "text_coverage": 100,
                            "table_coverage": 100,
                            "visual_coverage": 25,
                            "ocr_confidence": 0,
                            "openai_vision_enabled": False,
                            "unverified_visual_count": 1,
                            "signals": ["Visual analysis was partial."],
                        },
                        "artifacts": {"full_markdown": "sample.full.md"},
                    }
                ),
                encoding="utf-8",
            )

            self.run_script("assemble_report.py", manifest_path, "--out", report_path)

            report = report_path.read_text(encoding="utf-8")
            self.assertIn("- Completeness score: 72/100", report)
            self.assertIn("## Read Completeness", report)
            self.assertIn("- Visual coverage: 25/100", report)
            self.assertIn("- Unverified visual items: 1", report)

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
            self.assertEqual(list(out_dir.glob(".lo_profile_*")), [])

    def test_render_preview_skips_unhealthy_com_backend_from_health_memory(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            health_path = tmp_path / "backend-health.json"
            fake_soffice = tmp_path / "soffice.cmd"
            make_docx(source)
            health_path.write_text(
                json.dumps(
                    {
                        "preview": {
                            ".docx": {
                                "office-com": {
                                    "state": "unhealthy",
                                    "reason": "timeout",
                                    "timeout_seconds": 1,
                                    "updated_at": "2026-06-02T00:00:00Z",
                                }
                            }
                        }
                    }
                ),
                encoding="utf-8",
            )
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
            env["OFFICE_READER_PREVIEW_HEALTH_PATH"] = str(health_path)

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
                    "5",
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
            messages = " ".join(data["messages"])
            self.assertIn("Skipping Office COM preview", messages)
            self.assertNotIn("timed out after 5 seconds", messages)

    def test_render_preview_timeout_records_unhealthy_com_backend(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            health_path = tmp_path / "backend-health.json"
            make_docx(source)
            script = SCRIPTS_DIR / "render_preview.ps1"
            env = os.environ.copy()
            env["OFFICE_READER_PREVIEW_HEALTH_PATH"] = str(health_path)

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
            self.assertIn(proc.returncode, {0, 1}, proc.stderr)
            data = json.loads(proc.stdout)
            messages = " ".join(data.get("messages", []))
            if "Office COM preview timed out" not in messages:
                self.skipTest("Office COM preview did not time out on this machine.")
            health = json.loads(health_path.read_text(encoding="utf-8"))
            entry = health["preview"][".docx"]["office-com"]
            self.assertEqual(entry["state"], "unhealthy")
            self.assertEqual(entry["reason"], "timeout")
            self.assertEqual(entry["timeout_seconds"], 1)

    def test_render_preview_rejects_health_path_that_would_overwrite_regular_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            protected_file = tmp_path / "ordinary-file.txt"
            make_docx(source)
            protected_file.write_text("do not overwrite", encoding="utf-8")
            script = SCRIPTS_DIR / "render_preview.ps1"
            env = os.environ.copy()
            env["OFFICE_READER_PREVIEW_HEALTH_PATH"] = str(protected_file)

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
            self.assertNotEqual(proc.returncode, 0)
            data = json.loads(proc.stdout)
            self.assertEqual(data["status"], "failed")
            self.assertIn("preview health path", " ".join(data["messages"]).lower())
            self.assertEqual(protected_file.read_text(encoding="utf-8"), "do not overwrite")

    def test_render_preview_refuses_to_overwrite_existing_pdf(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "board-memo.docx"
            out_dir = tmp_path / "preview"
            existing_pdf = out_dir / "board-memo.pdf"
            make_docx(source)
            out_dir.mkdir()
            existing_pdf.write_text("existing preview", encoding="utf-8")
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
                    "-ContinueAfterComFailure",
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
            )
            self.assertNotEqual(proc.returncode, 0)
            data = json.loads(proc.stdout)
            self.assertEqual(data["status"], "failed")
            self.assertIn("already exists", " ".join(data["messages"]).lower())
            self.assertEqual(existing_pdf.read_text(encoding="utf-8"), "existing preview")

    def test_legacy_conversion_refuses_to_overwrite_existing_normalized_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "legacy.doc"
            out_dir = tmp_path / "converted"
            existing_docx = out_dir / "legacy.docx"
            source.write_bytes(b"not a real legacy document")
            out_dir.mkdir()
            existing_docx.write_text("existing normalized document", encoding="utf-8")
            script = SCRIPTS_DIR / "convert_legacy_office.ps1"

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
                ],
                text=True,
                encoding="utf-8",
                errors="replace",
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
                timeout=30,
            )
            self.assertNotEqual(proc.returncode, 0)
            data = json.loads(proc.stdout)
            self.assertEqual(data["status"], "failed")
            self.assertIn("already exists", " ".join(data["messages"]).lower())
            self.assertEqual(existing_docx.read_text(encoding="utf-8"), "existing normalized document")

    def test_legacy_conversion_cleans_libreoffice_profile_directory(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            source = tmp_path / "legacy.doc"
            out_dir = tmp_path / "converted"
            fake_soffice = tmp_path / "soffice.cmd"
            source.write_bytes(b"not a real legacy document")
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
                "echo fake docx> \"%OUTDIR%\\legacy.docx\"\r\n"
                "exit /b 0\r\n",
                encoding="utf-8",
            )
            script = SCRIPTS_DIR / "convert_legacy_office.ps1"
            env = os.environ.copy()
            env["PATH"] = str(tmp_path) + os.pathsep + env.get("PATH", "")

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
            self.assertEqual(data["status"], "success")
            self.assertEqual(data["backend"], "libreoffice")
            self.assertTrue((out_dir / "legacy.docx").exists())
            self.assertEqual(list(out_dir.glob(".lo_profile_*")), [])


if __name__ == "__main__":
    unittest.main()
