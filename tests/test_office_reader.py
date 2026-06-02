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
            self.assertEqual(manifest["tables"][0]["rows"][1], ["ARR", "$10M"])
            self.assertTrue(manifest["visual_findings"][0]["requires_visual_review"])

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
            self.assertIn("Clarify launch date.", report)

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


if __name__ == "__main__":
    unittest.main()
