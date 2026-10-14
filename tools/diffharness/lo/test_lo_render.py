import hashlib
import json
from pathlib import Path
import tempfile
import sys
import time
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import lo_render


class RenderTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.source = self.folder / "tracked.docx"
        with ZipFile(self.source, "w") as package:
            package.writestr("word/header1.xml", '''
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
 xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
 <w:p><w:ins w:id="7"><w:r><w:drawing><wp:anchor><wpg:wgp><wps:wsp>
 <wps:txbx/></wps:wsp></wpg:wgp></wp:anchor></w:drawing></w:r></w:ins></w:p>
</w:hdr>''')
        self.original = self.source.read_bytes()
        self.output = self.folder / "pdf"

    def run_render(self, conversion):
        calls = []

        def execute(arguments, timeout):
            calls.append(arguments)
            if "--version" in arguments:
                return dict(returncode=0, stdout="LibreOffice test-version", stderr="", timed_out=False)
            return conversion(arguments)

        with patch.object(lo_render, "_execute", side_effect=execute):
            result = lo_render.render(self.source, self.output)
        self.assertEqual(self.original, self.source.read_bytes())
        self.assertTrue(result["source_preserved"])
        self.assertEqual(hashlib.sha256(self.original).hexdigest(), result["source_sha256"])
        self.assertEqual("LibreOffice test-version", result["renderer_version"])
        self.assertNotEqual(self.source, Path(calls[-1][-1]))
        json.dumps(result)
        return result

    @staticmethod
    def completed(code=0, stderr="", timed_out=False):
        return dict(returncode=code, stdout="", stderr=stderr, timed_out=timed_out)

    def test_failed_tracked_drawing_reports_renderer_error_and_retains_docx(self):
        result = self.run_render(lambda args: self.completed(1, "Unspecified Application Error"))
        self.assertEqual("failed", result["status"])
        self.assertEqual("renderer_failed", result["failure_code"])
        self.assertEqual("Unspecified Application Error", result["stderr"])
        self.assertEqual([dict(part="word/header1.xml", revision="ins", anchored=True,
                               grouped=True, textbox=True)], result["tracked_drawings"])

    def test_zero_exit_without_pdf_is_failure(self):
        result = self.run_render(lambda args: self.completed())
        self.assertEqual("missing_pdf", result["failure_code"])

    def test_stale_pdf_is_not_evidence_of_success(self):
        self.output.mkdir()
        old = self.output / "tracked.pdf"
        old.write_bytes(b"%PDF-old artifact")
        result = self.run_render(lambda args: self.completed())
        self.assertEqual("failed", result["status"])
        self.assertEqual(b"%PDF-old artifact", old.read_bytes())

    def test_fresh_pdf_is_published_after_success(self):
        def conversion(args):
            copy = Path(args[-1])
            self.assertEqual(self.original, copy.read_bytes())
            out = Path(args[args.index("--outdir") + 1])
            (out / "tracked.pdf").write_bytes(b"%PDF-1.7\n%%EOF\n")
            return self.completed()
        result = self.run_render(conversion)
        self.assertEqual("rendered", result["status"])
        self.assertIsNone(result["failure_code"])
        self.assertEqual(b"%PDF-1.7\n%%EOF\n", Path(result["pdf"]).read_bytes())

    def test_invalid_pdf_is_failure(self):
        def conversion(args):
            out = Path(args[args.index("--outdir") + 1])
            (out / "tracked.pdf").write_text("not a PDF")
            return self.completed()
        result = self.run_render(conversion)
        self.assertEqual("invalid_pdf", result["failure_code"])

    def test_pdf_destination_link_cannot_overwrite_source(self):
        self.output.mkdir()
        destination = self.output / "tracked.pdf"
        destination.symlink_to(self.source)
        def conversion(args):
            out = Path(args[args.index("--outdir") + 1])
            (out / "tracked.pdf").write_bytes(b"%PDF-1.7\n%%EOF\n")
            return self.completed()
        result = self.run_render(conversion)
        self.assertEqual("rendered", result["status"])
        self.assertFalse(destination.is_symlink())
        self.assertEqual(b"%PDF-1.7\n%%EOF\n", destination.read_bytes())

    def test_timeout_is_reported(self):
        result = self.run_render(lambda args: self.completed(-15, timed_out=True))
        self.assertEqual("renderer_timeout", result["failure_code"])

    def test_missing_renderer_is_reported(self):
        with patch.object(lo_render, "_execute", side_effect=FileNotFoundError("missing soffice")):
            result = lo_render.render(self.source, self.output)
        self.assertEqual("unavailable", result["status"])
        self.assertEqual("renderer_unavailable", result["failure_code"])
        self.assertEqual(self.original, self.source.read_bytes())

    def test_output_write_failure_is_not_reported_as_missing_renderer(self):
        self.output.write_text("a file blocks the output directory")
        def conversion(args):
            out = Path(args[args.index("--outdir") + 1])
            (out / "tracked.pdf").write_bytes(b"%PDF-1.7\n%%EOF\n")
            return self.completed()
        result = self.run_render(conversion)
        self.assertEqual("failed", result["status"])
        self.assertEqual("output_unwritable", result["failure_code"])

    def test_temporary_workspace_failure_is_reported_separately(self):
        with patch.object(lo_render, "_execute", return_value=self.completed()), \
                patch.object(lo_render.tempfile, "TemporaryDirectory", side_effect=OSError("temporary disk full")):
            result = lo_render.render(self.source, self.output)
        self.assertEqual("failed", result["status"])
        self.assertEqual("render_workspace_failed", result["failure_code"])
        self.assertTrue(result["source_preserved"])

    def test_timeout_terminates_renderer_children(self):
        started = time.monotonic()
        result = lo_render._execute([sys.executable, "-c", (
            "import subprocess, sys, time; "
            "p = subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(20)']); "
            "print(p.pid, flush=True); time.sleep(20)")], 1)
        self.assertTrue(result["timed_out"])
        self.assertLess(time.monotonic() - started, 5)
        child = int(result["stdout"].strip())
        state = Path("/proc") / str(child) / "stat"
        try:
            child_state = state.read_text().split()[2]
        except (FileNotFoundError, ProcessLookupError):
            return
        self.assertIn(child_state, ["Z", "X"])


if __name__ == "__main__":
    unittest.main()
