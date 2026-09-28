"""Live Word/MathType regression tests for the 1.2.0 fixes. Run with MATHTYPE_OFFICE_LIVE_TEST=1
(tests/run-tests.ps1 -IncludeLiveOffice). Each test uses a normal folder under the user's Documents,
because Word opens files under %TEMP% in Protected View."""
from __future__ import annotations

import json
import os
import shutil
import sys
import unittest
import uuid
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

LIVE = os.environ.get("MATHTYPE_OFFICE_LIVE_TEST") == "1" or os.environ.get("MATHTYPE_WORD_LIVE_TEST") == "1"


@unittest.skipUnless(LIVE, "live Office test not requested")
class LiveRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        base = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents"
        cls.work = base / f"mathtype-for-word-live-测试-{uuid.uuid4().hex[:8]}"
        cls.work.mkdir(parents=True)
        os.environ["MATHTYPE_JOBS_DIR"] = str(cls.work / "jobs")

    @classmethod
    def tearDownClass(cls) -> None:
        os.environ.pop("MATHTYPE_JOBS_DIR", None)
        if os.environ.get("MATHTYPE_KEEP_LIVE_ARTIFACTS") != "1":
            shutil.rmtree(cls.work, ignore_errors=True)

    def build(self, name: str, body: str) -> Path:
        from test_v12_features import make_docx

        return make_docx(self.work / name, body)

    def test_table_markers_chinese_path_and_batches(self) -> None:
        import mcp_server
        from test_v12_features import para

        cells = "".join(f"<w:tc>{para('表中 {{MATH:t%d}} 值' % i)}</w:tc>" for i in range(1, 4))
        merged = ('<w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' + para("合并 {{MATH:t4}}") + "</w:tc>"
                  + "<w:tc>" + para("x") + "</w:tc>")
        # exact 15.6 pt on the paragraph itself: Word replaces the minimal test styles.xml when it saves
        exact = '<w:p><w:pPr><w:spacing w:line="312" w:lineRule="exact"/></w:pPr><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p>'
        body = "".join(exact.format("正文 {{MATH:e%d}} 结束。" % i) for i in range(1, 121))
        body += f"<w:tbl><w:tr>{cells}</w:tr><w:tr>{merged}</w:tr></w:tbl>"
        source = self.build("输入 文档.docx", body)
        equations = [{"id": f"e{i}", "marker": "{{MATH:e%d}}" % i, "tex": r"\frac{a_{%d}}{b}" % i,
                      "layout": "inline", "numbered": False} for i in range(1, 121)]
        equations += [{"id": f"t{i}", "marker": "{{MATH:t%d}}" % i, "tex": r"\sigma_{%d}" % i,
                       "layout": "inline", "numbered": False} for i in range(1, 5)]
        manifest = self.work / "清单.json"
        manifest.write_text(json.dumps({"schema_version": 1, "equations": equations}, ensure_ascii=False),
                            encoding="utf-8")
        output = self.work / "输出-mathtype.docx"
        result = mcp_server._render_word({"input_path": str(source), "output_path": str(output),
                                          "manifest_path": str(manifest), "batch_size": 50})
        self.assertTrue(result.get("ok"), json.dumps(result, ensure_ascii=False)[:3000])
        self.assertEqual(result["chunks"], 3)
        self.assertEqual(result["validation"]["counts"]["mathtype_objects"], 124)
        self.assertTrue(result["validation"]["equation_format"]["ok"])
        steps = result["post_processing"]
        self.assertGreater(steps["inline_line_spacing"]["paragraphs_fixed"], 0)   # \frac clips 15.6 pt lines

    def test_partial_render_and_format_check_with_inch_sizes(self) -> None:
        import mcp_server
        from test_v12_features import para

        source = self.build("partial.docx", para("A {{MATH:a}} B {{MATH:b}}"))
        manifest = self.work / "partial.json"
        manifest.write_text(json.dumps({"schema_version": 1, "equations": [
            {"id": "a", "marker": "{{MATH:a}}", "tex": "x^{2}", "layout": "inline", "numbered": False}]}),
            encoding="utf-8")
        output = self.work / "partial-out.docx"
        strict = mcp_server._render_word({"input_path": str(source), "output_path": str(output),
                                          "manifest_path": str(manifest), "resume": False})
        self.assertFalse(strict.get("ok"))
        partial = mcp_server._render_word({"input_path": str(source), "output_path": str(output),
                                           "manifest_path": str(manifest), "resume": False, "overwrite": True,
                                           "allow_unresolved_markers": True})
        self.assertTrue(partial.get("ok"), json.dumps(partial, ensure_ascii=False)[:2000])
        self.assertTrue(any("partial render" in w for w in partial["validation"]["warnings"]))
        # Word writes exactly 72 pt as "1in"; the format check must read it instead of crashing.
        inch = self.work / "inch.docx"
        with zipfile.ZipFile(output) as zin, zipfile.ZipFile(inch, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "word/document.xml":
                    text = data.decode("utf-8")
                    start = text.index('style="width:') + len('style="width:')
                    end = text.index(";", start)
                    data = (text[:start] + "1in" + text[end:]).encode("utf-8")
                zout.writestr(item, data)
        import mathtype_prefs

        check = mathtype_prefs.check_docx(str(inch))
        self.assertEqual(check["equations_checked"], 1)
        self.assertFalse(check["ok"])                      # 72 pt is the wrong size, reported as one line
        self.assertEqual(len(check["errors"]), 1)
        fixed = mathtype_prefs.apply_to_docx(str(inch), str(self.work / "inch-fixed.docx"))
        self.assertTrue(fixed["ok"])
        self.assertTrue(mathtype_prefs.check_docx(str(self.work / "inch-fixed.docx"))["ok"])


if __name__ == "__main__":
    unittest.main()
