"""Unit tests for the 1.2.0 features: scanning, marker preparation, batched rendering, line spacing,
CJGE body formatting, CSS lengths, UTF-8 output and validation summaries. No Office required."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""
RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:line="312" w:lineRule="exact"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
<w:rPr><w:b/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="a3"><w:name w:val="List Bullet"/><w:basedOn w:val="Normal"/></w:style>
</w:styles>"""
EQUATION = (
    '<w:r><w:object w:dxaOrig="400" w:dyaOrig="560">'
    '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:20pt;height:28pt" o:ole="">'
    '<v:imagedata r:id="rId9" o:title=""/></v:shape>'
    '<o:OLEObject Type="Embed" ProgID="Equation.DSMT4" ShapeID="_x0000_i1025" DrawAspect="Content" '
    'ObjectID="_1" r:id="rId8"/></w:object></w:r>'
)


def make_docx(path: Path, body: str) -> Path:
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:v="urn:schemas-microsoft-com:vml" '
        f'xmlns:o="urn:schemas-microsoft-com:office:office" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<w:body>{body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:top="1247" w:right="1247" w:bottom="1247" w:left="1247" w:header="720" w:footer="720" w:gutter="0"/>'
        f'</w:sectPr></w:body></w:document>'
    )
    with zipfile.ZipFile(path, "w") as package:
        package.writestr("[Content_Types].xml", CONTENT_TYPES)
        package.writestr("_rels/.rels", RELS)
        package.writestr("word/document.xml", document)
        package.writestr("word/styles.xml", STYLES)
    return path


def para(text: str, style: str | None = None) -> str:
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f'<w:p>{ppr}<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def body_text(path: Path) -> str:
    with zipfile.ZipFile(path) as package:
        root = ET.fromstring(package.read("word/document.xml"))
    return "\n".join("".join(t.text or "" for t in p.iter(f"{{{W}}}t")) for p in root.iter(f"{{{W}}}p"))


class ScanTests(unittest.TestCase):
    def spans(self, text: str) -> list[str]:
        import mathtype_scan

        return [text[a:b] for a, b, _ in mathtype_scan.find_spans(text)]

    def tex(self, text: str) -> list[str]:
        import mathtype_scan

        return [mathtype_scan.to_tex(t) for _, _, t in mathtype_scan.find_spans(text)]

    def test_false_positives_seen_in_the_field(self) -> None:
        self.assertEqual(self.spans("(a) 碎石级配；(b) 格栅照片；(c) 格栅"), [])      # sub-figure labels
        self.assertEqual(self.spans("T1 碎石性质；T2 格栅；P1、P2"), [])              # table / paper labels
        self.assertEqual(self.spans("1.643/1.535 g/cm³ — clarify"), [])              # g in a unit
        self.assertEqual(self.spans("the columns and specimens"), [])                # s at word ends
        self.assertEqual(self.spans("strengths within [x]% and"), [])                # placeholder
        self.assertEqual(self.spans("多子图（a–d）的图"), [])                           # letter range
        self.assertEqual(self.spans("e.g. Table 3 and T30 ×3"), [])

    def test_units_and_values(self) -> None:
        self.assertEqual(self.spans("试样直径（R = 0.25 m）"), ["R = 0.25 m"])
        self.assertEqual(self.tex("R = 0.25 m"), [r"R=0.25\ \mathrm{m}"])
        self.assertEqual(self.tex("φ = 43.2°"), [r"\varphi =43.2{}^{\circ}"])
        self.assertEqual(self.tex("c_col(0.125 m) ≈ 120 kPa"),
                         [r"c_{\mathrm{col}}(0.125\ \mathrm{m})\approx 120\ \mathrm{kPa}"])

    def test_comma_subscripts_and_prefix_minus(self) -> None:
        self.assertEqual(self.spans("Δσ3 = σ1f/Kp,g − σ3 (Kp,g from"), ["Δσ3 = σ1f/Kp,g − σ3", "Kp,g"])
        self.assertEqual(self.tex("(c_col, φ_col, E50,col)"),
                         [r"(c_{\mathrm{col}},\varphi_{\mathrm{col}},E_{50,\mathrm{col}})"])
        self.assertEqual(self.tex("Δσ3,exp"), [r"\Delta \sigma_{3,\mathrm{exp}}"])

    def test_expressions(self) -> None:
        self.assertEqual(self.tex("c_app = Δσ3·√Kp/2"),
                         [r"c_{\mathrm{app}}=\Delta \sigma_{3}\cdot \sqrt{K_{\mathrm{p}}}/2"])
        self.assertEqual(self.tex("E50 = K_E·p_a·(σ3/p_a)^n"),
                         [r"E_{50}=K_{\mathrm{E}}\cdot p_{\mathrm{a}}\cdot (\sigma_{3}/p_{\mathrm{a}})^{n}"])
        self.assertEqual(self.tex("m = (d/d_e)²"), [r"m=(d/d_{\mathrm{e}})^{2}"])
        self.assertEqual(self.tex("tan φ_sp = m·μ_p"), [r"\tan \varphi_{\mathrm{sp}}=m\cdot \mu_{\mathrm{p}}"])
        self.assertEqual(self.spans("0.16×150 = 24 kPa，与"), ["0.16×150 = 24 kPa"])

    def test_ranges_and_prose_signs(self) -> None:
        self.assertEqual(self.tex("n ≈ 1.4–1.8"), [r"n\approx 1.4\sim 1.8"])
        self.assertEqual(self.tex("E50–σ3 power law"), [r"E_{50}-\sigma_{3}"])
        self.assertEqual(self.spans("E50 +25–44% at 50 kPa"), ["E50"])

    def test_greek_control_word_keeps_space(self) -> None:
        import mathtype_scan

        self.assertEqual(mathtype_scan.symbol_tex("SYMG", "β", ("i",)), r"\beta ")
        self.assertEqual(mathtype_scan.symbol_tex("SYMU", "E_i", ("i",)), "E_{i}")
        self.assertEqual(mathtype_scan.symbol_tex("SYMU", "E_i", ()), r"E_{\mathrm{i}}")


class PrepareTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="mt-prepare-"))

    def tearDown(self) -> None:
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_scan_prepare_in_runs_tables_and_cjge_italics(self) -> None:
        import mathtype_scan

        body = (
            para("取 E50 = K_E·p_a 与 φ 值")
            # the expression is split across two runs
            + '<w:p><w:r><w:t xml:space="preserve">其中 c_app = Δσ3</w:t></w:r>'
              '<w:r><w:t xml:space="preserve">·√Kp/2 成立</w:t></w:r></w:p>'
            + '<w:tbl><w:tr><w:tc>' + para("表中 T_ult 与 σ3") + '</w:tc></w:tr></w:tbl>'
            + para("参考文献", "1") + para("Smith J (2020) with c = 5 kPa.")
        )
        source = make_docx(self.tmp / "中文 源.docx", body)
        result = mathtype_scan.scan(str(source), "cjge")
        texts = [c["text"] for c in result["candidates"]]
        self.assertEqual(texts, ["E50 = K_E·p_a", "φ", "c_app = Δσ3·√Kp/2", "T_ult", "σ3"])  # no reference list
        actions = {c["text"]: c["action"] for c in result["candidates"]}
        self.assertEqual(actions["φ"], "italic_text")
        self.assertEqual(actions["E50 = K_E·p_a"], "mathtype")
        self.assertTrue(result["candidates"][3]["in_table"])
        candidates = self.tmp / "c.json"
        candidates.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
        out = self.tmp / "out.docx"
        manifest = self.tmp / "m.json"
        prepared = mathtype_scan.prepare(str(source), str(candidates), str(out), str(manifest))
        self.assertTrue(prepared["ok"], prepared)
        self.assertEqual(prepared["mathtype_markers"], 2)
        self.assertEqual(prepared["italic_text_symbols"], 3)
        text = body_text(out)
        self.assertIn("取 {{MATH:eq001}} 与 φ 值", text)
        self.assertIn("其中 {{MATH:eq002}} 成立", text)
        data = json.loads(manifest.read_text(encoding="utf-8"))
        self.assertEqual([e["id"] for e in data["equations"]], ["eq001", "eq002"])
        self.assertEqual(data["equations"][1]["tex"], r"c_{\mathrm{app}}=\Delta \sigma_{3}\cdot \sqrt{K_{\mathrm{p}}}/2")
        with zipfile.ZipFile(out) as package:
            xml = package.read("word/document.xml").decode("utf-8")
        self.assertIn('w:vertAlign w:val="subscript"', xml)      # T_ult -> T + subscript ult
        self.assertIn("<w:t xml:space=\"preserve\">ult</w:t>", xml)

    def test_prepare_refuses_changed_document(self) -> None:
        import mathtype_scan

        source = make_docx(self.tmp / "a.docx", para("E = J/t"))
        result = mathtype_scan.scan(str(source), "all")
        candidates = self.tmp / "c.json"
        candidates.write_text(json.dumps(result), encoding="utf-8")
        make_docx(source, para("E = J/t changed"))
        with self.assertRaises(ValueError):
            mathtype_scan.prepare(str(source), str(candidates), str(self.tmp / "o.docx"), str(self.tmp / "m.json"))


class BatchTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(tempfile.mkdtemp(prefix="mt-batch-"))
        os.environ["MATHTYPE_JOBS_DIR"] = str(self.tmp / "jobs")
        self.source = make_docx(self.tmp / "输入.docx", para("x"))
        self.manifest = self.tmp / "m.json"
        equations = [{"id": f"e{i}", "marker": "{{MATH:e%d}}" % i, "tex": "x", "layout": "inline", "numbered": False}
                     for i in range(1, 11)]
        equations.append({"id": "n1", "marker": "{{MATH:n1}}", "tex": "y", "layout": "display", "numbered": True})
        self.manifest.write_text(json.dumps({"schema_version": 1, "equations": equations,
                                             "references": [{"marker": "{{EQREF:r1}}", "target": "n1"}]}),
                                 encoding="utf-8")
        self.calls: list[list[str]] = []

    def tearDown(self) -> None:
        os.environ.pop("MATHTYPE_JOBS_DIR", None)
        shutil.rmtree(self.tmp, ignore_errors=True)

    def bridge(self, fail_first_chunk_of_size: int | None = None, crash_on_call: int | None = None):
        def run(action, arguments, timeout=None):
            part = json.loads(Path(arguments["manifest_path"]).read_text(encoding="utf-8"))
            ids = [e["id"] for e in part["equations"]]
            self.calls.append(ids)
            self.assertTrue(arguments["allow_unresolved_markers"])
            self.assertGreater(timeout, 0)
            if crash_on_call and len(self.calls) == crash_on_call:
                raise KeyboardInterrupt("simulated crash")
            if fail_first_chunk_of_size and len(ids) == fail_first_chunk_of_size and len(self.calls) == 1:
                return {"ok": False, "error": "RPC server unavailable"}
            shutil.copy2(arguments["input_path"], arguments["output_path"])
            return {"ok": True}
        return run

    @staticmethod
    def post(output, manifest, manifest_path, allow):
        return {"ok": True, "validation": {"ok": True}}

    def test_plan_puts_references_with_their_targets_last(self) -> None:
        import mathtype_batch

        manifest = mathtype_batch.load_manifest(str(self.manifest))
        chunks = mathtype_batch.plan_chunks(manifest, 4)
        self.assertEqual(chunks, [["e1", "e2", "e3", "e4"], ["e5", "e6", "e7", "e8"], ["e9", "e10"], ["n1"]])

    def test_failed_chunk_is_split_and_job_completes(self) -> None:
        import mathtype_batch

        out = self.tmp / "输出.docx"
        result = mathtype_batch.render(str(self.source), str(out), str(self.manifest), 4,
                                       bridge=self.bridge(fail_first_chunk_of_size=4), postprocess=self.post)
        self.assertTrue(result["ok"], result)
        self.assertEqual(self.calls[0], ["e1", "e2", "e3", "e4"])
        self.assertEqual(self.calls[1], ["e1", "e2"])
        self.assertEqual(result["retries"], 1)
        self.assertTrue(out.is_file())
        status = mathtype_batch.status(str(out), cleanup=True)
        self.assertEqual(status["status"], "completed")
        self.assertFalse(mathtype_batch.job_dir(str(out)).exists())

    def test_interrupted_job_resumes_from_checkpoint(self) -> None:
        import mathtype_batch

        out = self.tmp / "out.docx"
        with self.assertRaises(KeyboardInterrupt):
            mathtype_batch.render(str(self.source), str(out), str(self.manifest), 4,
                                  bridge=self.bridge(crash_on_call=3), postprocess=self.post)
        self.assertEqual(mathtype_batch.status(str(out))["status"], "interrupted")
        self.calls.clear()
        result = mathtype_batch.render(str(self.source), str(out), str(self.manifest), 4,
                                       bridge=self.bridge(), postprocess=self.post)
        self.assertTrue(result["ok"], result)
        self.assertEqual(self.calls, [["e9", "e10"], ["n1"]])   # chunks 1 and 2 are not repeated

    def test_existing_output_requires_overwrite(self) -> None:
        import mathtype_batch

        out = self.tmp / "exists.docx"
        out.write_bytes(b"x")
        with self.assertRaises(FileExistsError):
            mathtype_batch.render(str(self.source), str(out), str(self.manifest), 4,
                                  bridge=self.bridge(), postprocess=self.post)


class LineSpacingTests(unittest.TestCase):
    def test_inherited_exact_spacing_is_fixed_only_where_clipped(self) -> None:
        import docx_postprocess

        tmp = Path(tempfile.mkdtemp(prefix="mt-spacing-"))
        try:
            body = f"<w:p><w:r><w:t>a </w:t></w:r>{EQUATION}</w:p>" + para("plain")
            source = make_docx(tmp / "s.docx", body)
            check = docx_postprocess.check(str(source))
            self.assertEqual(check["clipped_paragraphs"], 1)
            self.assertTrue(check["warnings"])
            fixed = docx_postprocess.fix(str(source), str(tmp / "f.docx"))
            self.assertEqual(fixed["paragraphs_fixed"], 1)
            self.assertEqual(docx_postprocess.check(str(tmp / "f.docx"))["clipped_paragraphs"], 0)
            with zipfile.ZipFile(tmp / "f.docx") as package:
                xml = package.read("word/document.xml").decode("utf-8")
            self.assertIn('w:line="312" w:lineRule="atLeast"', xml)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class BodyFormatTests(unittest.TestCase):
    ORDER_TABLES = ("RPR_ORDER", "PPR_ORDER", "TBLPR_ORDER", "TCPR_ORDER")

    def test_apply_then_report_conforms_and_keeps_schema_order(self) -> None:
        import cjge_body_format as body_format

        tmp = Path(tempfile.mkdtemp(prefix="mt-body-"))
        try:
            body = (para("论文题目") + para("1  引言", "1") + f"<w:p><w:r><w:t>正文 </w:t></w:r>{EQUATION}</w:p>"
                    + para("列表项", "a3")
                    + "<w:tbl><w:tr><w:tc>" + para("表头") + "</w:tc></w:tr><w:tr><w:tc>" + para("数据")
                    + "</w:tc></w:tr></w:tbl>")
            source = make_docx(tmp / "b.docx", body)
            before = body_format.report(str(source), "cjge")
            self.assertFalse(before["conforms"])
            result = body_format.apply(str(source), str(tmp / "o.docx"))
            self.assertEqual(result["paragraphs"]["title"], 1)
            self.assertEqual(result["paragraphs"]["list"], 1)
            self.assertEqual(result["equation_runs_untouched"], 1)
            after = body_format.report(str(tmp / "o.docx"), "cjge")
            self.assertTrue(after["conforms"], after.get("deviations"))
            with zipfile.ZipFile(tmp / "o.docx") as package:
                root = ET.fromstring(package.read("word/document.xml"))
            orders = {"rPr": body_format.RPR_ORDER, "pPr": body_format.PPR_ORDER,
                      "tblPr": body_format.TBLPR_ORDER, "tcPr": body_format.TCPR_ORDER}
            for tag, order in orders.items():
                for element in root.iter(f"{{{W}}}{tag}"):
                    names = [child.tag.split("}")[-1] for child in element]
                    ranks = [order.index(n) for n in names if n in order]
                    self.assertEqual(ranks, sorted(ranks), f"{tag} children out of order: {names}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class CssAndSummaryTests(unittest.TestCase):
    def test_css_lengths(self) -> None:
        import mathtype_prefs

        self.assertEqual(mathtype_prefs._css_length_pt("width:1in;height:18.9pt", "width"), 72.0)
        self.assertAlmostEqual(mathtype_prefs._css_length_pt("height:1cm", "height"), 28.3465, places=3)
        self.assertIsNone(mathtype_prefs._css_length_pt("margin-left:3pt", "left"))
        self.assertEqual(mathtype_prefs._set_css_length_pt("width:1in;height:2pt", "width", 40.5),
                         "width:40.5pt;height:2pt")

    def test_summary_groups_repeated_errors(self) -> None:
        import mathtype_prefs

        lines = mathtype_prefs._summarize([f"obj{i}" for i in range(12)], "equation(s) are off", "fix it")
        self.assertEqual(len(lines), 1)
        self.assertIn("12 equation(s) are off", lines[0])
        self.assertIn("and 7 more", lines[0])


class Utf8Tests(unittest.TestCase):
    def test_cli_and_server_emit_utf8_with_cjk_paths(self) -> None:
        env = {k: v for k, v in os.environ.items() if k not in ("PYTHONIOENCODING", "PYTHONUTF8")}
        missing = str(Path(tempfile.gettempdir()) / "不存在的目录" / "文档.docx")
        completed = subprocess.run([sys.executable, str(SCRIPTS / "docx_postprocess.py"), "check", missing],
                                   capture_output=True, env=env, check=False)
        result = json.loads(completed.stdout.decode("utf-8"))
        self.assertFalse(result["ok"])
        self.assertIn("不存在的目录", result["error"])
        server = subprocess.Popen([sys.executable, str(SCRIPTS / "mcp_server.py")], stdin=subprocess.PIPE,
                                  stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
        try:
            request = {"jsonrpc": "2.0", "id": 1, "method": "tools/call",
                       "params": {"name": "get_mathtype_render_status", "arguments": {"output_path": missing}}}
            server.stdin.write((json.dumps(request, ensure_ascii=False) + "\n").encode("utf-8"))
            server.stdin.flush()
            response = json.loads(server.stdout.readline().decode("utf-8"))
            self.assertIn("文档.docx", response["result"]["structuredContent"]["output_path"])
        finally:
            server.kill()
            server.communicate()


if __name__ == "__main__":
    unittest.main()
