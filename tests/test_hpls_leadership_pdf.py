from pathlib import Path
import sys
import tempfile
import unittest

import pdfplumber
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import build_hpls_leadership_pdf as report


def image_xobject_count(page):
    resources = page.get('/Resources') or {}
    xobjects = resources.get('/XObject') or {}
    return sum(
        1
        for reference in xobjects.values()
        if reference.get_object().get('/Subtype') == '/Image'
    )


class LeadershipPdfContractTests(unittest.TestCase):
    def test_precheck_evidence_page_contains_image_and_management_conclusions(self):
        output = ROOT / "tmp" / "pdfs" / "hpls_leadership" / "precheck-evidence.pdf"
        report.register_fonts()
        document = canvas.Canvas(str(output), pagesize=landscape(A4), pageCompression=1)
        report.draw_evidence_precheck(document)
        document.showPage()
        document.save()

        reader = PdfReader(str(output))
        page = reader.pages[0]
        text = page.extract_text() or ""
        for required in ["输入预检与失败保护", "缺文件即停止", "问题可定位", "修复可执行"]:
            self.assertIn(required, text)
        self.assertGreaterEqual(image_xobject_count(page), 1)

    def test_evidence_assets_are_stable_and_missing_assets_fail_clearly(self):
        self.assertEqual(
            {key: path.name for key, path in report.EVIDENCE_ASSETS.items()},
            {
                "precheck": "pe100-input-check.png",
                "success": "xp-generation-success.png",
                "review": "manual-review-warning.png",
            },
        )
        self.assertIs(report.validate_evidence_assets(), report.EVIDENCE_ASSETS)
        self.assertTrue(all(path.is_file() for path in report.EVIDENCE_ASSETS.values()))

        with tempfile.TemporaryDirectory() as temp_dir:
            missing = Path(temp_dir) / "missing-evidence.png"
            with self.assertRaisesRegex(FileNotFoundError, "missing-evidence\\.png"):
                report.validate_evidence_assets({"missing": missing})
            with self.assertRaisesRegex(FileNotFoundError, "missing-string-evidence\\.png"):
                report.validate_evidence_assets(
                    {"missing": str(Path(temp_dir) / "missing-string-evidence.png")}
                )

    def test_measured_metrics_use_interval_midpoints(self):
        metrics = report.compute_metrics()
        self.assertEqual(metrics["manual_mid_minutes"], 90.0)
        self.assertEqual(metrics["auto_mid_minutes"], 1.5)
        self.assertEqual(metrics["minutes_saved"], 88.5)
        self.assertAlmostEqual(metrics["time_reduction_percent"], 98.33, places=2)

    def test_page_contract_matches_approved_seven_sections(self):
        self.assertEqual(
            report.PAGE_TITLES,
            ["封面", "成果摘要", "业务痛点", "方案设计", "人机分工", "量化成效", "复用方案"],
        )

    def test_solution_page_has_four_protection_mechanisms(self):
        self.assertEqual(
            report.GUARDRAILS,
            ["输入契约", "失败即中止", "输出规则校验", "源表与原模板不改动"],
        )

    def test_solution_step_copy_uses_readable_body_type(self):
        self.assertGreaterEqual(report.SOLUTION_BODY_FONT_SIZE, 10)
        self.assertEqual(
            report.SOLUTION_STAGES[2][2],
            "缺失 / 重复 / 模板 / 依赖 / 路径",
        )

    def test_pages_three_to_seven_use_readable_small_type(self):
        self.assertGreaterEqual(report.CONTENT_BODY_FONT_SIZE, 10)
        self.assertGreaterEqual(report.CONTENT_COMPACT_FONT_SIZE, 9.5)

    def test_enlarged_callouts_use_copy_that_fits_without_dangling_punctuation(self):
        self.assertEqual(
            report.HUMAN_MACHINE_PRINCIPLE,
            "机器执行确定性规则；业务判断与质量责任保留给人。",
        )
        self.assertEqual(
            report.ORGANIZATION_VALUE,
            "专业人员转向异常判断、质量复核与规则优化。",
        )
        self.assertEqual(
            report.PLATFORM_ADAPTATION,
            "按平台适配源表字段、匹配规则、分组逻辑、计算公式、模板与业务阈值。",
        )

    def test_impact_row_label_and_duration_share_a_baseline(self):
        output = ROOT / "tmp" / "pdfs" / "hpls_leadership" / "impact-alignment.pdf"
        report.build_pdf(output)

        with pdfplumber.open(output) as pdf:
            words = pdf.pages[5].extract_words(keep_blank_chars=True, extra_attrs=["size"])
        label = next(word for word in words if abs(word["x0"] - 55.0) < 0.2 and 300 < word["top"] < 340)
        duration = next(word for word in words if abs(word["x0"] - 154.0) < 0.2 and 300 < word["top"] < 340)

        self.assertAlmostEqual(label["bottom"], duration["bottom"], delta=0.5)

    def test_generated_pdf_has_seven_pages_and_required_content(self):
        output = ROOT / "tmp" / "pdfs" / "hpls_leadership" / "contract.pdf"
        report.build_pdf(output)

        reader = PdfReader(str(output))
        self.assertEqual(len(reader.pages), 7)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        for required in [
            "业务痛点",
            "方案设计",
            "人机分工",
            "量化成效",
            "复用方案",
            "约 98%",
            "88.5 分钟",
        ]:
            self.assertIn(required, text)


if __name__ == "__main__":
    unittest.main()
