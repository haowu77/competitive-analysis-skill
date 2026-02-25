from __future__ import annotations

import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path


class CompetitiveAnalysisTest(unittest.TestCase):
    def setUp(self) -> None:
        self.root = Path(__file__).resolve().parents[1]
        self.script = self.root / "scripts" / "build_benchmark_xlsx.py"

    def _read_zip_text(self, xlsx: Path, member: str) -> str:
        with zipfile.ZipFile(xlsx, "r") as zf:
            return zf.read(member).decode("utf-8")

    def test_auto_lang_zh_sheet_names(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "benchmark.xlsx"
            subprocess.run(
                [
                    "python3",
                    str(self.script),
                    "--output",
                    str(out),
                    "--brief",
                    "这是一个用于家庭做饭规划的AI产品，包含备菜和库存管理。",
                    "--lang",
                    "auto",
                    "--lang-source",
                    "brief",
                ],
                check=True,
            )

            workbook_xml = self._read_zip_text(out, "xl/workbook.xml")
            self.assertIn('sheet name="摘要"', workbook_xml)
            self.assertIn('sheet name="竞品基准"', workbook_xml)
            self.assertIn('sheet name="功能矩阵"', workbook_xml)
            self.assertIn('sheet name="定价-GTM"', workbook_xml)
            self.assertIn('sheet name="证据来源"', workbook_xml)
            self.assertEqual(workbook_xml.count("<sheet name="), 5)

    def test_force_en_sheet_names(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "benchmark.xlsx"
            subprocess.run(
                [
                    "python3",
                    str(self.script),
                    "--output",
                    str(out),
                    "--brief",
                    "这是一个用于家庭做饭规划的AI产品，包含备菜和库存管理。",
                    "--lang",
                    "en",
                ],
                check=True,
            )

            workbook_xml = self._read_zip_text(out, "xl/workbook.xml")
            self.assertIn('sheet name="Summary"', workbook_xml)
            self.assertIn('sheet name="Benchmark"', workbook_xml)
            self.assertIn('sheet name="Feature-Matrix"', workbook_xml)
            self.assertIn('sheet name="Pricing-GTM"', workbook_xml)
            self.assertIn('sheet name="Sources"', workbook_xml)


if __name__ == "__main__":
    unittest.main()
