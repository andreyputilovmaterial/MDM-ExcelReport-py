"""Validate the committed example inputs of ``mdmexcelreportpy``.

Each ``tests/fixtures/*.mdd.json`` is an example *input* (from ``mdmreadpy`` or the
diff tool). This package renders them to Excel; these tests re-prepare the DataFrames
and write a workbook to a temp file (no Windows/COM required).
"""

import json
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'src'))

import report_create  # noqa: E402

FIXTURES = ROOT / 'tests' / 'fixtures'

INPUTS = [
    'a1_v1.mdd.json',
    'a1_v2.mdd.json',
    'b1_today.mdd.json',
    'b1_tomorrow.mdd.json',
    'report.diff.a1_-v1-v2-.mdd.json',
    'report.diff.b1_-today-tomorrow-.mdd.json',
    'report.diff.report.diff.-a1_-b1_---v1-today---v2-tomorrow--.mdd.json',
]


class TestReportExampleInputs(unittest.TestCase):
    def test_every_input_prepares_dataframes(self):
        for name in INPUTS:
            with self.subTest(name=name):
                report = json.loads((FIXTURES / name).read_text(encoding='utf-8'))
                doc = report_create.ReportDocument(report, {})
                self.assertTrue(doc.dataframes)
                for df in doc.dataframes:
                    self.assertIn('name', df)
                    self.assertIn('df', df)
                    self.assertIn('columns', df)

    def test_every_input_writes_a_workbook(self):
        for name in INPUTS:
            with self.subTest(name=name):
                report = json.loads((FIXTURES / name).read_text(encoding='utf-8'))
                doc = report_create.ReportDocument(report, {})
                with tempfile.TemporaryDirectory() as tmpdir:
                    out = Path(tmpdir) / 'out.xlsx'
                    doc.write_to_file(str(out))
                    self.assertTrue(out.is_file(), out)
                    self.assertGreater(out.stat().st_size, 0)


if __name__ == '__main__':
    unittest.main()
