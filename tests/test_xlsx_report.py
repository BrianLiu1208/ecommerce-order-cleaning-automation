"""Test the XLSX auto-refresh output without third-party packages."""
import csv
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile
from clean_orders import process
from generate_sample_data import generate
from xlsx_report import build_workbook

ROOT = Path(__file__).resolve().parents[1]
NS = {'m': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def cached_value(report, sheet_index, address):
    with ZipFile(report) as archive:
        root = ET.fromstring(archive.read(f'xl/worksheets/sheet{sheet_index}.xml'))
    node = root.find(f".//m:c[@r='{address}']/m:v", NS)
    return float(node.text) if node is not None else None


class TestExcelReport(unittest.TestCase):
    def test_report_refreshes_after_new_order(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary)
            src, output = path / 'raw', path / 'output'
            report = path / 'report.xlsx'
            generate(src)
            process(src, output, ROOT / 'rules/product_mapping.csv')
            build_workbook(output, report)
            self.assertEqual(cached_value(report, 1, 'A6'), 77)
            self.assertEqual(cached_value(report, 1, 'E6'), 117585)
            with ZipFile(report) as archive:
                self.assertIsNone(archive.testzip())
                self.assertEqual(sum(x.startswith('xl/worksheets/sheet') for x in archive.namelist()), 7)
            with (src / 'orders_shop.csv').open('a', newline='', encoding='utf-8-sig') as file:
                csv.writer(file).writerow(['WEB-9999', '2026-10-25', '無線滑鼠', '2', '500', 'C999', '完成'])
            metrics = process(src, output, ROOT / 'rules/product_mapping.csv')
            self.assertEqual(metrics['accepted_rows'], 78)
            self.assertEqual(metrics['completed_sales_twd'], '118585.00')
            build_workbook(output, report)
            self.assertEqual(cached_value(report, 1, 'A6'), 78)
            self.assertEqual(cached_value(report, 1, 'E6'), 118585)
            with ZipFile(report) as archive:
                self.assertIsNone(archive.testzip())

if __name__ == '__main__':
    unittest.main()
