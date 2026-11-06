"""Verify customer-facing HTML content reflects actual cleaned outputs."""
import csv
import tempfile
import unittest
from pathlib import Path

from clean_orders import process
from customer_report import render
from generate_sample_data import generate

ROOT = Path(__file__).resolve().parents[1]

class TestCustomerReport(unittest.TestCase):
    def test_customer_showcase_refreshes_with_new_order(self):
        with tempfile.TemporaryDirectory() as tmp:
            temp = Path(tmp)
            input_dir = temp / 'raw'
            output_dir = temp / 'output'
            html_path = temp / 'customer_case_study.html'
            generate(input_dir)
            process(input_dir, output_dir, ROOT / 'rules/product_mapping.csv')
            render(input_dir, output_dir, html_path)
            page = html_path.read_text(encoding='utf-8')
            self.assertIn('117,585', page)
            self.assertIn('POS-1003', page)
            self.assertIn('Wireless Mouse', page)
            self.assertIn('MKT-1005', page)
            self.assertIn('87 筆資料最後去了哪裡', page)
            self.assertIn('清洗前後對照', page)
            self.assertNotIn('https://', page)  # fully offline, no third-party resources
            with (input_dir / 'orders_shop.csv').open('a', newline='', encoding='utf-8-sig') as f:
                csv.writer(f).writerow(['WEB-9999', '2026-10-25', '無線滑鼠', '2', '500', 'C999', '完成'])
            process(input_dir, output_dir, ROOT / 'rules/product_mapping.csv')
            render(input_dir, output_dir, html_path)
            page = html_path.read_text(encoding='utf-8')
            self.assertIn('118,585', page)
            self.assertIn('88 筆資料最後去了哪裡', page)
            self.assertNotIn('117,585', page)
            self.assertIn('78', page)

if __name__ == '__main__':
    unittest.main()
