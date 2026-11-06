"""Run from project root: python3 -m unittest discover -s tests -v"""
import csv
import tempfile
import unittest
from pathlib import Path

from clean_orders import process, parse_date, parse_amount, parse_quantity
from generate_sample_data import generate

ROOT = Path(__file__).resolve().parents[1]


class TestOrderPipeline(unittest.TestCase):
    def test_date_formats(self):
        self.assertEqual(parse_date('10/09/2026'), '2026-10-09')
        self.assertEqual(parse_date('2026年10月09日'), '2026-10-09')
        with self.assertRaises(ValueError):
            parse_date('10/32/2026')

    def test_prices_and_quantities(self):
        self.assertEqual(str(parse_amount('NT$1,299')), '1299')
        self.assertEqual(parse_quantity('2'), 2)
        with self.assertRaises(ValueError):
            parse_quantity('-2')
        with self.assertRaises(ValueError):
            parse_amount('unknown')

    def test_pipeline_reconciliation_and_conflict_quarantine(self):
        with tempfile.TemporaryDirectory() as tmp:
            src, out = Path(tmp) / 'input', Path(tmp) / 'output'
            generate(src)
            metrics = process(src, out, ROOT / 'rules/product_mapping.csv')
            self.assertEqual(metrics['raw_rows'], 87)
            self.assertEqual(metrics['exact_duplicate_rows'], 2)
            self.assertEqual(metrics['conflicting_duplicate_rows'], 2)
            self.assertEqual(metrics['rejected_rows'], 6)
            self.assertEqual(metrics['raw_rows'], metrics['accepted_rows'] + metrics['exact_duplicate_rows'] + metrics['conflicting_duplicate_rows'] + metrics['rejected_rows'])
            with (out / 'clean_orders.csv').open(encoding='utf-8-sig', newline='') as f:
                cleaned = list(csv.DictReader(f))
            self.assertTrue(any(r['quality_flags'] for r in cleaned))
            self.assertFalse(any(r['order_id'] == 'POS-1003' for r in cleaned))
            self.assertTrue(all(float(r['line_amount_twd']) == float(r['unit_price_twd']) * int(r['quantity']) for r in cleaned))


if __name__ == '__main__':
    unittest.main()
