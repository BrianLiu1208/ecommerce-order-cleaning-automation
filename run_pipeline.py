"""Run the complete demo: clean orders + refresh Excel + customer showcase."""
import argparse
from pathlib import Path
from clean_orders import ROOT, process
from xlsx_report import build_workbook
from customer_report import render


def main():
    parser = argparse.ArgumentParser(description='Clean orders and refresh Excel in a single command.')
    parser.add_argument('--input-dir', type=Path, default=ROOT / 'data/raw')
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'data/output')
    parser.add_argument('--mapping', type=Path, default=ROOT / 'rules/product_mapping.csv')
    parser.add_argument('--report', type=Path, default=ROOT / 'reports/ecommerce_cleaning_demo.xlsx')
    parser.add_argument('--customer-report', type=Path, default=ROOT / 'reports/customer_case_study.html')
    args = parser.parse_args()
    metrics = process(args.input_dir, args.output_dir, args.mapping)
    report = build_workbook(args.output_dir, args.report)
    showcase = render(args.input_dir, args.output_dir, args.customer_report)
    print('Pipeline finished successfully:')
    for key in ['raw_rows','accepted_rows','rejected_rows','exact_duplicate_rows',
                'conflicting_duplicate_rows','completed_sales_twd']:
        print(f'  {key}: {metrics[key]}')
    print('Excel:', report)
    print('Customer showcase:', showcase)


if __name__ == '__main__':
    main()
