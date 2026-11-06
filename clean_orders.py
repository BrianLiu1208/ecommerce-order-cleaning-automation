#!/usr/bin/env python3
"""Clean three demo order exports with Python's standard library.

Safety principles: invalid rows are quarantined, conflicting duplicates are not
silently overwritten, and every input row gets an audit trail.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
from collections import Counter, defaultdict
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SOURCES = {
    'orders_shop.csv': {'channel': '官網', 'columns': ('訂單編號', '下單日期', '商品', '數量', '單價', '客戶ID', '狀態')},
    'orders_marketplace.csv': {'channel': '電商平台', 'columns': ('Order_ID', 'Purchased_At', 'Item', 'Quantity', 'Price', 'Buyer_ID', 'Order_Status')},
    'orders_pos.csv': {'channel': '實體門市', 'columns': ('交易編號', '交易時間', '品項', '件數', '售價', '會員編號', '交易狀態')},
}
STATUS_MAP = {
    '完成': 'Completed', '已完成': 'Completed', 'completed': 'Completed', 'paid': 'Completed', 'fulfilled': 'Completed',
    '取消': 'Cancelled', '已取消': 'Cancelled', 'cancelled': 'Cancelled', 'canceled': 'Cancelled',
    '退款': 'Refunded', '已退款': 'Refunded', 'refunded': 'Refunded',
}
DATE_FORMATS = ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%m/%d/%Y', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M', '%Y年%m月%d日')
CLEAN_FIELDS = ['channel', 'order_id', 'order_date', 'product', 'original_product', 'quantity', 'unit_price_twd', 'line_amount_twd', 'customer_id', 'status', 'quality_flags', 'source_file', 'source_row']
AUDIT_FIELDS = ['source_file', 'source_row', 'channel', 'order_id', 'result', 'reason']
REVIEW_FIELDS = ['source_file', 'source_row', 'channel', 'order_id', 'reason', 'raw_record_json']


def read_csv(path: Path):
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, fieldnames: list[str], rows: list[dict]):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)


def normalize_text(value: str) -> str:
    return re.sub(r'\s+', ' ', (value or '').strip()).casefold()


def load_product_mapping(path: Path) -> dict[str, str]:
    rows = read_csv(path)
    mapping = {}
    for row in rows:
        alias = normalize_text(row['alias'])
        if alias:
            mapping[alias] = row['canonical'].strip()
    return mapping


def parse_date(value: str) -> str:
    value = (value or '').strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue
    raise ValueError('invalid_date')


def parse_amount(value: str) -> Decimal:
    cleaned = re.sub(r'(?i)(?:NT\$|TWD|NTD|\$|,|\s)', '', (value or '').strip())
    try:
        amount = Decimal(cleaned)
    except (InvalidOperation, ValueError) as e:
        raise ValueError('invalid_unit_price') from e
    if not amount.is_finite() or amount <= 0 or amount.as_tuple().exponent < -2:
        raise ValueError('invalid_unit_price')
    return amount


def parse_quantity(value: str) -> int:
    try:
        raw = Decimal((value or '').strip())
    except InvalidOperation as e:
        raise ValueError('invalid_quantity') from e
    if not raw.is_finite() or raw <= 0 or raw != raw.to_integral_value():
        raise ValueError('invalid_quantity')
    return int(raw)


def normalize_row(raw: dict, filename: str, row_number: int, config: dict, product_map: dict):
    keys = config['columns']
    order_id, raw_date, raw_product, raw_quantity, raw_price, customer_id, raw_status = [str(raw.get(k) or '').strip() for k in keys]
    if not order_id:
        raise ValueError('missing_order_id')
    if not raw_product:
        raise ValueError('missing_product')
    date = parse_date(raw_date)
    quantity = parse_quantity(raw_quantity)
    price = parse_amount(raw_price)
    status = STATUS_MAP.get(normalize_text(raw_status))
    if not status:
        raise ValueError('unknown_status')
    canonical = product_map.get(normalize_text(raw_product), re.sub(r'\s+', ' ', raw_product).strip())
    flags = []
    if not customer_id:
        flags.append('missing_optional_customer_id')
    if normalize_text(raw_product) not in product_map:
        flags.append('unmapped_product')
    return {
        'channel': config['channel'], 'order_id': order_id, 'order_date': date,
        'product': canonical, 'original_product': raw_product, 'quantity': quantity,
        'unit_price_twd': f'{price:.2f}', 'line_amount_twd': f'{price * quantity:.2f}',
        'customer_id': customer_id, 'status': status, 'quality_flags': ';'.join(flags),
        'source_file': filename, 'source_row': row_number,
    }


def fingerprint(row: dict) -> tuple:
    return tuple(row[k] for k in ('channel', 'order_id', 'order_date', 'product', 'quantity', 'unit_price_twd', 'customer_id', 'status'))


def process(input_dir: Path, output_dir: Path, mapping_file: Path) -> dict:
    product_map = load_product_mapping(mapping_file)
    parsed = []
    rejected = []
    audit = []
    total_rows = 0

    for filename, config in SOURCES.items():
        file_path = input_dir / filename
        if not file_path.exists():
            raise FileNotFoundError(f'Missing required source: {file_path}')
        with file_path.open('r', encoding='utf-8-sig', newline='') as f:
            reader = csv.DictReader(f)
            missing_headers = set(config['columns']) - set(reader.fieldnames or ())
            if missing_headers:
                raise ValueError(f'{filename}: missing column(s): {sorted(missing_headers)}')
            for row_no, raw in enumerate(reader, 2):
                total_rows += 1
                basic = {
                    'source_file': filename, 'source_row': row_no,
                    'channel': config['channel'], 'order_id': str(raw.get(config['columns'][0]) or '').strip(),
                }
                try:
                    clean = normalize_row(raw, filename, row_no, config, product_map)
                    parsed.append((clean, raw, basic))
                except ValueError as error:
                    reason = str(error)
                    rejected.append({**basic, 'reason': reason, 'raw_record_json': json.dumps(raw, ensure_ascii=False)})
                    audit.append({**basic, 'result': 'rejected', 'reason': reason})

    groups = defaultdict(list)
    for record in parsed:
        clean = record[0]
        groups[(clean['channel'], clean['order_id'])].append(record)

    accepted = []
    duplicate_review = []
    for records in groups.values():
        if len(set(fingerprint(record[0]) for record in records)) > 1:
            for clean, raw, basic in records:
                duplicate_review.append({**basic, 'reason': 'conflicting_duplicate', 'raw_record_json': json.dumps(raw, ensure_ascii=False)})
                audit.append({**basic, 'result': 'conflicting_duplicate', 'reason': 'same_channel_order_id_different_values'})
        else:
            for ix, (clean, raw, basic) in enumerate(records):
                if ix == 0:
                    accepted.append(clean)
                    audit.append({**basic, 'result': 'accepted', 'reason': clean['quality_flags']})
                else:
                    duplicate_review.append({**basic, 'reason': 'exact_duplicate', 'raw_record_json': json.dumps(raw, ensure_ascii=False)})
                    audit.append({**basic, 'result': 'exact_duplicate', 'reason': 'same_channel_order_id_same_values'})

    accepted.sort(key=lambda r: (r['order_date'], r['channel'], r['order_id']))
    audit.sort(key=lambda r: (r['source_file'], int(r['source_row'])))
    rejected.sort(key=lambda r: (r['source_file'], int(r['source_row'])))
    duplicate_review.sort(key=lambda r: (r['source_file'], int(r['source_row'])))

    counts = Counter(row['result'] for row in audit)
    metrics = {
        'raw_rows': total_rows,
        'accepted_rows': counts['accepted'],
        'rejected_rows': counts['rejected'],
        'exact_duplicate_rows': counts['exact_duplicate'],
        'conflicting_duplicate_rows': counts['conflicting_duplicate'],
        'kept_with_warnings': sum(bool(r['quality_flags']) for r in accepted),
        'completed_orders': sum(r['status'] == 'Completed' for r in accepted),
        'cancelled_orders': sum(r['status'] == 'Cancelled' for r in accepted),
        'refunded_orders': sum(r['status'] == 'Refunded' for r in accepted),
        'completed_sales_twd': f"{sum((Decimal(r['line_amount_twd']) for r in accepted if r['status'] == 'Completed'), Decimal('0')):.2f}",
    }
    assert sum(counts.values()) == total_rows

    write_csv(output_dir / 'clean_orders.csv', CLEAN_FIELDS, accepted)
    write_csv(output_dir / 'rejected_rows.csv', REVIEW_FIELDS, rejected)
    write_csv(output_dir / 'duplicate_review.csv', REVIEW_FIELDS, duplicate_review)
    write_csv(output_dir / 'audit_log.csv', AUDIT_FIELDS, audit)
    definitions = {
        'raw_rows': '三份原始檔資料列數（不含表頭）',
        'accepted_rows': '通過驗證且鍵值唯一的保留資料列數',
        'rejected_rows': '必要欄位缺失或格式無法解析而隔離的資料列數',
        'exact_duplicate_rows': '完全相同的重複訂單中移除的額外列數',
        'conflicting_duplicate_rows': '相同來源＋訂單號但欄位內容衝突，整組隔離的資料列數',
        'kept_with_warnings': '保留但具有非必要欄位缺漏或無法對應品名的資料列數',
        'completed_orders': '已完成的訂單列數',
        'cancelled_orders': '取消的訂單列數',
        'refunded_orders': '退款的訂單列數',
        'completed_sales_twd': '只計已完成訂單的商品金額（數量×單價），不含運費/折扣/稅',
    }
    write_csv(output_dir / 'quality_summary.csv', ['metric', 'value', 'definition'], [
        {'metric': k, 'value': v, 'definition': definitions[k]} for k, v in metrics.items()
    ])
    return metrics


def main():
    parser = argparse.ArgumentParser(description='Clean demo order data and generate audited CSV outputs.')
    parser.add_argument('--input-dir', type=Path, default=ROOT / 'data/raw')
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'data/output')
    parser.add_argument('--mapping', type=Path, default=ROOT / 'rules/product_mapping.csv')
    args = parser.parse_args()
    metrics = process(args.input_dir, args.output_dir, args.mapping)
    print('Data cleaning complete:')
    for key, value in metrics.items():
        print(f'  {key}: {value}')
    print(f'Output folder: {args.output_dir}')


if __name__ == '__main__':
    main()
