#!/usr/bin/env python3
"""Generate deterministic *synthetic* source exports for the portfolio demo.

Warning: --force overwrites only the three known synthetic CSV filenames.
"""
from __future__ import annotations

import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent


def generate(root: Path, force=False):
    root.mkdir(parents=True, exist_ok=True)
    rng = random.Random(20261106)
    products = [
        ('無線滑鼠', 'Wireless Mouse', '滑鼠 - 無線', 699),
        ('機械鍵盤', 'Mechanical Keyboard', '機械式鍵盤', 1890),
        ('保溫杯', 'Insulated Bottle', '真空保溫杯', 490),
        ('USB-C 集線器', 'USB-C Hub', 'USB C 集線器', 1299),
        ('筆記本', 'Notebook', 'A5 筆記本', 129),
    ]
    configs = [
        ('orders_shop.csv', ['訂單編號', '下單日期', '商品', '數量', '單價', '客戶ID', '狀態'], 'WEB', 0),
        ('orders_marketplace.csv', ['Order_ID', 'Purchased_At', 'Item', 'Quantity', 'Price', 'Buyer_ID', 'Order_Status'], 'MKT', 1),
        ('orders_pos.csv', ['交易編號', '交易時間', '品項', '件數', '售價', '會員編號', '交易狀態'], 'POS', 2),
    ]
    all_counts = {}
    for filename, headers, prefix, source_idx in configs:
        output = root / filename
        if output.exists() and not force:
            raise FileExistsError(f'{output} exists. Use --force to regenerate DEMO inputs.')
        rows = []
        for ix in range(25):
            pname, ename, altname, baseprice = products[(ix + source_idx) % len(products)]
            names = (pname, ename, altname)
            day = date(2026, 10, 1) + timedelta(days=(ix * 7 + source_idx * 3) % 31)
            if source_idx == 0:
                date_text = (day.strftime('%Y/%m/%d') if ix % 2 else day.isoformat())
                price = f'NT${baseprice:,}' if ix % 3 else str(baseprice)
                status = '完成' if ix % 7 else '取消'
            elif source_idx == 1:
                date_text = day.strftime('%m/%d/%Y') if ix % 2 else day.strftime('%Y-%m-%d %H:%M')
                price = f'{baseprice:,.2f}' if ix % 3 else f'${baseprice:,}'
                status = 'paid' if ix % 6 else 'refunded'
            else:
                date_text = day.strftime('%Y.%m.%d') if ix % 2 else day.strftime('%Y年%m月%d日')
                price = str(baseprice)
                status = '已完成' if ix % 8 else '已退款'
            quantity = str(1 + (ix % 3))
            customer = f'C{(ix * 13 + source_idx * 9) % 41 + 1:03}'
            rows.append([f'{prefix}-{1001 + ix}', date_text, names[ix % 3], quantity, price, customer, status])
        if source_idx == 0:
            rows.append(list(rows[0]))  # exact duplicate
            rows += [
                [' ', '2026/10/11', '無線滑鼠', '1', '699', 'C110', '完成'],
                ['WEB-X1', '2026/10/12', '機械鍵盤', '1', '', 'C111', '完成'],
                ['WEB-X2', '2026/10/12', '旅行轉接頭', '1', '399', 'C112', '完成'],  # unmapped, kept with warning
            ]
        elif source_idx == 1:
            rows.append(list(rows[1]))  # exact duplicate
            rows += [
                ['MKT-X1', '2026-33-99', 'Notebook', '2', '129', 'C113', 'paid'],
                ['MKT-X2', '10/15/2026', 'USB-C Hub', '1', 'unknown', 'C114', 'paid'],
                ['MKT-X3', '10/17/2026', 'Insulated Bottle', '1', '490', '', 'paid'], # optional blank
            ]
        else:
            conflict = list(rows[2])
            conflict[4] = str(int(conflict[4]) + 100)  # conflicting duplicate, both quarantined
            rows.append(conflict)
            rows += [
                ['POS-X1', '2026.10.05', 'A5 筆記本', '-2', '129', 'C115', '已完成'],
                ['POS-X2', '2026.10.15', '無線滑鼠', '1', '699', 'C116', '待處理'],
                ['POS-X3', '2026.10.19', '機械式鍵盤', '1', '1890', '', '已完成'],
            ]
        with output.open('w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(rows)
        all_counts[filename] = len(rows)
    return all_counts


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'data/raw')
    parser.add_argument('--force', action='store_true', help='Overwrite only the three demo source files.')
    args = parser.parse_args()
    for filename, count in generate(args.output_dir, args.force).items():
        print(f'{filename}: {count} sample records')
