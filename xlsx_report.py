"""Portable Excel XLSX report writer. Uses only Python's standard library.

The cleaning logic remains in clean_orders.py. This file packages CSV outputs
into an Open XML spreadsheet, with cached KPI values and Excel formulas.
"""
from __future__ import annotations

import csv
from collections import Counter, defaultdict
from datetime import datetime, timezone
from decimal import Decimal
from pathlib import Path
import re
from tempfile import NamedTemporaryFile
from xml.sax.saxutils import escape
from zipfile import ZipFile, ZIP_DEFLATED

ROOT = Path(__file__).resolve().parent
SHEETS = [
    'Overview', 'Clean_Orders', 'Quality_Checks', 'Rejected_Rows',
    'Duplicate_Review', 'Audit_Log', 'Dictionary'
]

# Colors and number formats below belong to the spreadsheet document.
STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;NT$&quot;#,##0.00"/></numFmts>
<fonts count="4"><font><sz val="11"/><name val="Aptos"/></font>
<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><b/><sz val="18"/><color rgb="FFFFFFFF"/><name val="Aptos Display"/></font>
<font><b/><sz val="16"/><color rgb="FF172D47"/><name val="Aptos"/></font></fonts>
<fills count="8"><fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF172D47"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF0C827B"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE8F5F2"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF4F7FB"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFE9C5"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFCE6E8"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="12">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="5" borderId="0" xfId="0" applyFill="1" applyFont="1"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="0" fontId="3" fillId="4" borderId="0" xfId="0" applyFill="1" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
<xf numFmtId="164" fontId="3" fillId="5" borderId="0" xfId="0" applyFill="1" applyFont="1" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def read_csv(path):
    with path.open(encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def colname(index):
    letters = ''
    while index:
        index, digit = divmod(index - 1, 26)
        letters = chr(65 + digit) + letters
    return letters


def clean_xml_text(value):
    text = str(value or '')
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)
    return escape(text)


class Sheet:
    def __init__(self, name, widths=None, freeze=1, tab=None):
        self.name, self.cells = name, {}
        self.widths = widths or {}
        self.freeze, self.merges, self.filter_ref, self.tab = freeze, [], None, tab

    def put(self, row, col, value, style=0, formula=None):
        self.cells[(row, col)] = (value, style, formula)

    def matrix(self, start_row, start_col, matrix, header=False, styles=None):
        for r, values in enumerate(matrix, start_row):
            for c, value in enumerate(values, start_col):
                style = 1 if header and r == start_row else (styles.get(c, 0) if styles else 0)
                self.put(r, c, value, style)

    def merge(self, ref):
        self.merges.append(ref)

    def xml(self):
        if self.cells:
            max_row = max(r for r, _ in self.cells)
            max_col = max(c for _, c in self.cells)
        else:
            max_row, max_col = 1, 1
        output = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
                  f'<dimension ref="A1:{colname(max_col)}{max_row}"/>',
                  '<sheetViews><sheetView workbookViewId="0" showGridLines="0">']
        if self.freeze:
            output.append(f'<pane ySplit="{self.freeze}" topLeftCell="A{self.freeze + 1}" activePane="bottomLeft" state="frozen"/>')
        output.append('</sheetView></sheetViews>')
        if self.widths:
            output.append('<cols>')
            for col, width in sorted(self.widths.items()):
                output.append(f'<col min="{col}" max="{col}" width="{width}" customWidth="1"/>')
            output.append('</cols>')
        output.append('<sheetData>')
        rows = defaultdict(list)
        for (row, col), spec in self.cells.items():
            rows[row].append((col, spec))
        for row in sorted(rows):
            output.append(f'<row r="{row}" ht="{27 if row == 1 else 19}" customHeight="1">')
            for col, (value, style, formula) in sorted(rows[row]):
                loc = f'{colname(col)}{row}'
                if formula is not None:
                    val = value if value is not None else 0
                    output.append(f'<c r="{loc}" s="{style}"><f>{clean_xml_text(formula.lstrip("="))}</f><v>{val}</v></c>')
                elif isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
                    output.append(f'<c r="{loc}" s="{style}"><v>{value}</v></c>')
                else:
                    text = clean_xml_text(value)
                    output.append(f'<c r="{loc}" t="inlineStr" s="{style}"><is><t xml:space="preserve">{text}</t></is></c>')
            output.append('</row>')
        output.append('</sheetData>')
        if self.merges:
            output.append(f'<mergeCells count="{len(self.merges)}">')
            output.extend(f'<mergeCell ref="{ref}"/>' for ref in self.merges)
            output.append('</mergeCells>')
        if self.filter_ref:
            output.append(f'<autoFilter ref="{self.filter_ref}"/>')
        output.append('<pageMargins left="0.35" right="0.35" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        output.append('</worksheet>')
        return ''.join(output).encode('utf-8')


def table_sheet(name, headers, records, widths=None, money_cols=(), warn_column=None):
    sheet = Sheet(name, widths=widths)
    sheet.matrix(1, 1, [headers], header=True)
    for idx, row in enumerate(records, 2):
        for col, value in enumerate(row, 1):
            style = 5 if col in money_cols else 0
            if warn_column == col and value:
                style = 6
            sheet.put(idx, col, value, style)
    sheet.filter_ref = f'A1:{colname(len(headers))}{max(1, len(records) + 1)}'
    return sheet


def build_workbook(output_dir: Path, report_path: Path):
    output_dir, report_path = Path(output_dir), Path(report_path)
    clean = read_csv(output_dir / 'clean_orders.csv')
    quality = read_csv(output_dir / 'quality_summary.csv')
    rejected = read_csv(output_dir / 'rejected_rows.csv')
    duplicated = read_csv(output_dir / 'duplicate_review.csv')
    audit = read_csv(output_dir / 'audit_log.csv')
    metrics = {r['metric']: r['value'] for r in quality}
    n = lambda key: int(metrics[key])
    money = Decimal(metrics['completed_sales_twd'])
    by_channel = Counter()
    channel_money = defaultdict(Decimal)
    statuses = Counter()
    for row in clean:
        statuses[row['status']] += 1
        if row['status'] == 'Completed':
            by_channel[row['channel']] += 1
            channel_money[row['channel']] += Decimal(row['line_amount_twd'])
    assert n('raw_rows') == n('accepted_rows') + n('rejected_rows') + n('exact_duplicate_rows') + n('conflicting_duplicate_rows'), 'Source reconciliation failed'
    assert sum(channel_money.values(), Decimal(0)) == money, 'Completed sales mismatch'

    overview = Sheet('Overview', widths={1:27, 2:16, 3:22, 4:11, 5:25, 6:16, 7:21, 8:24}, freeze=3)
    overview.put(1, 1, 'ECOMMERCE  /  DATA QUALITY & ORDER CLEANING', 2)
    overview.merge('A1:H2')
    overview.put(3, 1, 'DEMO DATA · 3 CHANNELS · TWD · SYNTHETIC INFORMATION ONLY', 3)
    overview.merge('A3:H3')
    for col, label, value, formula, style in [
        (1, '保留訂單', n('accepted_rows'), 'COUNTA(Clean_Orders!B2:B5000)', 4),
        (3, '完成訂單', n('completed_orders'), 'COUNTIF(Clean_Orders!J2:J5000,"Completed")', 4),
        (5, '完成商品金額 TWD', float(money), 'SUMIF(Clean_Orders!J2:J5000,"Completed",Clean_Orders!H2:H5000)', 10),
        (8, '隔離待確認', n('rejected_rows') + n('conflicting_duplicate_rows'), 'COUNTIF(Audit_Log!E2:E5000,"rejected")+COUNTIF(Audit_Log!E2:E5000,"conflicting_duplicate")', 4),
    ]:
        overview.put(5, col, label, 3)
        overview.put(6, col, value, style, formula)
    overview.put(9, 1, '銷售渠道', 1)
    overview.put(9, 2, '完成筆數', 1)
    overview.put(9, 3, '商品金額 (TWD)', 1)
    for ix, channel in enumerate(('官網','電商平台','實體門市'), 10):
        overview.put(ix, 1, channel)
        overview.put(ix, 2, by_channel[channel], formula=f'COUNTIFS(Clean_Orders!A2:A5000,A{ix},Clean_Orders!J2:J5000,"Completed")')
        overview.put(ix, 3, float(channel_money[channel]), 5, formula=f'SUMIFS(Clean_Orders!H2:H5000,Clean_Orders!A2:A5000,A{ix},Clean_Orders!J2:J5000,"Completed")')
    overview.put(15, 1, '訂單狀態', 1)
    overview.put(15, 2, '筆數', 1)
    for ix, status in enumerate(('Completed','Cancelled','Refunded'), 16):
        overview.put(ix, 1, status)
        overview.put(ix, 2, statuses[status], formula=f'COUNTIF(Clean_Orders!J2:J5000,A{ix})')
    overview.put(22, 1, '資料品質核對', 1)
    overview.put(22, 2, '筆數', 1)
    for ix, (label, key) in enumerate([('原始資料', 'raw_rows'), ('保留資料', 'accepted_rows'), ('無效資料隔離', 'rejected_rows'), ('完全重複移除', 'exact_duplicate_rows'), ('衝突重複隔離', 'conflicting_duplicate_rows')],23):
        overview.put(ix, 1, label)
        overview.put(ix, 2, n(key))
    overview.put(29, 1, '差異（應為 0）', 8)
    overview.put(29, 2, 0, 8, 'B23-SUM(B24:B27)')
    overview.put(32, 1, '說明：完成商品金額 = 完成訂單的數量 × 單價；未包含折扣、運費、稅及會計認列。', 6)
    overview.merge('A32:H32')
    overview.put(34, 1, '非真實交易資料。顏色標記的錯誤與衝突資料不會納入清洗完成表。', 3)
    overview.merge('A34:H34')

    # Data tab with numeric prices, calculated Excel line-amount formulas and source traceability.
    headers = ['渠道','訂單號','日期','標準商品','原始商品名稱','數量','單價 TWD','商品金額 TWD','顧客 ID','狀態','品質警示','原始檔案','來源行號']
    rows = []
    for item in clean:
        rows.append([item['channel'], item['order_id'], item['order_date'], item['product'], item['original_product'],
                    int(item['quantity']), float(item['unit_price_twd']), float(item['line_amount_twd']),
                    item['customer_id'], item['status'], item['quality_flags'], item['source_file'], int(item['source_row'])])
    orders = table_sheet('Clean_Orders', headers, rows,
                         {1:18,2:20,3:17,4:23,5:23,6:12,7:18,8:19,9:17,10:16,11:37,12:29,13:15},
                         money_cols=(7,8), warn_column=11)
    for idx, item in enumerate(rows, 2):
        orders.put(idx, 8, item[7], 5, f'F{idx}*G{idx}')

    qc = table_sheet('Quality_Checks', ['Metric','Value','定義'],
                     [[row['metric'], float(row['value']) if row['metric'] == 'completed_sales_twd' else int(row['value']), row['definition']] for row in quality],
                     {1:29,2:22,3:80})
    qc.put(13, 1, '原始資料對帳差異（應為 0）', 8)
    qc.put(13, 2, 0, 8, 'B2-SUM(B3:B6)')
    qc.put(11, 2, float(money), 5)

    cols = ['原始檔案','來源行號','渠道','訂單號','原因','原始 JSON（人工審核）']
    rejected_sheet = table_sheet('Rejected_Rows', cols,
           [[x['source_file'],int(x['source_row']),x['channel'],x['order_id'],x['reason'],x['raw_record_json']] for x in rejected],
           {1:27,2:17,3:19,4:19,5:31,6:85}, warn_column=5)
    dup_sheet = table_sheet('Duplicate_Review',cols,
           [[x['source_file'],int(x['source_row']),x['channel'],x['order_id'],x['reason'],x['raw_record_json']] for x in duplicated],
           {1:27,2:17,3:19,4:19,5:31,6:85},warn_column=5)
    audit_sheet = table_sheet('Audit_Log', ['原始檔案','行號','渠道','訂單號','處理結果','原因／警示'],
           [[x['source_file'],int(x['source_row']),x['channel'],x['order_id'],x['result'],x['reason']] for x in audit],
           {1:29,2:14,3:18,4:22,5:29,6:47},warn_column=6)
    definitions = [
        ('channel','來源渠道','官網／電商平台／實體門市','channel + order_id 作為唯一鍵'),
        ('order_id','來源內的訂單識別碼','POS-1003','相同來源但內容不同的訂單整組隔離'),
        ('order_date','標準化日期','2026-10-21','不可解析則隔離'),
        ('product','標準商品名稱','A5 筆記本','使用商品對照表'),
        ('quantity','數量','3','必須為正整數'),
        ('unit_price_twd','台幣單價','129.00','必須為正數'),
        ('line_amount_twd','商品金額','387.00','Excel 公式：單價 × 數量'),
        ('status','訂單狀態','Completed','Completed／Cancelled／Refunded'),
        ('quality_flags','資料品質警示','unmapped_product','非致命問題保留並標示'),
        ('source_file','來源 CSV','orders_pos.csv','可回溯到原始資料'),
        ('source_row','原始行號','4','原始資料含表頭，從第 2 行起算'),
    ]
    dictionary = table_sheet('Dictionary', ['程式欄位','中文定義','範例','資料規則'], definitions, {1:29,2:30,3:30,4:60})
    sheets = [overview, orders, qc, rejected_sheet, dup_sheet, audit_sheet, dictionary]

    content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
      '<Default Extension="xml" ContentType="application/xml"/>',
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>']
    for i in range(1, len(sheets)+1):
        content_types.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    content_types.append('</Types>')
    workbook_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      ''.join(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"/>' for i, name in enumerate(SHEETS,1)) +
      '</sheets><calcPr calcMode="auto" fullCalcOnLoad="1"/></workbook>')
    workbook_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets)+1)) +
      f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>')
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
      '</Relationships>')
    now = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    core_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
      'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Ecommerce Order Cleaning Demo</dc:title>'
      f'<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created></cp:coreProperties>')

    report_path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(dir=report_path.parent, suffix='.xlsx', delete=False) as tmp:
        tmp_path = Path(tmp.name)
    try:
        with ZipFile(tmp_path, 'w', compression=ZIP_DEFLATED) as archive:
            archive.writestr('[Content_Types].xml', ''.join(content_types))
            archive.writestr('_rels/.rels', root_rels)
            archive.writestr('docProps/core.xml', core_xml)
            archive.writestr('xl/workbook.xml', workbook_xml)
            archive.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
            archive.writestr('xl/styles.xml', STYLES)
            for i, sheet in enumerate(sheets, 1):
                archive.writestr(f'xl/worksheets/sheet{i}.xml', sheet.xml())
        with ZipFile(tmp_path) as archive:
            if archive.testzip() is not None:
                raise RuntimeError('XLSX archive integrity check failed')
        tmp_path.chmod(0o644)
        tmp_path.replace(report_path)
    finally:
        tmp_path.unlink(missing_ok=True)
    return report_path


if __name__ == '__main__':
    result = build_workbook(ROOT / 'data/output', ROOT / 'reports/ecommerce_cleaning_demo.xlsx')
    print('Excel report updated:', result)
