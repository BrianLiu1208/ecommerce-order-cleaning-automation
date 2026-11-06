"""Create a portable, offline, customer-facing HTML case study from pipeline CSVs.

Only Python standard library; no customer data is sent anywhere.
All values are derived from the current input/output and HTML-escaped.
"""
from __future__ import annotations

import csv
import html
import json
from collections import Counter, defaultdict
from decimal import Decimal
from pathlib import Path

from clean_orders import SOURCES


def read_csv(path: Path) -> list[dict]:
    with path.open(encoding='utf-8-sig', newline='') as source:
        return list(csv.DictReader(source))


def h(value) -> str:
    return html.escape(str(value), quote=True)


def money(value) -> str:
    return f"NT$ {Decimal(str(value)):,.0f}"


def number(value) -> str:
    return f"{int(value):,}"


def pill(label: str, kind: str = 'neutral') -> str:
    return f'<span class="pill {h(kind)}">{h(label)}</span>'


def make_rows(rows: list[tuple[str, str]], muted: bool = False) -> str:
    return ''.join(
        f'<div class="kv"><span>{h(key)}</span><strong{(" class=\"soft\"" if muted else "")}>{h(val)}</strong></div>'
        for key, val in rows
    )


def render(input_dir: Path, output_dir: Path, destination: Path) -> Path:
    summary = read_csv(output_dir / 'quality_summary.csv')
    metrics = {r['metric']: r['value'] for r in summary}
    cleaned = read_csv(output_dir / 'clean_orders.csv')
    rejected = read_csv(output_dir / 'rejected_rows.csv')
    duplicate = read_csv(output_dir / 'duplicate_review.csv')
    audited = read_csv(output_dir / 'audit_log.csv')

    raw = int(metrics['raw_rows'])
    accepted = int(metrics['accepted_rows'])
    rejected_n = int(metrics['rejected_rows'])
    exact_n = int(metrics['exact_duplicate_rows'])
    conflicts_n = int(metrics['conflicting_duplicate_rows'])
    warning_n = int(metrics['kept_with_warnings'])
    if len(audited) != raw or len(cleaned) != accepted or len(rejected) != rejected_n or len(duplicate) != exact_n + conflicts_n:
        raise ValueError('HTML report input mismatch: row counts do not reconcile')
    if raw != accepted + rejected_n + exact_n + conflicts_n:
        raise ValueError('HTML report input mismatch: reconciliation failed')

    completed_actual = sum(row['status'] == 'Completed' for row in cleaned)
    amount_actual = sum((Decimal(r['line_amount_twd']) for r in cleaned if r['status'] == 'Completed'), Decimal('0'))
    if completed_actual != int(metrics['completed_orders']) or amount_actual != Decimal(metrics['completed_sales_twd']):
        raise ValueError('HTML report input mismatch: completed orders or amount do not reconcile')

    # Live before/after example: both sides are read from this run's source data.
    exemplar = next((r for r in cleaned if r['source_file'] == 'orders_marketplace.csv' and r['order_id'] == 'MKT-1005'), None)
    if exemplar is None:
        exemplar = next((r for r in cleaned if r['source_file'] == 'orders_marketplace.csv'), None)
    if exemplar is None:
        exemplar = cleaned[0] if cleaned else None
    if exemplar is None:
        raise ValueError('The demo requires at least one accepted row')

    path = input_dir / exemplar['source_file']
    originals = read_csv(path)
    source_index = int(exemplar['source_row']) - 2
    if not 0 <= source_index < len(originals):
        raise ValueError(f'Invalid source row link for {exemplar["order_id"]}')
    original = originals[source_index]
    keys = SOURCES[exemplar['source_file']]['columns']
    before = [(key, original.get(key, '')) for key in keys]
    after = [(label, exemplar[field]) for label, field in [
        ('來源渠道', 'channel'), ('訂單編號', 'order_id'), ('統一日期', 'order_date'),
        ('標準商品', 'product'), ('數量', 'quantity'), ('台幣單價', 'unit_price_twd'),
        ('商品金額', 'line_amount_twd'), ('訂單狀態', 'status'),
    ]]

    conflict_rows = [r for r in duplicate if r['reason'] == 'conflicting_duplicate']
    conflict_group = defaultdict(list)
    for record in conflict_rows:
        conflict_group[(record['channel'], record['order_id'])].append(record)
    conflict_key = next(iter(conflict_group), None)
    conflicts_html = ''
    if conflict_key:
        entries = conflict_group[conflict_key]
        info = SOURCES[entries[0]['source_file']]
        price_column = info['columns'][4]
        conflict_items = []
        for item in entries:
            record = json.loads(item['raw_record_json'])
            conflict_items.append(f'<div class="conflict-line"><span>{h(item["source_file"])} · 第 {h(item["source_row"])} 列</span>'
                                  f'<strong>{h(record.get(price_column, ""))} 元／件</strong></div>')
        conflicts_html = f'''<div class="alert">
            <div class="alert-header">{pill('需要人工確認', 'warning')} <strong>訂單 {h(conflict_key[1])}</strong></div>
            {''.join(conflict_items)}
            <p>同一渠道、同一訂單編號卻有不同售價。程式將整組移到人工審核清單，<strong>不擅自選擇任何一筆</strong>。</p>
        </div>'''

    buckets = [
        ('有效保留', accepted, 'ok'),
        ('無效資料隔離', rejected_n, 'bad'),
        ('完全重複移除', exact_n, 'neutral'),
        ('衝突資料隔離', conflicts_n, 'warning'),
    ]
    bars = ''.join(
        f'''<div class="progress-row"><div class="progress-title"><span>{pill(name, tone)}</span><b>{number(n)} 筆 <small>({n/raw*100:.1f}%)</small></b></div>
        <div class="track"><div class="fill {tone}" style="width:{n/raw*100:.3f}%"></div></div></div>'''
        for name, n, tone in buckets
    )
    reasons = Counter(r['reason'] for r in rejected)
    translations = {
        'invalid_date': '無效日期', 'invalid_unit_price': '無效單價',
        'invalid_quantity': '無效數量', 'unknown_status': '未知訂單狀態',
        'missing_order_id': '缺少訂單編號', 'missing_product': '缺少商品名稱',
    }
    reason_rows = ''.join(f'<tr><td>{h(translations.get(key, key))}</td><td class="num">{number(n)}</td></tr>' for key, n in sorted(reasons.items()))
    channel_data = defaultdict(lambda: {'count': 0, 'completed': 0, 'amount': Decimal('0')})
    for r in cleaned:
        obj = channel_data[r['channel']]
        obj['count'] += 1
        if r['status'] == 'Completed':
            obj['completed'] += 1
            obj['amount'] += Decimal(r['line_amount_twd'])
    channel_rows = ''.join(
        f'<tr><td>{h(channel)}</td><td class="num">{number(info["count"])}</td>'
        f'<td class="num">{number(info["completed"])}</td><td class="num">{h(money(info["amount"]))}</td></tr>'
        for channel, info in sorted(channel_data.items())
    )

    page = f'''<!doctype html>
<html lang="zh-Hant"><head><meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>電商訂單資料清洗｜客戶成果展示</title>
<style>
:root {{--ink:#142c3d;--muted:#607488;--line:#dfebed;--bg:#f4f8fa;--teal:#087e83;--deep:#0c3546;--green:#176d55;--amber:#9a6700;--red:#b64e4e;}}
* {{box-sizing:border-box}}
body {{margin:0;background:var(--bg);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Noto Sans TC","PingFang TC","Microsoft JhengHei",sans-serif;font-size:15px;line-height:1.65}}
a {{color:var(--teal)}}.shell{{max-width:1100px;margin:auto;padding:36px 24px 68px}}
.hero{{background:linear-gradient(112deg,#123c4c 0%,#0c5761 68%,#087e83 100%);color:#fff;border-radius:26px;padding:42px 44px;box-shadow:0 14px 44px #13435a24}}
.eyebrow{{font-weight:700;letter-spacing:.12em;text-transform:uppercase;font-size:12px;opacity:.82}}
h1{{font-size:36px;line-height:1.25;margin:10px 0 16px;letter-spacing:-.03em}}h2{{font-size:22px;margin:0 0 8px;letter-spacing:-.02em}}h3{{font-size:17px;margin:0 0 8px}}p{{margin:7px 0}}
.hero p{{max-width:780px;color:#e8f8fa;font-size:16px}}.hero-meta{{margin-top:24px;display:flex;gap:8px;flex-wrap:wrap}}
.hero-tag{{background:#ffffff20;padding:5px 12px;border-radius:999px;border:1px solid #ffffff40;font-size:12px}}
section{{margin-top:26px}}.head{{display:flex;align-items:end;justify-content:space-between;gap:12px;margin:0 0 14px}}
.desc{{color:var(--muted);font-size:14px;margin:3px 0 0}}.grid{{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}}
.kpi{{background:white;border:1px solid var(--line);border-radius:16px;padding:18px 19px;min-height:126px}}
.kpi .k{{font-size:13px;color:var(--muted)}}.kpi .v{{font-weight:800;font-size:29px;line-height:1.25;letter-spacing:-.03em;margin-top:9px;white-space:nowrap}}
.kpi .hint{{color:var(--muted);font-size:12px;margin-top:6px}}.card{{background:#fff;border:1px solid var(--line);border-radius:20px;padding:24px}}
.twocol{{display:grid;grid-template-columns:1fr 1fr;gap:16px}}.progress-row{{margin:18px 0}}.progress-title{{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;font-size:14px}}
.progress-title small{{color:var(--muted);font-weight:400}}.track{{height:9px;background:#e6eef1;border-radius:99px;overflow:hidden}}.fill{{height:100%;border-radius:99px;background:var(--teal);min-width:3px}}.fill.ok{{background:#1d8a75}}.fill.bad{{background:#c96666}}.fill.neutral{{background:#8fa1b5}}.fill.warning{{background:#d3a150}}
.pill{{display:inline-block;border-radius:999px;background:#eef3f5;padding:3px 10px;font-weight:700;font-size:12px;color:#496476}}
.pill.ok{{background:#e7f5ef;color:var(--green)}}.pill.bad{{background:#ffeded;color:var(--red)}}.pill.warning{{background:#fff4dd;color:var(--amber)}}
.kv{{display:flex;justify-content:space-between;gap:20px;border-bottom:1px solid #ecf1f3;padding:10px 0;font-size:13px}}.kv span{{color:var(--muted)}}.kv strong{{text-align:right;overflow-wrap:anywhere}}.kv:last-child{{border-bottom:0}}
.card-label{{font-size:12px;font-weight:800;letter-spacing:.04em;text-transform:uppercase;color:var(--teal);margin-bottom:10px}}
.tagline{{font-size:12px;color:var(--muted)}}.arrow{{display:flex;justify-content:center;align-items:center;color:var(--teal);font-weight:700}}
.alert{{padding:18px;background:#fff9ed;border:1px solid #efd7a7;border-radius:14px;margin-top:12px}}
.alert-header{{display:flex;align-items:center;gap:9px;margin-bottom:9px}}.conflict-line{{display:flex;justify-content:space-between;gap:10px;padding:6px 0;border-bottom:1px solid #f3e8d1;font-size:13px}}
.conflict-line:last-of-type{{border-bottom:0}}.alert p{{color:#6e5c3d;font-size:13px;margin-top:12px}}
table{{width:100%;border-collapse:collapse;font-size:13px}}th{{color:var(--muted);text-align:left;background:#f5f8fa;font-weight:700}}td,th{{padding:10px 12px;border-bottom:1px solid #e9eff1}}.num{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
.note{{font-size:12px;color:var(--muted);margin-top:13px}}.footer{{border-top:1px solid var(--line);padding-top:18px;color:var(--muted);font-size:12px;margin-top:30px}}
@media(max-width:800px){{.grid{{grid-template-columns:repeat(2,minmax(0,1fr))}}.twocol{{grid-template-columns:1fr}}h1{{font-size:27px}}.hero{{padding:30px 24px}}.shell{{padding:16px 12px 40px}}}}
@media(max-width:450px){{.grid{{grid-template-columns:1fr 1fr;gap:9px}}.kpi{{padding:14px 12px}}.kpi .v{{font-size:23px}}.card{{padding:17px}}}}
@media print{{body{{background:#fff}}.shell{{max-width:none;padding:0}}.hero{{box-shadow:none}}.card,.kpi{{break-inside:avoid}}}}
</style></head><body><main class="shell">
<header class="hero"><div class="eyebrow">DATA CLEANING · MINI CASE STUDY</div>
<h1>讓三個銷售渠道的訂單<br/>變成一份可信任的報表</h1>
<p>原始匯出檔有不同欄位名稱、日期格式、商品別名與重複訂單。用 Python 自動統一格式、隔離有疑慮的資料，並產出能追溯來源的 Excel 成果。</p>
<div class="hero-meta"><span class="hero-tag">Python 標準函式庫</span><span class="hero-tag">Excel 自動更新</span><span class="hero-tag">可追溯稽核紀錄</span><span class="hero-tag">100% 合成示範資料</span></div></header>
<section><div class="head"><div><h2>處理成果總覽</h2><p class="desc">所有數字皆從這次執行的 CSV 清洗結果計算，不是固定展示數字。</p></div></div>
<div class="grid">
<div class="kpi"><div class="k">原始資料列數</div><div class="v">{number(raw)}</div><div class="hint">來自 3 個銷售渠道</div></div>
<div class="kpi"><div class="k">有效保留</div><div class="v">{number(accepted)}</div><div class="hint">佔原始資料 {accepted/raw*100:.1f}%</div></div>
<div class="kpi"><div class="k">需處理／審核</div><div class="v">{number(rejected_n+exact_n+conflicts_n)}</div><div class="hint">{rejected_n} 無效 + {exact_n} 重複 + {conflicts_n} 衝突</div></div>
<div class="kpi"><div class="k">完成訂單商品金額</div><div class="v" style="font-size:24px">{h(money(amount_actual))}</div><div class="hint">{number(completed_actual)} 筆完成訂單 · 非財報營收</div></div>
</div></section>
<section class="twocol"><div class="card"><h2>{number(raw)} 筆資料最後去了哪裡？</h2><p class="desc">每一筆來源資料都有明確處理結果，不會默默消失。</p>
{bars}<p class="note">對帳：{raw} = {accepted} + {rejected_n} + {exact_n} + {conflicts_n}；保留資料中另有 {warning_n} 筆非致命警示。</p></div>
<div class="card"><h2>異常不是刪掉就好</h2><p class="desc">錯誤資料保留來源檔案、來源列號和原因，方便人工修正。</p>
<table><thead><tr><th>無效資料原因</th><th class="num">筆數</th></tr></thead><tbody>{reason_rows}</tbody></table>
{conflicts_html}</div></section>
<section><div class="head"><div><h2>清洗前後對照</h2><p class="desc">實際從來源 CSV 與本次 clean_orders.csv 讀取同一筆訂單：{h(exemplar['order_id'])}。</p></div></div>
<div class="twocol"><div class="card"><div class="card-label">BEFORE · 原始資料</div><h3>{h(exemplar['source_file'])}</h3>
<p class="tagline">原始檔欄位名稱與內容，保留原貌</p>{make_rows(before)}</div>
<div class="card"><div class="card-label">AFTER · 統一格式</div><h3>clean_orders.csv</h3>
<p class="tagline">與其他渠道合併後的標準欄位及數值</p>{make_rows(after)}</div></div></section>
<section class="twocol"><div class="card"><h2>清洗後跨渠道概覽</h2><p class="desc">完成訂單金額 = 數量 × 單價，只計 Completed。</p>
<table><thead><tr><th>銷售渠道</th><th class="num">保留筆數</th><th class="num">完成筆數</th><th class="num">商品金額</th></tr></thead><tbody>{channel_rows}</tbody></table></div>
<div class="card"><h2>客戶收到什麼？</h2>
<div class="kv"><span>clean_orders.csv</span><strong>可分析訂單明細</strong></div>
<div class="kv"><span>rejected_rows.csv</span><strong>無效資料清單</strong></div>
<div class="kv"><span>duplicate_review.csv</span><strong>重複與衝突清單</strong></div>
<div class="kv"><span>audit_log.csv</span><strong>每列來源與處理結果</strong></div>
<div class="kv"><span>quality_summary.csv</span><strong>品質統計</strong></div>
<div class="kv"><span>ecommerce_cleaning_demo.xlsx</span><strong>Excel 成果報表</strong></div>
<p class="note">執行 <code>python3 run_pipeline.py</code> 可同時更新 CSV、Excel 與本展示頁。</p></div></section>
<section class="card"><h2>適用情境與限制</h2><p>適合需要合併官網、電商平台、實體門市固定格式 CSV 的小型營運團隊。每筆資料預設為一筆單品訂單；真實專案需依客戶的訂單鍵、幣別、商品映射、驗收規則客製化。</p>
<p class="desc">本作品使用合成資料，沒有測量人工節省時間，因此不宣稱「節省 80% 工時」等未經驗證的績效。退款、折扣、運費、稅費均未納入正式財務營收認列。</p></section>
<footer class="footer">展示用途 · Synthetic Data Only ｜所有指標以本次 pipeline 的輸出為準。資料不會上傳至外部服務。可離線開啟此 HTML。</footer>
</main></body></html>'''
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(page, encoding='utf-8')
    return destination


if __name__ == '__main__':
    from clean_orders import ROOT
    print(render(ROOT / 'data/raw', ROOT / 'data/output', ROOT / 'reports/customer_case_study.html'))
