# Ecommerce Order Cleaning Automation ｜電商訂單資料清洗作品（Demo v3）

> **資料全部為程式產生的模擬資料，非真實客戶或個資。**
>
> 目標客戶：需要合併官網／電商平台／門市訂單的電商與小型營運團隊。

## 這個專案解決什麼問題？

三個來源的 Excel/CSV 匯出檔，欄名、日期、價格、品名及訂單狀態不一致。手動整理費時，容易重複計算，也難以追查問題來源。

本作品用 **Python 標準函式庫** 實作可重複執行的清洗流程，並產生可每次重新更新的 Excel `.xlsx` 報表及離線客戶展示 HTML。整套程式可在 macOS 的 Python 3.12.2 執行，不依賴 Windows Excel，也不需要另裝套件。

### 功能

1. 讀取三種來源與不同欄名的 UTF-8 CSV。
2. 統一日期、訂單狀態、商品名稱、台幣單價及數量。
3. 移除完全重複資料；對於 **相同渠道＋訂單號、但內容衝突的訂單**，全部移至人工審核清單，不任意選一筆。
4. 核心欄位無效的訂單另存 `rejected_rows.csv`，不偷偷填造資料。
5. 非必要欄位缺失、未知商品名稱則保留並加上旗標。
6. 為每筆原始資料建立稽核紀錄，確保**輸入列數 = 保留＋拒收＋完全重複＋衝突重複**。
7. 提供處理結果 CSV 和用 Excel 開啟的示範分析報表。

### 實際測試結果（隨附資料）

| 指標 | 結果 |
|---|---:|
| 三份原始訂單資料 | 87 筆 |
| 保留且唯一訂單 | 77 筆 |
| 無效資料隔離 | 6 筆 |
| 完全重複而移除 | 2 筆 |
| 衝突訂單需人工審核 | 2 筆 |
| 保留但帶有警示旗標 | 3 筆 |
| 完成狀態的訂單 | 64 筆 |
| 完成訂單商品金額 | NT$117,585 |

> 「商品金額」僅等於數量 × 單價，且只納入 `Completed`；沒有運費、折扣、退貨會計處理或稅務認列，不可當作正式財務營收。

## 專案結構

```text
ecommerce_data_cleaning_v3/
├── README.md
├── generate_sample_data.py       # 產生固定的模擬資料
├── clean_orders.py               # 正式資料清洗邏輯（純 Python）
├── run_pipeline.py               # 一鍵執行：清洗 + Excel + HTML 展示更新（建議使用）
├── customer_report.py            # 從實際輸出資料建立離線客戶展示 HTML
├── xlsx_report.py                # 使用標準函式庫產生 Excel `.xlsx`
├── rules/
│   └── product_mapping.csv       # 可編輯的商品名稱對照表
├── data/
│   ├── raw/
│   │   ├── orders_shop.csv
│   │   ├── orders_marketplace.csv
│   │   └── orders_pos.csv
│   └── output/
│       ├── clean_orders.csv
│       ├── rejected_rows.csv
│       ├── duplicate_review.csv
│       ├── audit_log.csv
│       └── quality_summary.csv
├── reports/
│   ├── ecommerce_cleaning_demo.xlsx
│   └── customer_case_study.html
└── tests/
    ├── test_clean_orders.py
    ├── test_xlsx_report.py
    └── test_customer_report.py
```

## Mac 快速開始（不用另裝 Python 套件）

本專案所有程式只使用 Python 標準函式庫，建議 Python **3.10+**。請**先查詢你目前的 Python 版本與虛擬環境**，不需要重裝 Python。

```bash
cd ecommerce-order-cleaning-automation
python3 --version
python3 run_pipeline.py
python3 -m unittest discover -s tests -v
```

執行成功後，新的 `data/output/` CSV、`reports/ecommerce_cleaning_demo.xlsx` 和 `reports/customer_case_study.html` **會同時更新**，無須手動匯入、無須第三方套件。使用 Mac 的 Microsoft Excel 或 Numbers 開啟 Excel 檔即可。建議更新報表前先關閉舊版報表，避免編輯衝突。

### 重新產生模擬資料

```bash
python3 generate_sample_data.py --force
python3 run_pipeline.py
```

⚠️ `--force` 會覆蓋 `data/raw/` 內的三份**同名示範 CSV**，不要在同名檔案中存放正式客戶資料。

## 下一步：驗證 Excel 會隨來源資料更新

1. 執行 `python3 run_pipeline.py`，再打開 `reports/ecommerce_cleaning_demo.xlsx`。
2. 未修改範例資料時，Overview 的保留訂單應是 **77**，完成商品金額應是 **NT$117,585**，對帳差異為 **0**。
3. 若修改了 `data/raw/` 內的來源資料，先關閉 Excel，再執行同一條指令，報表就會依新的結果重建。
4. 自動測試 `test_xlsx_report.py` 會在暫存資料夾額外加入一筆完成訂單，檢查報表確實更新為 **78 筆**、**NT$118,585**。不會修改你下載的示範來源資料。

本版 Excel 含 `Overview`、`Clean_Orders`、`Quality_Checks`、`Rejected_Rows`、`Duplicate_Review`、`Audit_Log`、`Dictionary` 七個工作表。報表含 Excel 公式與已計算好的 KPI 值；不依賴巨集。

## 本版驗收條件

- 執行清洗程式沒有錯誤。
- `raw_rows = accepted_rows + rejected_rows + exact_duplicate_rows + conflicting_duplicate_rows`。
- 相同渠道＋訂單號沒有重複出現在 `clean_orders.csv`。
- 衝突的 `POS-1003` 不會被隨意保留。
- `completed_sales_twd` 僅統計 `Completed` 訂單，示範資料結果為 `117585.00`。
- 單元測試全部通過。

## 常見欄位與規則

| 欄位 | 說明 |
|---|---|
| `channel` | 固定為官網、電商平台、實體門市之一 |
| `order_id` | 與 `channel` 一起形成唯一訂單鍵 |
| `order_date` | ISO `YYYY-MM-DD`，未解析者拒收 |
| `product` | 依 `rules/product_mapping.csv` 對應標準品名；未對應則保留並標示 |
| `unit_price_twd` | 台幣正數、小數最多兩位，清除 `NT$` / `$` / 千分位 |
| `line_amount_twd` | 數量 × 單價，不含運費／折扣／稅 |
| `status` | `Completed`、`Cancelled` 或 `Refunded`，未知狀態隔離 |
| `quality_flags` | 非致命警示，例如客戶 ID 空白、未知商品名 |
| `source_file`, `source_row` | 回溯到原始檔案及行號 |

## 已知限制與下一版方向

- v3 版已支援 `python3 run_pipeline.py` **自動更新 Excel 和客戶展示頁**。若只執行 `python3 clean_orders.py`，仍只會更新 CSV。
- 報表以內建 Excel 公式及數值儲存，未提供大型動態圖表或 PivotTable。
- 每筆來源資料預設是一筆單品訂單，未涵蓋單張訂單多個品項的情境。
- 目前未處理折扣、運費、退款金額、幣別轉換、資料庫同步、排程或多使用者權限。
- 狀態 `paid` 在本案例視為完成訂單，實際客戶的收入認列規則須另外確認。
- CSV 若含敏感個資，應先評估保存位置、遮罩方式及安全權限；不要把正式客戶資料上傳 GitHub。
- 交付實際客戶前，需針對對方的欄位規格、訂單鍵及驗收方式重新設定。

## 接案展示用一句話

> 將多個平台的訂單 CSV 自動統一、清洗與核對，輸出可追蹤的 Excel 報表及錯誤明細，減少人工整理時間與重複計算風險。


## v3 新增｜客戶展示頁（Phase 4）

`python3 run_pipeline.py` 現在除了 CSV 和 Excel，還會產生完全離線的 `reports/customer_case_study.html`。

可以在 Mac 上直接開啟：

```bash
open reports/customer_case_study.html
```

展示頁會自動產出：

- 原始資料、保留、錯誤隔離、重複處理的對帳指標和視覺化。
- **一筆實際來源訂單**的清洗前後對照，依來源行號連結。
- 衝突訂單實例（例如 `POS-1003` 售價不一致）與人工處理理由。
- 各銷售渠道已完成訂單之商品金額彙總。
- 客戶會收到的檔案清單、已知限制、資料處理說明。

所有數值從本次 `data/output` 讀取，不是固定 KPI；新增訂單後 HTML 跟 Excel 一起更新。
測試 `test_customer_report.py` 驗證 HTML 會跟著新增的 1 筆、NT$1,000 訂單更新。

**展示使用注意事項：**
1. 全部都是合成示範資料，勿當真實客戶成果或業績。
2. 沒有測試人工處理時間，因此**不要宣稱節省多少工時百分比**。
3. 商品金額只計 `Completed`，不等於正式營業收入。
4. HTML 有資料來源內容，不要將未去識別化的客戶個資公開上網。
5. 本階段沒有修改 v2 清洗的核心規則與既有 Excel 格式，只新增一個展示輸出。


## GitHub 公開展示說明

此儲存庫所附輸入檔均為**模擬資料**，並無真實客戶資料。正式接案的私有資料與密鑰請勿提交至此公開儲存庫。接案用途、報價及提案範例見 [PORTFOLIO.md](PORTFOLIO.md)。
