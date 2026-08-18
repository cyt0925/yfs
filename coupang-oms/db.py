"""SQLite 連線與 schema。

設計重點：
1. UNIQUE(po_number, sku_id) 才是真正的防重鍵。一張 PO 有多個 SKU，
   若把 po_number 設成 UNIQUE，一張 29 個 SKU 的單只會留下 1 列。
2. delivery_date / warehouse 是「整張 PO 共用、而且會被改」的欄位，
   絕對不能進主鍵，否則酷澎改倉改期會被誤判成全新單而產生幽靈列。
3. WAL 模式讓多人同時讀寫不互卡；業務層的覆蓋另外用 version 樂觀鎖擋。
"""

import os
import shutil
import sqlite3
import datetime as _dt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 資料與設定放在程式資料夾「外面」的兄弟資料夾。
#
# 這樣更新版本時只要把整個 coupang-oms 解壓縮覆蓋掉就好，不必挑檔案、
# 也不會洗掉已匯入的訂單和改好的設定。程式碼歸程式碼，資料歸資料。
DATA_DIR = os.environ.get("COUPANG_OMS_DATA") or os.path.join(
    os.path.dirname(BASE_DIR), "資料與設定")
DEFAULTS_DIR = os.path.join(BASE_DIR, "defaults")

DB_PATH = os.path.join(DATA_DIR, "database.db")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")

# 舊版把這些檔案放在程式資料夾裡，第一次啟動時自動搬過去
_MIGRATE = ("database.db", "config.json", "export_profiles.json")
_SEED = ("config.json", "export_profiles.json")

SCHEMA = """
CREATE TABLE IF NOT EXISTS orders (
    id                INTEGER PRIMARY KEY AUTOINCREMENT,

    -- === 識別（不可變）===
    po_number         TEXT NOT NULL,
    sku_id            TEXT NOT NULL,

    -- === 酷澎來源欄位：只有匯入能改，改動一律寫歷程 ===
    order_type        TEXT DEFAULT '',      -- 一般 / NS / 補單 / 拆單
    parent_po         TEXT DEFAULT '',      -- 拆單的原訂單，供血緣追蹤
    line              TEXT DEFAULT '',      -- 線別
    brand             TEXT DEFAULT '',
    product_name      TEXT DEFAULT '',
    barcode           TEXT DEFAULT '',
    yf_sku            TEXT DEFAULT '',      -- 永豐料號
    warehouse         TEXT DEFAULT '',      -- 到貨倉別（會變）
    address           TEXT DEFAULT '',
    delivery_date     TEXT DEFAULT '',      -- 交付日期 ISO（會變）
    qty_coupang       INTEGER,              -- 下單數量（酷澎原始值）
    unit              TEXT DEFAULT '',
    box_size          INTEGER,
    unit_price        REAL,
    expiry_note       TEXT DEFAULT '',
    seq_no            INTEGER,              -- 整合表的 NO，僅供參考排序

    -- === OP 自有欄位：匯入永不覆蓋 ===
    qty_ship          INTEGER,              -- 出貨數量（實際要出的）
    qty_ship_overridden INTEGER NOT NULL DEFAULT 0,
    remarks           TEXT DEFAULT '',
    receiving_note    TEXT DEFAULT '',      -- 這個品項的驗收註記（短驗/溢收…）

    -- === 警示（掛在 SKU 上：酷澎是逐品項改的）===
    needs_review      INTEGER NOT NULL DEFAULT 0,
    alert_level       TEXT DEFAULT '',      -- '' / changed / changed_after_pull / missing
    review_reason     TEXT DEFAULT '',

    -- === 稽核 ===
    version           INTEGER NOT NULL DEFAULT 1,
    source_file       TEXT DEFAULT '',
    first_seen_at     TEXT DEFAULT '',
    last_seen_at      TEXT DEFAULT '',
    created_at        TEXT DEFAULT '',
    updated_at        TEXT DEFAULT '',

    UNIQUE (po_number, sku_id)
);

CREATE INDEX IF NOT EXISTS idx_orders_po        ON orders(po_number);
CREATE INDEX IF NOT EXISTS idx_orders_delivery  ON orders(delivery_date);
CREATE INDEX IF NOT EXISTS idx_orders_brand     ON orders(brand);
CREATE INDEX IF NOT EXISTS idx_orders_line      ON orders(line);
CREATE INDEX IF NOT EXISTS idx_orders_wh        ON orders(warehouse);
CREATE INDEX IF NOT EXISTS idx_orders_review    ON orders(needs_review);

-- PO 層級的狀態。OP 是「整張單一起」拋 ERP、一起交倉庫、一起確認的，
-- 所以這三個狀態天生屬於整張 PO，不屬於個別 SKU。存成一張獨立的表
-- 而不是複製到每個 SKU 上，就不會有「同一張單的 29 個品項狀態不一致」
-- 這種資料矛盾的可能。
CREATE TABLE IF NOT EXISTS po_headers (
    po_number        TEXT PRIMARY KEY,
    po_status        TEXT NOT NULL DEFAULT '已建立',
    receiving_status TEXT NOT NULL DEFAULT '未驗收',
    is_pulled        INTEGER NOT NULL DEFAULT 0,
    pulled_at        TEXT DEFAULT '',
    pulled_by        TEXT DEFAULT '',
    pulled_batch_id  INTEGER,
    -- 建檔日：整合表沒有酷澎的實際開單日，這裡記的是「我們第一次看到
    -- 這張單的日期」，第一次匯入時填當天，之後匯入一律不再更動。
    filed_date       TEXT DEFAULT '',
    version          INTEGER NOT NULL DEFAULT 1,
    created_at       TEXT DEFAULT '',
    updated_at       TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_poh_status   ON po_headers(po_status);
CREATE INDEX IF NOT EXISTS idx_poh_pulled   ON po_headers(is_pulled);
CREATE INDEX IF NOT EXISTS idx_poh_recv     ON po_headers(receiving_status);

-- 讀取用的扁平檢視：把 PO 層級狀態接回每一列 SKU，讓查詢、篩選、匯出
-- 都能像以前一樣當成一張大表來用。寫入一律針對底層的兩張實體表，
-- 各自負責自己該負責的欄位。
CREATE VIEW IF NOT EXISTS order_rows AS
SELECT
    o.*,
    h.po_status,
    h.receiving_status,
    h.is_pulled,
    h.pulled_at,
    h.pulled_by,
    h.pulled_batch_id,
    h.filed_date,
    h.version AS po_version
FROM orders o
JOIN po_headers h ON h.po_number = o.po_number;

-- 欄位級 append-only 歷程。只新增，永不修改刪除：它最終是拿去跟酷澎
-- 對帳、釐清倉庫出錯責任的證據。
CREATE TABLE IF NOT EXISTS edit_logs (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id    INTEGER NOT NULL,
    po_number   TEXT NOT NULL,
    sku_id      TEXT NOT NULL,
    field       TEXT NOT NULL,          -- 欄位代碼，可被查詢統計
    field_label TEXT NOT NULL,          -- 中文顯示名
    old_value   TEXT DEFAULT '',
    new_value   TEXT DEFAULT '',
    operator    TEXT NOT NULL,
    source      TEXT NOT NULL,          -- import / manual / system
    note        TEXT DEFAULT '',
    changed_at  TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_logs_order ON edit_logs(order_id);
CREATE INDEX IF NOT EXISTS idx_logs_po    ON edit_logs(po_number);
CREATE INDEX IF NOT EXISTS idx_logs_field ON edit_logs(field);
CREATE INDEX IF NOT EXISTS idx_logs_time  ON edit_logs(changed_at);

-- 匯入批次：先預覽、確認後才寫入，全有全無
CREATE TABLE IF NOT EXISTS import_batches (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    filename       TEXT DEFAULT '',
    operator       TEXT DEFAULT '',
    rows_total     INTEGER DEFAULT 0,
    rows_new       INTEGER DEFAULT 0,
    rows_updated   INTEGER DEFAULT 0,
    rows_identical INTEGER DEFAULT 0,
    rows_error     INTEGER DEFAULT 0,
    committed      INTEGER NOT NULL DEFAULT 0,
    preview_json   TEXT DEFAULT '',
    created_at     TEXT DEFAULT '',
    committed_at   TEXT DEFAULT ''
);

-- 匯出批次：倉庫拿到的到底是哪一版，事後要能回答
CREATE TABLE IF NOT EXISTS export_batches (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    operator    TEXT DEFAULT '',
    profile     TEXT DEFAULT '',
    filename    TEXT DEFAULT '',
    row_count   INTEGER DEFAULT 0,
    mark_pulled INTEGER NOT NULL DEFAULT 0,
    filter_json TEXT DEFAULT '',
    created_at  TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS export_batch_items (
    batch_id  INTEGER NOT NULL,
    order_id  INTEGER NOT NULL,
    po_number TEXT NOT NULL,
    sku_id    TEXT NOT NULL,
    qty_ship  INTEGER
);

CREATE INDEX IF NOT EXISTS idx_expitems_batch ON export_batch_items(batch_id);
CREATE INDEX IF NOT EXISTS idx_expitems_order ON export_batch_items(order_id);
"""


def now():
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_conn():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")      # 多人同時讀寫不互卡
    conn.execute("PRAGMA busy_timeout=8000")     # 瞬間鎖衝突改成等待而非報錯
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def ensure_data_dir():
    """建立資料資料夾、從舊位置搬檔、補上預設設定檔。

    每次啟動都跑，但都是冪等的：已經存在的東西一律不動，
    所以使用者改過的 config.json 永遠不會被預設值蓋掉。
    """
    os.makedirs(DATA_DIR, exist_ok=True)

    moved = []
    for name in _MIGRATE:
        old = os.path.join(BASE_DIR, name)
        new = os.path.join(DATA_DIR, name)
        if os.path.exists(old) and not os.path.exists(new):
            shutil.move(old, new)
            moved.append(name)

    old_backups = os.path.join(BASE_DIR, "backups")
    if os.path.isdir(old_backups) and not os.path.isdir(BACKUP_DIR):
        shutil.move(old_backups, BACKUP_DIR)
        moved.append("backups")

    for name in _SEED:
        target = os.path.join(DATA_DIR, name)
        source = os.path.join(DEFAULTS_DIR, name)
        if not os.path.exists(target) and os.path.exists(source):
            shutil.copy2(source, target)

    return moved


def init_db():
    ensure_data_dir()
    conn = get_conn()
    try:
        conn.executescript(SCHEMA)
        conn.commit()
    finally:
        conn.close()


def backup_db(tag="import"):
    """資料庫檔就是全部身家，誤刪等於歸零。每次寫入前先留一份。"""
    if not os.path.exists(DB_PATH):
        return ""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, f"database_{tag}_{stamp}.db")
    shutil.copy2(DB_PATH, dest)
    _prune_backups(keep=30)
    return dest


def _prune_backups(keep=30):
    try:
        files = sorted(
            (os.path.join(BACKUP_DIR, f) for f in os.listdir(BACKUP_DIR)
             if f.endswith(".db")),
            key=os.path.getmtime,
            reverse=True,
        )
        for path in files[keep:]:
            os.remove(path)
    except OSError:
        pass
