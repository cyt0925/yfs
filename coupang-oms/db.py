"""資料庫連線與 schema。

支援兩種後端：
1. SQLite（預設）——沒設 DATABASE_URL 環境變數時使用，檔案存在本機，
   給自己電腦上單機測試用。
2. PostgreSQL（Supabase）——設了 DATABASE_URL 就切過去，資料存在雲端，
   不管程式部署到哪台機器、重新部署幾次，資料都不會不見，多人也能
   同時連到同一份。正式讓同事共用時走這條路。

兩種後端共用同一套呼叫方式（`conn.execute(sql, params)`、`row["欄位"]`、
`?` 佔位符），呼叫端（app.py／importer.py）完全不用管現在接的是哪個
資料庫——差異全部封裝在這個檔案裡的 _PgConnection／_PgCursor。

設計重點：
1. UNIQUE(po_number, sku_id) 才是真正的防重鍵。一張 PO 有多個 SKU，
   若把 po_number 設成 UNIQUE，一張 29 個 SKU 的單只會留下 1 列。
2. delivery_date / warehouse 是「整張 PO 共用、而且會被改」的欄位，
   絕對不能進主鍵，否則酷澎改倉改期會被誤判成全新單而產生幽靈列。
3. SQLite 用 WAL 模式讓多人同時讀寫不互卡；PostgreSQL 本來就是多人
   資料庫，不需要這個。業務層的覆蓋另外用 version 樂觀鎖擋，兩邊都靠
   這一層，不依賴資料庫層的鎖。
"""

import json
import os
import shutil
import sqlite3
import datetime as _dt

import psycopg2
import psycopg2.extras

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATABASE_URL = (os.environ.get("DATABASE_URL") or "").strip()
IS_POSTGRES = bool(DATABASE_URL)

# 資料與設定放在程式資料夾「外面」的兄弟資料夾。
#
# 這樣更新版本時只要把整個 coupang-oms 解壓縮覆蓋掉就好，不必挑檔案、
# 也不會洗掉已匯入的訂單和改好的設定。程式碼歸程式碼，資料歸資料。
# （PostgreSQL 模式下訂單資料跟帳號設定都在雲端資料庫裡，這個資料夾
# 只剩備份檔還會用到。）
DATA_DIR = os.environ.get("COUPANG_OMS_DATA") or os.path.join(
    os.path.dirname(BASE_DIR), "資料與設定")
DEFAULTS_DIR = os.path.join(BASE_DIR, "defaults")

DB_PATH = os.path.join(DATA_DIR, "database.db")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")

# 舊版把這些檔案放在程式資料夾裡，第一次啟動時自動搬過去
_MIGRATE = ("database.db", "config.json", "export_profiles.json")
_SEED = ("config.json", "export_profiles.json", "users.json")

# ---------------------------------------------------------------- SQLite schema

SCHEMA_SQLITE = """
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
    -- OP 跟酷澎談好之後，會先在系統裡把交期／倉別改掉，但酷澎後台通常
    -- 要過一兩天才更新。標記成「人工調整過」之後，匯入就不再覆蓋這兩個
    -- 欄位，只會亮燈提醒「酷澎的檔案還是舊值、還沒同步」——既不會洗掉
    -- OP 談好的結果，也不會讓兩邊不一致這件事被藏起來。
    delivery_date_overridden INTEGER NOT NULL DEFAULT 0,
    warehouse_overridden     INTEGER NOT NULL DEFAULT 0,
    box_size_overridden      INTEGER NOT NULL DEFAULT 0,
    order_type_overridden    INTEGER NOT NULL DEFAULT 0,
    -- 備註已搬到 po_headers（整張單共用，不分品項）。這欄改名保留，
    -- 只是既有資料庫的搬遷來源，程式不會再讀寫它。
    remarks_legacy    TEXT DEFAULT '',
    receiving_note    TEXT DEFAULT '',      -- 這個品項的驗收註記（短驗/溢收…）

    -- 實際驗入數量：小真在酷澎後台按批次驗收工具抓回來的「收貨數量」
    -- （receivedQty），跟系統無關的外部資料來源，只能透過
    -- /api/sync/verified-qty 寫入，畫面上唯讀。NULL 代表這個品項還沒
    -- 同步過，跟「同步回來是 0」要分得開，所以不能用
    -- INTEGER NOT NULL DEFAULT 0。
    actual_verified_qty    INTEGER,
    actual_verified_at     TEXT DEFAULT '',

    -- === 警示（掛在 SKU 上：酷澎是逐品項改的）===
    needs_review      INTEGER NOT NULL DEFAULT 0,
    alert_level       TEXT DEFAULT '',      -- '' / changed / changed_after_pull / missing
    review_reason     TEXT DEFAULT '',

    -- 酷澎把這個品項的數量下修到 0 時，後台匯出的整合表會直接整列消失，
    -- 不會留一列數量 0——沒有這個欄位的話，「資料庫有、本次檔案沒有」
    -- 的既有保護規則會讓這個品項的舊數量永遠卡住，PO 總數就悄悄跟酷澎
    -- 兜不起來、也沒有任何警示。標記起來但不刪列，出貨數量歸零、不計入
    -- PO 總數，保留成稽核紀錄；之後如果這個品項又出現在檔案裡，會自動
    -- 解除標記、當成正常異動重新同步。
    removed_from_coupang INTEGER NOT NULL DEFAULT 0,

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
    -- 配送方式：整合表沒有這欄，是內部自己要用的分類，一律空白開始，
    -- 匯入永遠不會覆蓋。
    shipping_method  TEXT DEFAULT '',
    -- 個人標記：全公司共用一份，OP 自己想特別留意哪張單就點一下標起來，
    -- 只有開／關兩種狀態，不分是誰標的。
    flagged          INTEGER NOT NULL DEFAULT 0,
    -- 備註：整張單共用一份，不像驗收註記是逐品項各自記錄。原本存在
    -- orders.remarks（每個品項各自一份），改放這裡，比較符合「這是這
    -- 張單的備註」的實際用法。
    remarks          TEXT DEFAULT '',
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
--
-- 這裡故意不建 view：新裝機是一張空的 orders／po_headers，這個 schema
-- 腳本執行完馬上就有新欄位，但既有安裝的表可能還沒補到新欄位（要等
-- init_db() 稍後跑 _migrate_columns 才會補），如果在這裡就建 view、
-- 引用到還不存在的欄位，既有安裝會直接炸在「no such column」。
-- view 統一交給 init_db() 補完欄位之後，用 ORDER_ROWS_VIEW 建一次。

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

-- app_settings 只有 PostgreSQL 模式會用到（帳號密碼、下拉選單設定移進
-- 資料庫，不再放檔案）；SQLite 模式建這張表但不會有人寫入，留著純粹
-- 是為了讓兩邊 schema 盡量長一樣，之後改起來比較不會漏改一邊。
CREATE TABLE IF NOT EXISTS app_settings (
    key        TEXT PRIMARY KEY,
    value      TEXT NOT NULL,
    updated_at TEXT DEFAULT ''
);

-- 每個人自己的簽名圖：登入誰、蓋的就是誰的簽名，跟修改歷程記登入身分
-- 是同一套精神——共用一張圖的話，事後就查不出這份驗收單是誰簽的。
-- 圖片存成 base64 文字，SQLite 與 PostgreSQL 都通吃，不用處理二進位
-- 型別在兩邊不一樣的問題。
CREATE TABLE IF NOT EXISTS user_signatures (
    username    TEXT PRIMARY KEY,
    image_b64   TEXT NOT NULL,
    mime        TEXT DEFAULT 'image/png',
    filename    TEXT DEFAULT '',
    byte_size   INTEGER DEFAULT 0,
    updated_at  TEXT DEFAULT ''
);

-- 簽名批次與歸檔。metadata（誰、何時、簽了哪張 PO、簽成功幾處）永久
-- 保留，簽好的 PDF 本體另外存、可依保留天數清掉——PDF 很佔空間，
-- 資料庫容量有限，但「誰在什麼時候簽了什麼」這筆帳不能跟著被清掉。
CREATE TABLE IF NOT EXISTS sign_batches (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    operator    TEXT DEFAULT '',
    file_count  INTEGER DEFAULT 0,
    signed_count INTEGER DEFAULT 0,
    fail_count  INTEGER DEFAULT 0,
    keyword     TEXT DEFAULT '',
    created_at  TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS signed_docs (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    batch_id     INTEGER NOT NULL,
    operator     TEXT DEFAULT '',
    po_number    TEXT DEFAULT '',
    filename     TEXT DEFAULT '',
    sign_count   INTEGER DEFAULT 0,
    status       TEXT DEFAULT '',
    message      TEXT DEFAULT '',
    pdf_b64      TEXT DEFAULT '',
    byte_size    INTEGER DEFAULT 0,
    purged       INTEGER NOT NULL DEFAULT 0,
    created_at   TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_signdocs_batch ON signed_docs(batch_id);
CREATE INDEX IF NOT EXISTS idx_signdocs_po    ON signed_docs(po_number);
CREATE INDEX IF NOT EXISTS idx_signdocs_time  ON signed_docs(created_at);
"""

# ---------------------------------------------------------------- PostgreSQL schema
#
# 跟 SQLite 版本結構完全對應，只有兩處語法不一樣：
# 1. AUTOINCREMENT 主鍵改用 SERIAL。
# 2. CREATE VIEW 沒有 IF NOT EXISTS 這個寫法，改用 OR REPLACE。

SCHEMA_POSTGRES = """
CREATE TABLE IF NOT EXISTS orders (
    id                SERIAL PRIMARY KEY,

    po_number         TEXT NOT NULL,
    sku_id            TEXT NOT NULL,

    order_type        TEXT DEFAULT '',
    parent_po         TEXT DEFAULT '',
    line              TEXT DEFAULT '',
    brand             TEXT DEFAULT '',
    product_name      TEXT DEFAULT '',
    barcode           TEXT DEFAULT '',
    yf_sku            TEXT DEFAULT '',
    warehouse         TEXT DEFAULT '',
    address           TEXT DEFAULT '',
    delivery_date     TEXT DEFAULT '',
    qty_coupang       INTEGER,
    unit              TEXT DEFAULT '',
    box_size          INTEGER,
    unit_price        REAL,
    expiry_note       TEXT DEFAULT '',
    seq_no            INTEGER,

    qty_ship          INTEGER,
    qty_ship_overridden INTEGER NOT NULL DEFAULT 0,
    delivery_date_overridden INTEGER NOT NULL DEFAULT 0,
    warehouse_overridden     INTEGER NOT NULL DEFAULT 0,
    box_size_overridden      INTEGER NOT NULL DEFAULT 0,
    order_type_overridden    INTEGER NOT NULL DEFAULT 0,
    remarks_legacy    TEXT DEFAULT '',
    receiving_note    TEXT DEFAULT '',

    actual_verified_qty    INTEGER,
    actual_verified_at     TEXT DEFAULT '',

    needs_review      INTEGER NOT NULL DEFAULT 0,
    alert_level       TEXT DEFAULT '',
    review_reason     TEXT DEFAULT '',

    removed_from_coupang INTEGER NOT NULL DEFAULT 0,

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

CREATE TABLE IF NOT EXISTS po_headers (
    po_number        TEXT PRIMARY KEY,
    po_status        TEXT NOT NULL DEFAULT '已建立',
    receiving_status TEXT NOT NULL DEFAULT '未驗收',
    is_pulled        INTEGER NOT NULL DEFAULT 0,
    pulled_at        TEXT DEFAULT '',
    pulled_by        TEXT DEFAULT '',
    pulled_batch_id  INTEGER,
    filed_date       TEXT DEFAULT '',
    shipping_method  TEXT DEFAULT '',
    flagged          INTEGER NOT NULL DEFAULT 0,
    remarks          TEXT DEFAULT '',
    version          INTEGER NOT NULL DEFAULT 1,
    created_at       TEXT DEFAULT '',
    updated_at       TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_poh_status   ON po_headers(po_status);
CREATE INDEX IF NOT EXISTS idx_poh_pulled   ON po_headers(is_pulled);
CREATE INDEX IF NOT EXISTS idx_poh_recv     ON po_headers(receiving_status);

-- view 不在這裡建：既有安裝的 po_headers／orders 可能還沒被
-- _migrate_columns 補上新欄位，這時如果就在這個 schema 腳本裡建 view
-- 引用新欄位，會直接炸在 undefined column（這正是新裝機跟既有安裝
-- 唯一的差異，新裝機的表是空的、剛建就有全部欄位，既有安裝不是）。
-- 統一交給 init_db() 在 _migrate_columns 跑完之後、用 ORDER_ROWS_VIEW
-- 建一次。

CREATE TABLE IF NOT EXISTS edit_logs (
    id          SERIAL PRIMARY KEY,
    order_id    INTEGER NOT NULL,
    po_number   TEXT NOT NULL,
    sku_id      TEXT NOT NULL,
    field       TEXT NOT NULL,
    field_label TEXT NOT NULL,
    old_value   TEXT DEFAULT '',
    new_value   TEXT DEFAULT '',
    operator    TEXT NOT NULL,
    source      TEXT NOT NULL,
    note        TEXT DEFAULT '',
    changed_at  TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_logs_order ON edit_logs(order_id);
CREATE INDEX IF NOT EXISTS idx_logs_po    ON edit_logs(po_number);
CREATE INDEX IF NOT EXISTS idx_logs_field ON edit_logs(field);
CREATE INDEX IF NOT EXISTS idx_logs_time  ON edit_logs(changed_at);

CREATE TABLE IF NOT EXISTS import_batches (
    id             SERIAL PRIMARY KEY,
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

CREATE TABLE IF NOT EXISTS export_batches (
    id          SERIAL PRIMARY KEY,
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

-- 帳號密碼、下拉選單設定（原本的 config.json / users.json /
-- export_profiles.json）改存這裡——Render 每次重新部署都會把程式資料夾
-- 換成全新的，檔案放在裡面就跟訂單資料一樣，改過的東西一次就洗光。
-- 放進資料庫，就跟訂單資料一樣安全。
CREATE TABLE IF NOT EXISTS app_settings (
    key        TEXT PRIMARY KEY,
    value      TEXT NOT NULL,
    updated_at TEXT DEFAULT ''
);

-- 每個人自己的簽名圖：登入誰、蓋的就是誰的簽名，跟修改歷程記登入身分
-- 是同一套精神——共用一張圖的話，事後就查不出這份驗收單是誰簽的。
-- 圖片存成 base64 文字，SQLite 與 PostgreSQL 都通吃，不用處理二進位
-- 型別在兩邊不一樣的問題。
CREATE TABLE IF NOT EXISTS user_signatures (
    username    TEXT PRIMARY KEY,
    image_b64   TEXT NOT NULL,
    mime        TEXT DEFAULT 'image/png',
    filename    TEXT DEFAULT '',
    byte_size   INTEGER DEFAULT 0,
    updated_at  TEXT DEFAULT ''
);

-- 簽名批次與歸檔。metadata（誰、何時、簽了哪張 PO、簽成功幾處）永久
-- 保留，簽好的 PDF 本體另外存、可依保留天數清掉——PDF 很佔空間，
-- 資料庫容量有限，但「誰在什麼時候簽了什麼」這筆帳不能跟著被清掉。
CREATE TABLE IF NOT EXISTS sign_batches (
    id          SERIAL PRIMARY KEY,
    operator    TEXT DEFAULT '',
    file_count  INTEGER DEFAULT 0,
    signed_count INTEGER DEFAULT 0,
    fail_count  INTEGER DEFAULT 0,
    keyword     TEXT DEFAULT '',
    created_at  TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS signed_docs (
    id           SERIAL PRIMARY KEY,
    batch_id     INTEGER NOT NULL,
    operator     TEXT DEFAULT '',
    po_number    TEXT DEFAULT '',
    filename     TEXT DEFAULT '',
    sign_count   INTEGER DEFAULT 0,
    status       TEXT DEFAULT '',
    message      TEXT DEFAULT '',
    pdf_b64      TEXT DEFAULT '',
    byte_size    INTEGER DEFAULT 0,
    purged       INTEGER NOT NULL DEFAULT 0,
    created_at   TEXT DEFAULT ''
);

CREATE INDEX IF NOT EXISTS idx_signdocs_batch ON signed_docs(batch_id);
CREATE INDEX IF NOT EXISTS idx_signdocs_po    ON signed_docs(po_number);
CREATE INDEX IF NOT EXISTS idx_signdocs_time  ON signed_docs(created_at);
"""

SCHEMA = SCHEMA_SQLITE  # 保留舊名字，避免漏改到還在引用它的地方


# ---------------------------------------------------------------- 時間

# 時間一律綁台灣時區，不要用機器的本地時間。
#
# 在自己電腦上跑時兩者剛好一樣，看不出差別；一搬到雲端主機（Render 的
# 機器跑 UTC）就會整整慢 8 小時——下午兩點改的單，修改歷程寫成早上六點。
# 這份歷程是拿去跟酷澎對帳、釐清倉庫責任的證據，時間錯掉等於作廢，
# 所以寫死成 UTC+8，程式擺到哪台機器上記的都是同一個時間。
TAIPEI_TZ = _dt.timezone(_dt.timedelta(hours=8), "Asia/Taipei")


def _local_now():
    return _dt.datetime.now(TAIPEI_TZ)


def now():
    return _local_now().strftime("%Y-%m-%d %H:%M:%S")


def today():
    """今天的日期（台灣），ISO 格式。建檔日這類「日」欄位用這個。"""
    return _local_now().date().isoformat()


def file_stamp():
    """檔名用的時間戳，例如備份檔、匯出檔。"""
    return _local_now().strftime("%Y%m%d_%H%M%S")


# ---------------------------------------------------------------- PostgreSQL 相容層
#
# 呼叫端（app.py／importer.py）全部是照 sqlite3 的用法寫的：
# conn.execute(sql, params)、cur.lastrowid、row["欄位"]、dict(row)。
# 這一層把 psycopg2 包成同樣的用法，這樣不用把整個專案的 SQL 呼叫
# 全部重寫一次，也不會因為漏改某一處而炸掉。

# 這幾張表的 INSERT 會用到 cur.lastrowid 拿新產生的 id，其他表
# （例如 export_batch_items）沒有 id 欄位，硬加 RETURNING id 會直接
# 出錯，所以只白名單有 id 的這幾張。
_LASTROWID_TABLES = ("ORDERS", "IMPORT_BATCHES", "EXPORT_BATCHES",
                     "SIGN_BATCHES", "SIGNED_DOCS")


def _wants_returning_id(sql):
    s = sql.strip().upper()
    if "RETURNING" in s:
        return False
    return any(s.startswith(f"INSERT INTO {t}") for t in _LASTROWID_TABLES)


class _PgCursor:
    def __init__(self, cur):
        self._cur = cur
        self.lastrowid = None

    def execute(self, sql, params=()):
        sql2 = sql.replace("?", "%s")
        add_returning = _wants_returning_id(sql)
        if add_returning:
            sql2 = sql2.rstrip().rstrip(";") + " RETURNING id"
        self._cur.execute(sql2, params)
        if add_returning:
            row = self._cur.fetchone()
            self.lastrowid = row["id"] if row else None
        return self

    def executemany(self, sql, seq):
        self._cur.executemany(sql.replace("?", "%s"), list(seq))
        return self

    def fetchone(self):
        return self._cur.fetchone()

    def fetchall(self):
        return self._cur.fetchall()

    def __iter__(self):
        return iter(self._cur)

    @property
    def rowcount(self):
        return self._cur.rowcount


class _PgConnection:
    """包住 psycopg2 的連線，讓外面呼叫起來跟 sqlite3.Connection 一樣。"""

    def __init__(self, raw):
        self._raw = raw

    def _new_cursor(self):
        return _PgCursor(
            self._raw.cursor(cursor_factory=psycopg2.extras.RealDictCursor))

    def execute(self, sql, params=()):
        return self._new_cursor().execute(sql, params)

    def executemany(self, sql, seq):
        return self._new_cursor().executemany(sql, seq)

    def executescript(self, script):
        # psycopg2 的 cursor.execute 本來就能一次跑一整段用分號分隔的
        # 多條敘述（走 simple query protocol），不需要 sqlite3 才有的
        # executescript 額外處理。
        cur = self._raw.cursor()
        cur.execute(script)
        cur.close()

    def cursor(self):
        return self._new_cursor()

    def commit(self):
        self._raw.commit()

    def rollback(self):
        self._raw.rollback()

    def close(self):
        self._raw.close()


def get_conn():
    if IS_POSTGRES:
        raw = psycopg2.connect(DATABASE_URL)
        return _PgConnection(raw)

    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")      # 多人同時讀寫不互卡
    conn.execute("PRAGMA busy_timeout=8000")     # 瞬間鎖衝突改成等待而非報錯
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def ensure_data_dir():
    """建立資料資料夾、從舊位置搬檔、補上預設設定檔。

    PostgreSQL 模式下設定檔改存資料庫，這裡只需要確保備份資料夾存在；
    SQLite 模式維持原本的搬檔／補檔行為，冪等、不會洗掉使用者改過的
    設定。
    """
    os.makedirs(DATA_DIR, exist_ok=True)
    if IS_POSTGRES:
        return []

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


ORDER_ROWS_VIEW = """
DROP VIEW IF EXISTS order_rows;
CREATE VIEW order_rows AS
SELECT
    o.*,
    h.po_status,
    h.receiving_status,
    h.is_pulled,
    h.pulled_at,
    h.pulled_by,
    h.pulled_batch_id,
    h.filed_date,
    h.shipping_method,
    h.flagged,
    h.remarks,
    h.version AS po_version
FROM orders o
JOIN po_headers h ON h.po_number = o.po_number;
"""

# 保留舊名字，避免漏改到還在引用它的地方。
ORDER_ROWS_VIEW_POSTGRES = ORDER_ROWS_VIEW


def init_db():
    ensure_data_dir()
    conn = get_conn()
    try:
        conn.executescript(SCHEMA_POSTGRES if IS_POSTGRES else SCHEMA_SQLITE)
        _migrate_columns(conn)
        # orders／po_headers 表可能剛剛才被 _migrate_columns 補上新欄位，
        # 上面 executescript 建出來的 view 是舊欄位版本，要重建一次才會
        # 抓到新欄位（不然要等下次重啟才會生效）。SQLite 的
        # CREATE VIEW IF NOT EXISTS 對既有安裝也不會自動套用新欄位，
        # 所以兩種資料庫都要重跑，不是只有 PostgreSQL。
        conn.executescript(ORDER_ROWS_VIEW)
        conn.commit()
    finally:
        conn.close()


def _migrate_columns(conn):
    """CREATE TABLE IF NOT EXISTS 不會幫既有的表補新欄位——資料庫是
    使用者手上舊版跑出來的，每次新增欄位都要在這裡手動補一次
    ALTER TABLE，不然舊資料庫升級後會直接炸在「no such column」
    （PostgreSQL 則是 undefined column）。"""
    if IS_POSTGRES:
        existing = {r["column_name"] for r in conn.execute(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_name = 'orders'")}
    else:
        existing = {r["name"] for r in conn.execute("PRAGMA table_info(orders)")}

    if "box_size_overridden" not in existing:
        conn.execute(
            "ALTER TABLE orders ADD COLUMN box_size_overridden "
            "INTEGER NOT NULL DEFAULT 0")

    if "actual_verified_qty" not in existing:
        conn.execute("ALTER TABLE orders ADD COLUMN actual_verified_qty INTEGER")
    if "actual_verified_at" not in existing:
        conn.execute(
            "ALTER TABLE orders ADD COLUMN actual_verified_at TEXT DEFAULT ''")

    if "order_type_overridden" not in existing:
        conn.execute(
            "ALTER TABLE orders ADD COLUMN order_type_overridden "
            "INTEGER NOT NULL DEFAULT 0")

    if "removed_from_coupang" not in existing:
        conn.execute(
            "ALTER TABLE orders ADD COLUMN removed_from_coupang "
            "INTEGER NOT NULL DEFAULT 0")

    if IS_POSTGRES:
        poh_existing = {r["column_name"] for r in conn.execute(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_name = 'po_headers'")}
    else:
        poh_existing = {r["name"] for r in conn.execute("PRAGMA table_info(po_headers)")}

    if "shipping_method" not in poh_existing:
        conn.execute(
            "ALTER TABLE po_headers ADD COLUMN shipping_method TEXT DEFAULT ''")
    if "flagged" not in poh_existing:
        conn.execute(
            "ALTER TABLE po_headers ADD COLUMN flagged "
            "INTEGER NOT NULL DEFAULT 0")
    if "remarks" not in poh_existing:
        conn.execute("ALTER TABLE po_headers ADD COLUMN remarks TEXT DEFAULT ''")

    # 備註從「每個品項各自一份」搬到「整張單一份」：orders.remarks 改名
    # remarks_legacy，只在既有資料庫、這個欄位還沒改過名時跑一次——
    # 改完名之後 "remarks" 就不會再出現在 existing 裡，這個分支永遠只會
    # 跑這一次，不會每次啟動都重複合併。
    if "remarks" in existing and "remarks_legacy" not in existing:
        conn.execute("ALTER TABLE orders RENAME COLUMN remarks TO remarks_legacy")
        _consolidate_legacy_remarks(conn)


def _consolidate_legacy_remarks(conn):
    """把搬欄位前每個品項各自的備註，合併成整張單一份，不要無聲丟資料。

    同一張單底下可能好幾個品項各自寫了不同的備註，直接只留第一筆會
    默默丟掉其他人寫的內容；這裡改成把所有不重複、非空白的值用「／」
    串起來，一個字都不會不見，事後 OP 自己再去整理成一份就好。"""
    rows = conn.execute(
        """SELECT po_number, remarks_legacy FROM orders
           WHERE remarks_legacy IS NOT NULL AND remarks_legacy != ''
           ORDER BY id""").fetchall()
    combined = {}
    for row in rows:
        po, val = row["po_number"], row["remarks_legacy"]
        seen = combined.setdefault(po, [])
        if val not in seen:
            seen.append(val)
    for po, vals in combined.items():
        conn.execute("UPDATE po_headers SET remarks = ? WHERE po_number = ?",
                     ("／".join(vals), po))


# ---------------------------------------------------------------- 設定值（僅 PostgreSQL 模式）
#
# config.json / users.json / export_profiles.json 在 PostgreSQL 模式下
# 改存進 app_settings 表，不再放檔案——理由跟訂單資料要搬進資料庫一樣：
# Render 每次重新部署都會把程式資料夾換成全新的，檔案放在裡面就跟訂單
# 資料一樣，改過的密碼、加過的人、調過的下拉選單，一次就洗光。

_SETTINGS_DEFAULT_FILE = {
    "config": "config.json",
    "users": "users.json",
    "export_profiles": "export_profiles.json",
}


def load_setting(key, default):
    conn = get_conn()
    try:
        row = conn.execute(
            "SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
        if row is not None:
            return json.loads(row["value"])

        # 第一次啟動，資料庫裡還沒有這個設定：從程式內建的預設檔種一份
        # 進去，之後就都從資料庫讀了。
        fname = _SETTINGS_DEFAULT_FILE.get(key)
        seeded = default
        if fname:
            path = os.path.join(DEFAULTS_DIR, fname)
            try:
                with open(path, encoding="utf-8") as fh:
                    seeded = json.load(fh)
            except (OSError, ValueError):
                seeded = default
        save_setting(key, seeded)
        return seeded
    finally:
        conn.close()


def save_setting(key, value):
    conn = get_conn()
    try:
        payload = json.dumps(value, ensure_ascii=False)
        conn.execute(
            """INSERT INTO app_settings (key, value, updated_at) VALUES (?, ?, ?)
               ON CONFLICT (key) DO UPDATE
                   SET value = EXCLUDED.value, updated_at = EXCLUDED.updated_at""",
            (key, payload, now()))
        conn.commit()
    finally:
        conn.close()


# ---------------------------------------------------------------- 備份

def backup_db(tag="import"):
    """資料庫檔就是全部身家，誤刪等於歸零。每次寫入前先留一份。

    PostgreSQL 模式下資料不是本機檔案，這裡沒有東西好複製——Supabase
    本身有每日備份；本機這份函式在雲端模式下直接跳過，不當一回事地
    假裝成功。"""
    if IS_POSTGRES:
        return ""
    if not os.path.exists(DB_PATH):
        return ""
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = file_stamp()
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
