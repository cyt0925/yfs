"""商品主檔自動化：取代「訂單彙總表 → 專案報價檔 → 總表」那條人工流程。

OP 原本的做法（見 Word 需求文件）：
  1. 從凱特系統下載「訂單彙總表」（就是 OMS 吃的那份整合表）。
  2. 貼進「專案報價檔」，依交貨日一個分頁一個分頁放，不同 PO 用黃底隔開；
     永豐料號／品類／品牌／箱入數／單價 用 VLOOKUP 從總表撈。
  3. 訂單數量、交貨日一改，報價檔就要跟著手改。
  4. 改完重做樞紐：國條 × 出貨數量(箱)，算出每個交貨日的出貨總箱數。
  5. 再 VLOOKUP 回填「總表」對應日期欄位（要貼成值）。

這個模組的三步：
  ① 商品主檔（mst_products）：全公司一份、鍵是國條。從總表匯入品類、品牌、
     永豐料號、箱入數、COGS 單價、Note——就是同事原本 VLOOKUP 總表那幾欄。
     線別不是商品的屬性（總表沒有線別欄），lines_seen 是匯入訂單時自動學來的
     「這個國條出現過哪些線別」。
  ② 訂單明細（mst_orders）：鍵是 (PO, SKU)。線別是每一列自己帶的原始值，同一張
     PO 可以跨線別（CPG-潔品／CPG-紙品）；改交貨日整張 PO 一起搬、不分線別。
     人改過的欄位（出貨數量、交貨日、備註）立旗標，再匯入不覆蓋。
  ③ 總表：不存快照，每次現算。箱數 = 出貨數量 ÷ 箱入數，依交貨日 + 國條加總。

線別顯示名：檔案裡的原始線別（CPG-潔品、CPG-紙品）照實保留在每一列，畫面上
篩選、總表、匯出用顯示名（CPG-* → 紙潔、空白 → 未分類、其他照原名）。規則寫死在
LINE_GROUPS_DEFAULT，不做設定面板。
"""

# 套件拆法（原本是一個 1700 行的 master.py）：
#   common.py          共用的東西：Blueprint、常數、小工具、訂單查詢與篩選。其他模組一律 `from .common import *`。
#   page.py            頁面本身與「有哪些線別／月份」的下拉資料。
#   orders_import.py   ② 訂單明細的匯入：多檔合併解析 → 預覽（比對新增／有變／消失）→ 確認寫入。
#   orders_api.py      ② 訂單明細的查詢與修改：清單＋篩選面、就地編輯、PO 改期、PO 視窗整張存。
#   products.py        ① 商品主檔：清單、手動存、刪除、從酷澎主檔／寶僑總表匯入（含把總表存成匯出樣式）。
#   board.py           ③ 總表看板：一個月一頁，商品層（供需／庫存／金額）與品牌層（目標達成），總表每欄的公式在這裡現算。
#   sources.py         ① 外部來源檔：寶僑 supply 表／Coupang Master／庫存銷售表 → GIV、NIV、每月供需與進銷；庫存現算。
#   summary.py         ③ 總表：現算每日箱數、月配額、匯出總表（有底稿就照底稿填、缺的日期自動插欄）。
#   quotation.py       匯出專案報價檔（一個交貨日一個分頁，A～R 欄）。
#   admin_api.py       清除資料、匯入歷程（各次匯入新增／有變）、修改歷程查詢。
#   stats.py           ④ 改單統計：每月幾張 PO、幾張被改過、改幾次、原因分佈、匯出 Excel（給主管算成本）。
# 路由都掛在 common.master_bp 上，這裡把各模組 import 進來讓路由生效，並保留 `master.master_bp` 給 app.py 用。
from .common import *  # noqa: F401,F403
from .page import *  # noqa: F401,F403
from .orders_import import *  # noqa: F401,F403
from .orders_api import *  # noqa: F401,F403
from .products import *  # noqa: F401,F403
from .sources import *  # noqa: F401,F403
from .board import *  # noqa: F401,F403
from .summary import *  # noqa: F401,F403
from .quotation import *  # noqa: F401,F403
from .admin_api import *  # noqa: F401,F403
from .stats import *  # noqa: F401,F403

# 給測試與腳本用的內部名稱（保持跟舊的 master.py 一樣可以 master._xxx 拿到）
from .page import master_page, api_lines  # noqa: F401
from .orders_import import _diff_import, _parse_uploads, api_import_preview, api_import_commit  # noqa: F401
from .orders_api import api_orders, _apply_item_edit, api_update_order, api_delete_order, _move_po_dates, api_update_po_date, api_po_detail, api_po_save  # noqa: F401
from .products import api_products, _upsert_product, api_save_product, api_delete_product, _HEADER_ALIASES, _PRODUCT_TEXT_FIELDS, _find_columns, api_import_products, _DATE_HDR, _template_row, _guess_line, _save_template  # noqa: F401
from .summary import api_set_quota, _build_summary, api_summary, _template_meta, api_template_get, _REF_PART, _shift_ref_text, _shift_formula, _insert_column, _fill_template, api_export  # noqa: F401
from .quotation import api_export_daily  # noqa: F401
from .admin_api import api_reset, api_imports, api_logs  # noqa: F401
from .stats import api_stats, api_stats_events, api_stats_export, _build_stats  # noqa: F401
