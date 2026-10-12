"""蝦皮特選寄倉訂單系統：照「蝦皮特選寄倉訂單系統需求初稿」（Jerry 2026-10-07 給的 Word）做，網址 /shopee。

第一段（2026-10-07）做到：
  parse.py   讀蝦皮採購單（PurchaseOrder_…）、入庫單（InboundOrder_…），認線別、單位代碼
  orders.py  上傳 → 預覽 → 選履約方式確認匯入；重傳比對差異、保留每一版、差異確認；訂單總覽；改履約方式；上傳紀錄與原始檔下載
第二段（2026-10-08）做到：
  purchase.py 價格本（轉換率）上傳；竹運採購進貨、供應商直送 → EIP 採購單＋採購下採明細（寶僑、紙潔）；
              除不盡轉換率的人工修改；已產出的 EIP 檔、回填 EIP 單號
瑪氏直送（2026-10-12）：
  purchase.py 瑪氏的 EIP 採購單（單位數量照蝦皮）；mars_po.py 回填單號後產瑪氏制式採購單、蝦皮倉資料
竹運出貨（2026-10-12）：
  ship.py     EC 即時庫存、竹運庫存效期表、竹運商品主檔上傳；照 Alice 的「核對竹運出貨效期及庫存 V5」判斷可出貨量、
              選效期批次、人工改出貨量；產訂單系統拋單檔＋下修刪單清單
缺貨維護與通知（2026-10-12）：
  short.py    寶僑（兩次）、紙潔（一次）人工填缺貨；竹運出貨的缺貨自動帶入；產 LINE 通知文字＋Excel 明細，通知過的那次鎖住
瑪氏用勇信效期表自動核對缺貨還沒做。
判讀依據、Jerry 定的規則、還沒問清楚的事寫在 docs/蝦皮特選_設計筆記.md。
"""
from flask import Blueprint

shopee_bp = Blueprint("shopee", __name__)

from . import mars_po, orders, purchase, ship, short  # noqa: E402,F401 — 讓路由生效

__all__ = ["shopee_bp"]
