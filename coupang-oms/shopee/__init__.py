"""蝦皮特選寄倉訂單系統：照「蝦皮特選寄倉訂單系統需求初稿」（Jerry 2026-10-07 給的 Word）做，網址 /shopee。

第一段（2026-10-07）做到：
  parse.py   讀蝦皮採購單（PurchaseOrder_…）、入庫單（InboundOrder_…），認線別、單位代碼
  orders.py  上傳 → 預覽 → 選履約方式確認匯入；重傳比對差異、保留每一版、差異確認；訂單總覽；改履約方式；上傳紀錄與原始檔下載
第二段（2026-10-08）做到：
  purchase.py 價格本（轉換率）上傳；竹運採購進貨、供應商直送 → EIP 採購單＋採購下採明細（寶僑、紙潔）；
              除不盡轉換率的人工修改；已產出的 EIP 檔、回填 EIP 單號
後面幾段（瑪氏直送、竹運出貨效期判斷、缺貨下修）還沒做。
判讀依據、Jerry 定的規則、還沒問清楚的事寫在 docs/蝦皮特選_設計筆記.md。
"""
from flask import Blueprint

shopee_bp = Blueprint("shopee", __name__)

from . import orders, purchase  # noqa: E402,F401 — 讓路由生效

__all__ = ["shopee_bp"]
