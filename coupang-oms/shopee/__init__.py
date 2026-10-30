"""蝦皮特選寄倉訂單系統：照「蝦皮特選寄倉訂單系統需求初稿」（Jerry 2026-10-07 給的 Word）做，網址 /shopee。

第一段（2026-10-07）做到：
  parse.py   讀蝦皮採購單（PurchaseOrder_…）、入庫單（InboundOrder_…），認線別、單位代碼
  orders.py  上傳 → 預覽 → 選履約方式確認匯入；重傳比對差異、保留每一版、差異確認；訂單總覽；改履約方式；上傳紀錄與原始檔下載
第二段（2026-10-08）做到：
  purchase.py 價格本（轉換率）上傳；竹運採購進貨、供應商直送 → EIP 採購單＋採購下採明細（寶僑、紙潔）；
              除不盡轉換率的人工修改；已產出的 EIP 檔、回填 EIP 單號
瑪氏直送（2026-10-12）：
  purchase.py 瑪氏的 EIP 採購單（單位數量照蝦皮）；mars_po.py 回填單號後產瑪氏制式採購單、蝦皮倉資料
  direct_po.py 寶僑、紙潔供應商直送：回填單號後產給採購轉原廠的直送採購單（Alice 2026-10-21 的範本）
竹運出貨（2026-10-12）：
  ship.py     EC 即時庫存、竹運庫存效期表、竹運商品主檔上傳；照 Alice 的「核對竹運出貨效期及庫存 V5」判斷可出貨量、
              選效期批次、人工改出貨量；產訂單系統拋單檔＋下修刪單清單
缺貨維護與通知（2026-10-12）：
  short.py    寶僑（兩次）、紙潔（一次）人工填缺貨；竹運出貨的缺貨自動帶入；產 LINE 通知文字＋Excel 明細，通知過的那次鎖住
瑪氏缺貨（2026-10-13）：
  yx.py       讀勇信配送明細表 PDF（跟酷澎瑪氏同一份格式），對回瑪氏 EIP 逐品比箱數，確認後寫成缺貨紀錄，再走 short.py 的通知
畫面改版（2026-10-13）：
  drop.py     一個上傳入口收所有檔案（看內容認是哪一種，再交給原本那一支）
  orders.py   overview() 算每張單的待辦（作業流程格子）
  side.py     右側面板：一張單的全部（品項、採購單計畫、竹運出貨、缺貨、EIP、差異、歷程）
刪除訂單、匯出明細（2026-10-30，Alice 要的，Jerry 看過示意圖）：
  trash.py    刪除／還原一張訂單（shp_deleted，不動 shp_orders；各步用 not_deleted() 跳過，重傳不會回來）
  po_pdf.py   蝦皮採購單 PDF 畫缺貨（不出貨整列畫線；短出劃掉數量、寫實際、蓋章；印章跟 MOMO 共用）
  items.py    清單上展開品項、直接改出貨量（一次拿好幾張單的品項；竹運出貨只給數字，效期自動挑）
  export.py   勾選訂單匯出訂單明細 Excel（一張工作表、一品一列、缺貨整列黃底；資料跟訂單視窗同一份 side_data）
判讀依據、Jerry 定的規則、還沒問清楚的事寫在 docs/蝦皮特選_設計筆記.md。
"""
from flask import Blueprint

shopee_bp = Blueprint("shopee", __name__)

from . import admin, direct_po, drop, export, items, mars_po, orders, po_pdf, purchase, refulfil, ship, short, side, trash, yx  # noqa: E402,F401 — 讓路由生效

__all__ = ["shopee_bp"]
