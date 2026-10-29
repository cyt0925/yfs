"""MOMO 訂單管理系統：照「MOMO 訂單系統需求書」（Jerry 2026-10-27 給的 Word）做，網址 /momo。

第一批（2026-10-28，Jerry 看過示意圖「好 開始做」）：
  parse.py    讀 MOMO 後台下載的未出貨清單（UnsendList），表頭用文字認；單據首列往下帶；單位代碼；看料號分紙品、潔品、瑪氏
  orders.py   一次丟好幾個檔 → 合在一起預覽 → 確認匯入；重傳比對差異；訂單總覽（作業流程格子）；出貨方式；上傳紀錄
  purchase.py 換成箱數（價格本、瑪氏商品總表跟蝦皮特選共用）、除不盡改數量；EIP 採購單；回填 EIP 單號；瑪氏採購單；直送採購單
  admin.py    清除資料（限管理員，2026-10-29）
  notice.py   ④ 待畫入庫通知單、嘜頭填箱數：MOMO 入庫通知單 PDF 照列畫線、畫掉數量寫實際數、蓋業務專用章（2026-10-29，Becky 的畫法）
  check.py    ③ 待核缺貨與效期：勇信配送明細表核缺貨、效期；填 MOMO 效期匯入檔（第二批，2026-10-28）
  side.py     右側面板：一張入庫單的全部
  settings.py 倉別主檔（需求第九章）、上傳入口認檔
判讀依據、預設的做法、還沒問清楚的事寫在 docs/MOMO_設計筆記.md。
"""
from flask import Blueprint

momo_bp = Blueprint("momo", __name__)

from . import admin, check, notice, orders, purchase, settings, side  # noqa: E402,F401 — 讓路由生效

__all__ = ["momo_bp"]
