"""瑪氏出貨：Alice 的「瑪氏酷澎出貨優化」規格（出貨彙總表 → 拆單 → EIP 採購單 → 瑪氏採購單 → EMMA → 勇信缺貨）。

入口在上方「線別工具」的瑪氏那欄，網址 /mars。跟寶僑的商品主檔自動化是兩個工具，但**訂單共用**：
拆單直接拿 ② 訂單明細裡已經匯進來的瑪氏訂單（mst_orders，線別歸到「瑪氏」的），不用再傳一次出貨彙總表，
人在 ② 改過的出貨數量也照改過的算。瑪氏商品總表另外存（mst_mars_products），不跟寶僑主檔混。

目前做到：
  products.py   瑪氏商品總表上傳（整份覆蓋）
  split.py      ① 拆單（品類／單位／中標／到貨日／到貨倉）→ 下載拆單表＋EIP 上傳用採購表
                ② 回填 EIP 採購單號、約倉時間（一份拆單表一張 EIP 採購單）
之後：③ 瑪氏採購單（V2 範本）④ EMMA 匯入檔 ⑤ 勇信效期表缺貨比對。
判讀的依據與還沒問清楚的事，寫在 docs/瑪氏出貨_設計筆記.md。
"""
from flask import Blueprint

mars_bp = Blueprint("mars", __name__)

from . import products, split  # noqa: E402,F401 — 讓路由生效

__all__ = ["mars_bp"]
