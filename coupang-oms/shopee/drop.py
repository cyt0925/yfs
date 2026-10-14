"""一個入口收所有檔案（Jerry 2026-10-13 改版：「真的要分為六個介面嗎」）。

蝦皮採購單／入庫單、價格本、EC 即時庫存／竹運庫存效期表／竹運商品主檔、勇信配送明細表 PDF 都丟同一個地方，
這裡看內容認是哪一種，再交給原本那一支處理；回應的 X-Shopee-Kind 告訴畫面接下來要做什麼
（蝦皮訂單要選履約方式、勇信要看比對結果，其他的直接存好）。

認的順序：PDF → 蝦皮採購單／入庫單（PR ID、Shopee SKU ID…）→ 瑪氏商品總表（瑪氏貨號＋永豐料號＋單位類別）→ 價格本（料號3＋箱入數）→ 庫存資料。
瑪氏商品總表跟瑪氏出貨頁共用同一份（Jerry 2026-10-14：「是否也可以在蝦皮這邊上傳」），這裡傳就是呼叫那邊同一支匯入。
價格本排在庫存前面：PG 價格本整合裡有一頁「庫存資料更新」，表頭跟 EC 即時庫存很像，先認庫存會把價格本當成 EC 庫存存掉。
"""
from flask import make_response

from mars.products import api_import_mars_products, parse_master

from .common import *  # noqa: F401,F403
from .parse import ParseError, parse_file
from .purchase import api_conv_import, parse_conv
from .ship import api_shopee_stock_import, parse_stock
from .orders import api_import_preview
from .yx import api_yx_compare

KIND_TEXT = {"orders": "蝦皮訂單", "mars": "瑪氏商品總表", "conv": "價格本", "stock": "庫存資料", "yx": "勇信配送明細表"}


def detect(data, filename):
    """回傳 orders／mars／conv／stock／yx；認不出來回空字串。"""
    if data[:5] == b"%PDF-":
        return "yx"
    try:
        if parse_file(data, filename)["rows"]:
            return "orders"
    except ParseError:
        pass
    except Exception:  # noqa: BLE001 — 開不起來的檔交給後面幾種再試
        pass
    try:
        if parse_master(io.BytesIO(data))[0]:
            return "mars"
    except ValueError:
        pass
    try:
        rows, _w = parse_conv(data, filename)
        if rows:
            return "conv"
    except ValueError:
        pass
    try:
        parse_stock(data, filename)
        return "stock"
    except ValueError:
        pass
    return ""


@shopee_bp.route("/api/shopee/drop", methods=["POST"])
def api_drop():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify({"error": "未收到檔案，請重新上傳。"}), 400
    data = f.read()
    kind = detect(data, f.filename)
    if not kind:
        return jsonify({"error": f"{f.filename}：認不出是哪一種檔案。可以上傳蝦皮採購單／入庫單、價格本、瑪氏商品總表、EC 即時庫存、"
                                 "竹運庫存效期表、竹運商品主檔、勇信配送明細表（PDF）。"}), 400
    f.stream.seek(0)
    handler = {"orders": api_import_preview, "mars": api_import_mars_products, "conv": api_conv_import, "stock": api_shopee_stock_import,
               "yx": api_yx_compare}[kind]
    resp = make_response(handler())
    resp.headers["X-Shopee-Kind"] = kind
    return resp


__all__ = ["detect", "api_drop", "KIND_TEXT"]
