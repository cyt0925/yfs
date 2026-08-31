"""採購表轉換：跟酷澎訂單管理系統完全獨立的新功能。

同一份酷澎「訂單匯入」Excel，依線別（PG／紙潔／瑪氏）轉換成各線別
自己要用的「產品採購表」格式。三個線別各自沿用公司既有的範本檔
（purchase_templates/ 底下那三個 .xls，連查表工作表、儲存格樣式都
原封不動保留），只覆寫資料列——這樣匯出的檔案結構才會跟公司系統
原本要吃的格式一模一樣，不會因為改用別的函式庫重新產生而跑掉。

不寫進資料庫、不碰任何一張酷澎訂單的表：上傳的檔案只在這次請求內
處理、解析結果全部回傳給瀏覽器暫存，使用者確認/編輯完再送回來產生
下載檔，處理完就丟，跟訂單管理系統的資料完全隔離。
"""
import io
import json
import os
import re
import zipfile

import openpyxl
import xlrd
from flask import Blueprint, jsonify, render_template, request, send_file
from xlutils.copy import copy as xl_copy

import db

purchase_bp = Blueprint("purchase", __name__)

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "purchase_templates")

# wipe_rows：範本檔主表原本的資料區域大概到第幾列（含預先編號的空白列），
# 產生新檔案時全部清成空白再重寫，兩份範例檔都留了不少預先編號、單位
# 已經填好「箱」的空白列在後面，這裡抓大一點，寧可多清幾列不要漏。
LINES = {
    "pg": {
        "label": "P&G",
        "template": os.path.join(TEMPLATE_DIR, "pg_template.xls"),
        "group_by": "po_number",
        "wipe_rows": 250,
        "default_remark_style": "buyout",
    },
    "paper": {
        "label": "紙潔",
        "template": os.path.join(TEMPLATE_DIR, "paper_template.xls"),
        "group_by": "po_number",
        "wipe_rows": 45,
        "default_remark_style": "blank",
    },
    "mars": {
        "label": "瑪氏",
        "template": os.path.join(TEMPLATE_DIR, "mars_template.xls"),
        "group_by": "ship_note",
        "wipe_rows": 10,
        "default_remark_style": "mars",
    },
}

# 欄位名稱比對用——酷澎匯出的「訂單匯入」檔在 P&G／紙潔／瑪氏三條線
# 欄位順序不一樣（瑪氏把出貨備註搬到最前面），一律用標題文字找欄位，
# 不能寫死欄位字母位置。
FIELD_HEADERS = {
    "po_number": ["訂單編號"],
    "material_no": ["料號"],
    "qty": ["數量"],          # 前綴比對，因為標題後面偶爾會黏著總和數字
    "ship_note": ["出貨備註"],
    "address": ["收件地址"],
}

WAREHOUSE_FILE = "purchase_warehouses.json"
DEFAULT_WAREHOUSES = {
    "桃園市楊梅區環東路200號C區4樓": "TXRC8",
    "桃園市大園區建國路102號4樓 B棟(倉儲棟)": "TAO1",
    "桃園市觀音區玉林路一段523號": "TAO4",
}


def _load_warehouses():
    path = os.path.join(db.DATA_DIR, WAREHOUSE_FILE)
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
            if isinstance(data, dict) and data:
                return data
    except (OSError, ValueError):
        pass
    return dict(DEFAULT_WAREHOUSES)


def _save_warehouses(mapping):
    os.makedirs(db.DATA_DIR, exist_ok=True)
    path = os.path.join(db.DATA_DIR, WAREHOUSE_FILE)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(mapping, fh, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def _find_column(headers, candidates):
    for i, h in enumerate(headers):
        if h is None:
            continue
        text = str(h).strip()
        for cand in candidates:
            if text == cand or text.startswith(cand):
                return i
    return None


def _guess_date_mmdd(ship_note):
    """從出貨備註開頭抓 MMDD（例如「0905從PG出貨至TXRC8」→ "0905"）。"""
    if not ship_note:
        return ""
    m = re.match(r"^(\d{2})(\d{2})", ship_note.strip())
    if not m:
        return ""
    month, day = int(m.group(1)), int(m.group(2))
    if 1 <= month <= 12 and 1 <= day <= 31:
        return m.group(0)
    return ""


def _guess_date_from_filename(filename):
    """出貨備註猜不到日期時（瑪氏就是這樣，出貨備註是內部採購單號，
    不含日期）退而看上傳檔名——三個線別的匯入檔名都遵守同一個慣例，
    帶著「MMDD到貨」，例如「酷澎訂單匯入_0904到貨_TAO1_TAO4.xlsx」。"""
    if not filename:
        return ""
    m = re.search(r"(\d{2})(\d{2})到貨", filename)
    if not m:
        return ""
    month, day = int(m.group(1)), int(m.group(2))
    if 1 <= month <= 12 and 1 <= day <= 31:
        return m.group(1) + m.group(2)
    return ""


def mmdd_to_md(mmdd):
    """0905 -> 9/5（備註文字慣用不補零的月/日）。"""
    if not mmdd or len(mmdd) != 4 or not mmdd.isdigit():
        return ""
    return f"{int(mmdd[:2])}/{int(mmdd[2:])}"


def _guess_warehouse_from_note(ship_note):
    """PG 的出貨備註格式是「MMDD從PG出貨至TXRC8」，倉別直接寫在裡面；
    其他線別的出貨備註格式不含倉別，抓不到就回傳空字串，交給地址對照表。"""
    if not ship_note:
        return ""
    m = re.search(r"至([A-Z0-9]+)\s*$", ship_note.strip())
    return m.group(1) if m else ""


class PurchaseImportError(Exception):
    pass


def parse_import_file(line_key, file_bytes, filename=""):
    """讀酷澎的「訂單匯入」Excel，依線別規則分組，回傳給前端預覽用的結構。

    filename 是上傳時的原始檔名，只用來在出貨備註猜不到日期時當備援
    （見 _guess_date_from_filename）——不影響解析內容，純粹輔助猜日期。"""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb[wb.sheetnames[0]]
    except Exception as exc:
        raise PurchaseImportError(f"檔案讀取失敗，請確認是有效的 Excel 檔：{exc}") from exc

    header_row = next(ws.iter_rows(min_row=1, max_row=1), None)
    if header_row is None:
        raise PurchaseImportError("檔案是空的，沒有表頭列。")
    headers = [c.value for c in header_row]

    cols = {}
    missing = []
    for field, candidates in FIELD_HEADERS.items():
        idx = _find_column(headers, candidates)
        if idx is None:
            missing.append(candidates[0])
        cols[field] = idx
    if missing:
        raise PurchaseImportError(
            f"檔案裡找不到欄位：{'、'.join(missing)}，請確認上傳的是酷澎「訂單匯入」檔。")

    warehouses = _load_warehouses()
    line_cfg = LINES[line_key]

    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        po = r[cols["po_number"]]
        material = r[cols["material_no"]]
        if po is None or material is None or str(po).strip() == "":
            continue
        po = str(po).strip()
        material = str(material).strip()
        raw_qty = r[cols["qty"]]
        try:
            qty = int(raw_qty)
        except (TypeError, ValueError):
            qty = 0
        ship_note = str(r[cols["ship_note"]] or "").strip() if cols["ship_note"] is not None else ""
        address = str(r[cols["address"]] or "").strip() if cols["address"] is not None else ""
        rows.append({
            "po_number": po, "material_no": material, "qty": qty,
            "ship_note": ship_note, "address": address,
        })

    if not rows:
        raise PurchaseImportError("這個檔案解析不到任何資料列，請確認上傳的是酷澎「訂單匯入」檔。")

    group_field = line_cfg["group_by"]
    groups_order = []
    groups = {}
    for row in rows:
        key = row[group_field] or row["po_number"]
        if key not in groups:
            groups[key] = []
            groups_order.append(key)
        groups[key].append(row)

    result = []
    for key in groups_order:
        grows = groups[key]
        po_numbers = sorted({r["po_number"] for r in grows})
        addresses = {r["address"] for r in grows if r["address"]}
        address = next(iter(addresses)) if len(addresses) == 1 else ""

        warehouse_guess = warehouses.get(address, "") if address else ""
        if not warehouse_guess:
            for r in grows:
                warehouse_guess = _guess_warehouse_from_note(r["ship_note"])
                if warehouse_guess:
                    break

        date_guess = ""
        for r in grows:
            date_guess = _guess_date_mmdd(r["ship_note"])
            if date_guess:
                break
        if not date_guess:
            # 瑪氏的出貨備註就是內部採購單號，不含日期，退而看上傳檔名——
            # 三個線別的匯入檔名都遵守同一個「MMDD到貨」慣例。
            date_guess = _guess_date_from_filename(filename)

        result.append({
            "key": str(key),
            "po_numbers": po_numbers,
            "address": address,
            "address_mismatch": len(addresses) > 1,
            "warehouse_guess": warehouse_guess,
            "date_guess": date_guess,
            "item_count": len(grows),
            "qty_total": sum(r["qty"] for r in grows),
            "rows": [{"material_no": r["material_no"], "qty": r["qty"]} for r in grows],
        })
    return result


def build_remark(line_key, po_numbers, date_mmdd, warehouse):
    """三個線別各自的備註樣板。日期／倉別任一沒填就先留空，讓使用者自己補，
    不要硬湊出一句看起來對、其實缺資料的備註。"""
    style = LINES[line_key]["default_remark_style"]
    md = mmdd_to_md(date_mmdd)
    if style == "blank":
        return ""
    if style == "buyout":
        po = po_numbers[0] if po_numbers else ""
        if not (po and warehouse and md):
            return ""
        return f"{po}({warehouse})_酷澎{md}買斷"
    if style == "mars":
        if not (md and warehouse):
            return ""
        return f"MARS入倉{md}瑪氏送酷澎-{warehouse}倉"
    return ""


def build_filename(line_key, po_numbers, date_mmdd, warehouse):
    """檔名裡的 PO 單號是用來區分「一張 PO 一個檔」時是哪一張，所以合併
    匯出時呼叫端會傳空的 po_numbers 進來，這裡就不放單號——這正是紙潔
    （本來就是多張 PO 合併）的真實檔名慣例：酷澎_產品採購表上傳_0902到貨。"""
    po = po_numbers[0] if po_numbers else ""
    if line_key == "pg":
        # 只取後 6 碼——訂單編號前面那一串每張都一樣，沒有辨識度，
        # 檔名太長也不好看。後 6 碼理論上有機會撞號（機率很低，同一批
        # 匯出裡撞到才會真的出事），撞了也不會互相覆蓋：多檔匯出時
        # 有另一層檔名去重機制兜底（見 api_purchase_export）。
        short_po = po[-6:] if po else ""
        suffix = f"({short_po})" if short_po else ""
        return f"酷澎XP&G_產品採購表上傳_{date_mmdd}到貨{suffix}.xls"
    if line_key == "paper":
        return f"酷澎_產品採購表上傳_{date_mmdd}到貨.xls"
    if line_key == "mars":
        wh = warehouse or "倉別"
        suffix = f"_{po}" if po else ""
        return f"永豐Mars採購單(箱單位)-GUM_{wh}{suffix}.xls"
    return f"採購表_{date_mmdd}.xls"


def fill_template(line_key, rows, remark):
    """載入該線別的公司範本檔，清空資料區、寫入新資料，回傳檔案位元組。

    用 xlutils.copy 而不是重新用 xlwt 從零產生，是為了讓查表工作表
    （表單選項／商品資料／客編）跟儲存格樣式完全保留，降低跟公司系統
    要求的格式對不起來的風險。

    備註是「逐列」的：合併多張 PO 時，每一列要保留自己那張 PO 的備註
    （P&G／瑪氏的備註帶著訂單編號），所以每列各自帶 remark 就用自己的，
    沒帶才退回用整份共用的那個。之前整份共用一個備註，合併後會把所有
    列都寫成第一張 PO 的單號，後面幾張的單號整個消失。"""
    cfg = LINES[line_key]
    rb = xlrd.open_workbook(cfg["template"], formatting_info=True)
    wb = xl_copy(rb)
    ws = wb.get_sheet(0)

    wipe_upto = max(cfg["wipe_rows"], len(rows) + 1)
    for r in range(1, wipe_upto):
        for c in range(10):
            ws.write(r, c, "")

    for i, row in enumerate(rows, start=1):
        ws.write(i, 0, i)                        # 項目
        ws.write(i, 1, str(row["material_no"]))   # 料號
        ws.write(i, 2, "")                         # 品名
        ws.write(i, 3, "箱")                        # 單位
        ws.write(i, 6, int(row["qty"]))            # 採購數量
        ws.write(i, 8, row.get("remark", remark))   # 備註（逐列，見上面說明）

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------- 路由

@purchase_bp.route("/purchase")
def purchase_page():
    # 這頁的識別色是綠色，logo 也另外用一張綠色版。檔案還沒放進去之前
    # 先用原本那張，等 static/logo_purchase.png 一丟進去就自動換掉，
    # 不用再改程式；萬一哪天被刪掉也只是變回原本的 logo，不會破圖。
    logo = ("logo_purchase.png"
            if os.path.exists(os.path.join(os.path.dirname(__file__),
                                            "static", "logo_purchase.png"))
            else "logo.png")
    return render_template("purchase.html", lines=LINES, logo_file=logo)


@purchase_bp.route("/api/purchase/lines")
def api_purchase_lines():
    return jsonify([{"key": k, "label": v["label"]} for k, v in LINES.items()])


@purchase_bp.route("/api/purchase/warehouses")
def api_purchase_warehouses():
    return jsonify(_load_warehouses())


@purchase_bp.route("/api/purchase/parse", methods=["POST"])
def api_purchase_parse():
    line_key = request.form.get("line", "")
    if line_key not in LINES:
        return jsonify({"error": "請選擇正確的線別。"}), 400
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    try:
        groups = parse_import_file(line_key, upload.read(), upload.filename)
    except PurchaseImportError as exc:
        return jsonify({"error": str(exc)}), 400

    for g in groups:
        g["remark"] = build_remark(line_key, g["po_numbers"], g["date_guess"], g["warehouse_guess"])
        g["filename"] = build_filename(line_key, g["po_numbers"], g["date_guess"], g["warehouse_guess"])

    return jsonify({"line": line_key, "groups": groups})


@purchase_bp.route("/api/purchase/export", methods=["POST"])
def api_purchase_export():
    payload = request.get_json(silent=True) or {}
    line_key = payload.get("line", "")
    if line_key not in LINES:
        return jsonify({"error": "請選擇正確的線別。"}), 400
    groups = payload.get("groups") or []
    if not isinstance(groups, list) or not groups:
        return jsonify({"error": "沒有要匯出的資料。"}), 400

    files = []
    for g in groups:
        rows = g.get("rows") or []
        if not rows:
            continue
        filename = (g.get("filename") or "").strip() or "採購表.xls"
        if not filename.lower().endswith(".xls"):
            filename += ".xls"
        remark = g.get("remark") or ""
        data = fill_template(line_key, rows, remark)
        files.append((filename, data))

    if not files:
        return jsonify({"error": "選取的項目裡沒有任何資料列。"}), 400

    if len(files) == 1:
        name, data = files[0]
        return send_file(io.BytesIO(data), as_attachment=True, download_name=name,
                          mimetype="application/vnd.ms-excel")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = set()
        for name, data in files:
            final = name
            n = 2
            while final in used_names:
                base, ext = os.path.splitext(name)
                final = f"{base}({n}){ext}"
                n += 1
            used_names.add(final)
            zf.writestr(final, data)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="採購表.zip",
                      mimetype="application/zip")


@purchase_bp.route("/api/purchase/warehouses", methods=["POST"])
def api_purchase_warehouses_save():
    payload = request.get_json(silent=True) or {}
    address = (payload.get("address") or "").strip()
    code = (payload.get("code") or "").strip()
    if not address or not code:
        return jsonify({"error": "地址跟倉別代碼都要填。"}), 400
    mapping = _load_warehouses()
    mapping[address] = code
    _save_warehouses(mapping)
    return jsonify({"ok": True, "warehouses": mapping})
