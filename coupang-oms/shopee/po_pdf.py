"""蝦皮採購單 PDF 畫缺貨（Alice 2026-10-30「蝦皮特選也需要同 MOMO 畫除缺貨的項目」；Jerry：短出要畫線、填上正確數量、要蓋章）。

上傳蝦皮後台下載的採購單 PDF（PURCHASE ORDER，Chromium 產的、有文字）→ 系統用 Supplier SKU ID 對回這張單的品項，
照缺貨紀錄（勇信核對、手填缺貨、竹運出貨的出貨量，跟訂單視窗同一份 side_data）畫：
  不出貨：整列一條紅線（照昱靜 2026-10-30 畫的範本，不蓋章）
  短出：Purchase QTY (unit)、(pcs) 兩格的數字劃掉、寫實際數量，Unit Name 那格蓋業務專用章
  最下面的總數不改（範本沒改）。
自己用「Microsoft Print to PDF」印出來的 PDF 字都變成圖形、讀不到，要用後台下載的原檔。
印章跟 MOMO 共用同一顆（mst_meta momo_stamp，在 MOMO「資料與設定」上傳）；還沒上傳就畫紅框寫「業務專用章」。
"""
import pymupdf

from momo.notice import _stamp_rect, load_stamp

from .common import *  # noqa: F401,F403
from .orders import overview

RED = (0.85, 0.1, 0.1)
COLS = {"no": "No.", "sku": "SupplierSKUID", "qu": "PurchaseQTY(unit)", "unit": "UnitName", "qp": "PurchaseQTY(pcs)"}


def _num(s):
    try:
        return float(str(s).replace(",", ""))
    except ValueError:
        return None


def is_po_pdf(data):
    if data[:5] != b"%PDF-":
        return False
    try:
        t = pymupdf.open(stream=data, filetype="pdf")[0].get_text()
    except Exception:  # noqa: BLE001
        return False
    return "PURCHASE ORDER" in t and "Purchase Order ID" in t


def has_text(data):
    try:
        return bool(pymupdf.open(stream=data, filetype="pdf")[0].get_text("words"))
    except Exception:  # noqa: BLE001
        return False


def _lines(page):
    """表格線：橫線 {y: 總長}、直線 {x: 總長}（Chromium 畫的是細長方形，也有畫線的）。"""
    hy, vx = collections.defaultdict(float), collections.defaultdict(float)
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "re":
                r = it[1]
                if r.height < 1.6 and r.width > 3:
                    hy[round(r.y0, 1)] += r.width
                elif r.width < 1.6 and r.height > 3:
                    vx[round(r.x0, 1)] += r.height
            elif it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.y - b.y) < 0.5:
                    hy[round(a.y, 1)] += abs(a.x - b.x)
                elif abs(a.x - b.x) < 0.5:
                    vx[round(a.x, 1)] += abs(a.y - b.y)

    def merge(d, need):
        out = []
        for k in sorted(k for k, v in d.items() if v > need):
            if out and k - out[-1] < 2:
                continue
            out.append(k)
        return out
    return merge(hy, 300), merge(vx, 60)


def parse_po_pdf(data, filename=""):
    """→ {"po_id", "rows": [{no, sku, qty_unit, qty_pcs, page, band, x0, x1, qu, qp, unit}]}；qu／qp 是那格的欄寬和數字的框。"""
    doc = pymupdf.open(stream=data, filetype="pdf")
    first = doc[0].get_text("words")
    po = next((w[4] for w in first if re.fullmatch(r"PO[A-Z0-9]{8,}", w[4])), "")
    if not po:
        raise ValueError(f"{filename}：讀不到 Purchase Order ID（PO 開頭的單號）。")
    rows = []
    for pi, page in enumerate(doc):
        ws = page.get_text("words")
        head = next((w for w in ws if w[4] == "No."), None)
        if head is None:
            continue
        ys, xs = _lines(page)
        cy = (head[1] + head[3]) / 2
        above, below = [y for y in ys if y < cy], [y for y in ys if y > cy]
        if not above or not below or len(xs) < 6:
            continue
        h0, h1 = above[-1], below[0]
        within = lambda x0, x1, y0, y1: sorted([w for w in ws if y0 < (w[1] + w[3]) / 2 < y1 and x0 <= (w[0] + w[2]) / 2 < x1], key=lambda w: (round(w[1]), w[0]))
        col = {}
        for x0, x1 in zip(xs, xs[1:]):
            label = "".join(w[4] for w in within(x0, x1, h0, h1))
            for k, v in COLS.items():
                if label == v:
                    col[k] = (x0, x1)
        if len(col) < len(COLS):
            raise ValueError(f"{filename}：第 {pi + 1} 頁找不到 {'、'.join(v for k, v in COLS.items() if k not in col)} 欄，採購單格式可能跟之前不一樣，請人工畫。")
        for y0, y1 in zip(below, below[1:]):
            nos = [w for w in within(*col["no"], y0, y1) if re.fullmatch(r"\d{1,4}", w[4])]
            if not nos:
                continue
            cell = lambda k: within(*col[k], y0, y1)
            box = lambda k: [min(w[0] for w in cell(k)), min(w[1] for w in cell(k)), max(w[2] for w in cell(k)), max(w[3] for w in cell(k))] if cell(k) else None
            rows.append({"no": int(nos[0][4]), "sku": "".join(w[4] for w in cell("sku")), "page": pi, "band": [y0, y1], "x0": xs[0], "x1": xs[-1],
                         "qty_unit": _num("".join(w[4] for w in cell("qu"))), "qty_pcs": _num("".join(w[4] for w in cell("qp"))),
                         "qu": {"col": list(col["qu"]), "box": box("qu")}, "qp": {"col": list(col["qp"]), "box": box("qp")}, "unit": list(col["unit"])})
    if not rows:
        raise ValueError(f"{filename}：讀不到採購單的品項表格。")
    return {"po_id": po, "rows": rows}


def po_marks(d, rows):
    """訂單視窗那份資料（side_data）＋ PDF 的列 → 要畫的 {列號: {kind, actual, qty, sku}}、提醒。
    出貨量跟訂單視窗一樣：竹運出貨看出貨計畫、其他看缺貨紀錄；兩邊都沒有就照訂單出、不畫。"""
    g = d["order"]
    by = lambda grp: {x["order_id"]: x for x in (grp or {}).get("items", [])}
    ship, short = by(d["ship"]), by(d["short"])
    items = {}
    for i in g["items"]:
        if not i["removed"]:
            items.setdefault(i["supplier_sku_id"], i)
    marks, warns = {}, []
    for r in rows:
        i = items.get(r["sku"])
        if i is None:
            warns.append(f"第 {r['no']} 列 {r['sku']}：系統裡這張單沒有這個品項（可能蝦皮改過單），不畫。")
            continue
        z, k = ship.get(i["id"]), short.get(i["id"])
        final = z["final"] if z else k["final"] if k else None
        qty = r["qty_pcs"] if r["qty_pcs"] is not None else i["qty"]
        if r["qty_pcs"] is not None and r["qty_pcs"] != i["qty"]:
            warns.append(f"第 {r['no']} 列 {r['sku']}：PDF 數量 {r['qty_pcs']:g}，系統是 {i['qty']:,}，請確認是最新的採購單。")
        if final is None or final >= qty:
            continue
        marks[r["no"]] = {"kind": "noship" if final <= 0 else "short", "actual": final, "qty": qty, "sku": r["sku"], "name": i["sku_name"]}
    return marks, warns


def draw_po(data, rows, marks, stamp_png=None):
    """原 PDF ＋ 要畫的 → 畫好的 PDF；回 (bytes, 畫了幾列)。"""
    doc = pymupdf.open(stream=data, filetype="pdf")
    size = None
    if stamp_png:
        pm = pymupdf.Pixmap(stamp_png); size = (pm.width, pm.height)

    def stamp(page, cx, cy, w, h):
        if stamp_png:
            page.insert_image(_stamp_rect(cx, cy, w, h, size), stream=stamp_png, keep_proportion=True, overlay=True)
        else:                                            # 還沒上傳印章：紅框寫字，先能用
            r = pymupdf.Rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
            page.draw_rect(r, color=RED, width=1)
            page.insert_textbox(r, "業務專用章", fontsize=6, fontname="china-t", color=RED, align=1)
    drawn = 0
    for r in rows:
        mk = marks.get(r["no"])
        if not mk:
            continue
        page = doc[r["page"]]
        y0, y1 = r["band"]; ym = (y0 + y1) / 2
        if mk["kind"] == "noship":
            page.draw_line((r["x0"] + 2, ym), (r["x1"] - 2, ym), color=RED, width=1.8)
        else:
            ratio = (r["qty_unit"] / r["qty_pcs"]) if r["qty_unit"] and r["qty_pcs"] else 1
            for key, val in (("qu", mk["actual"] * ratio), ("qp", mk["actual"])):
                c = r[key]
                if c["box"]:
                    bx0, by0, bx1, by1 = c["box"]
                    page.draw_line((bx0 - 2, by1 + 1), (bx1 + 2, by0 - 1), color=RED, width=1.4)
                    page.insert_text((bx0 - 1, by0 - 2.5), f"{val:g}", fontsize=10, fontname="hebo", color=(0, 0, 0))
                else:
                    page.insert_text((c["col"][0] + 3, ym), f"{val:g}", fontsize=10, fontname="hebo", color=(0, 0, 0))
            ux0, ux1 = r["unit"]
            stamp(page, (ux0 + ux1) / 2, ym, min(ux1 - ux0 + 6, 46), min(y1 - y0 - 2, 40))
        drawn += 1
    return doc.tobytes(), drawn


def desc_marks(marks):
    parts = []
    for v in marks.values():
        what = "不出貨" if v["kind"] == "noship" else f"{v['qty']:g}→{v['actual']:g}"
        parts.append(f"{v['sku']} {what}")
    return "、".join(parts)


@shopee_bp.route("/api/shopee/po_pdf", methods=["POST"])
def api_po_pdf_upload():
    """上傳區丟進來的 PDF 先到這裡：蝦皮採購單收下；不是採購單的（勇信配送明細表）回在 others，畫面再交給勇信核對。"""
    from .side import side_data
    files = [f for f in request.files.getlist("file") if f and f.filename]
    operator, stamp = _operator(), now()
    out, others, errs = [], [], []
    conn = get_conn()
    try:
        for f in files:
            data = f.read()
            if not is_po_pdf(data):
                if data[:5] == b"%PDF-" and not has_text(data):
                    errs.append(f"{f.filename}：這份 PDF 讀不到字（自己印成 PDF 的，字會變成圖形）。蝦皮採購單請上傳後台下載的原檔。")
                else:
                    others.append(f.filename)
                continue
            try:
                p = parse_po_pdf(data, f.filename)
            except ValueError as exc:
                errs.append(str(exc)); continue
            g = next((x for x in overview(conn, {"q": p["po_id"], "deleted": "1"}) if x["po_id"] == p["po_id"]), None)
            if g is None:
                errs.append(f"{f.filename}：系統裡沒有採購單 {p['po_id']}，請先上傳這張單的蝦皮採購單 Excel。"); continue
            no = g["order_no"]
            conn.execute("DELETE FROM shp_po_pdf WHERE order_no = ?", (no,))
            conn.execute("INSERT INTO shp_po_pdf (order_no, filename, pdf_b64, rows_json, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?)",
                         (no, f.filename, base64.b64encode(data).decode("ascii"), json.dumps(p["rows"], ensure_ascii=False), operator, stamp))
            _log(conn, g["line"], no, "", "", "shopee_po_pdf", "蝦皮採購單 PDF", "", f"{len(p['rows'])} 列", operator, "import", f.filename)
            marks, warns = po_marks(side_data(conn, no, {}), p["rows"])
            out.append({"order_no": no, "filename": f.filename, "rows": len(p["rows"]),
                        "noship": sum(1 for v in marks.values() if v["kind"] == "noship"), "short": sum(1 for v in marks.values() if v["kind"] == "short"),
                        "marks": desc_marks(marks), "warnings": warns})
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "orders": out, "others": others, "errors": errs})


@shopee_bp.route("/api/shopee/po_pdf/<no>/file")
def api_po_pdf_file(no):
    from .side import side_data
    conn = get_conn()
    try:
        pp = _row(conn.execute("SELECT * FROM shp_po_pdf WHERE order_no = ?", (no,)))
        if pp is None:
            return jsonify({"error": f"{no} 的蝦皮採購單 PDF 還沒上傳，請把蝦皮後台下載的 PDF 丟進上傳區。"}), 404
        d = side_data(conn, no, request.args)
        if d is None:
            return jsonify({"error": f"找不到訂單 {no}，請重新整理頁面。"}), 404
        rows = json.loads(pp["rows_json"] or "[]")
        marks, _w = po_marks(d, rows)
        st = load_stamp(conn)
        pdf, drawn = draw_po(base64.b64decode(pp["pdf_b64"]), rows, marks, st["png"] if st else None)
        operator, stamp = _operator(), now()
        conn.execute("UPDATE shp_po_pdf SET drawn_by = ?, drawn_at = ? WHERE order_no = ?", (operator, stamp, no))
        _log(conn, d["order"]["line"], no, "", "", "shopee_po_draw", "畫蝦皮採購單", "", desc_marks(marks) or "沒有缺貨，照原樣", operator, "manual", pp["filename"])
        conn.commit()
    finally:
        conn.close()
    base = pp["filename"].rsplit(".", 1)[0] if "." in pp["filename"] else pp["filename"]
    resp = send_file(io.BytesIO(pdf), as_attachment=True, download_name=f"{base}_已畫.pdf", mimetype="application/pdf")
    resp.headers["X-Shopee-Drawn"] = str(drawn)
    resp.headers["X-Shopee-Stamp"] = "1" if st else "0"
    return resp


__all__ = ["is_po_pdf", "has_text", "parse_po_pdf", "po_marks", "draw_po", "desc_marks", "api_po_pdf_upload", "api_po_pdf_file"]
