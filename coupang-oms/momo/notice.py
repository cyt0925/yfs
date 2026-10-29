"""④ 待畫入庫通知單（Jerry 2026-10-29 看過用 Becky 真檔畫的示意圖「OK 照這樣做 ④ 開工」）。只有瑪氏（③ 核過的單）。

- MOMO 後台下載的「入庫通知單」PDF 丟上傳區：通知單號＝入庫單號；表格是一格一格的矩形，照格子找列（一列一品）、找「數量」和「商品有效日期」兩欄。
  文字抓得到座標，所以不用猜位置；表格畫法變了就會對不到，會列出來。
- Becky 的畫法：不出貨（勇信沒出、或改成 0 不下採）→ 整列畫一條粗線、商品有效日期那格蓋業務專用章；
  實際出貨數跟 MOMO 訂的不一樣（勇信少出，或除不盡捨去之後本來就少）→ 數量畫掉、上面寫實際出貨數、蓋章。實際出貨數＝勇信箱數 × 箱入數（MOMO 的單位）。
- 業務專用章在「資料與設定」上傳（圖片，白底自動去掉，存 mst_meta momo_stamp）；沒上傳就畫一個紅框寫「業務專用章」，先能用。
- 原 PDF 存在 momo_notice（一張入庫單一份，再傳就換掉），下載時才畫，所以 ③ 的判斷後來改了，再下載一次就是新的。
"""
import base64
import datetime as _dt  # noqa: F401
import re

import pymupdf

from .check import check_orders
from .common import *  # noqa: F401,F403

NO_RE = re.compile(r"通知單號[:：]\s*(\d{8}-\d{4})")
STAMP_KEY = "momo_stamp"


def is_notice(data):
    try:
        doc = pymupdf.open(stream=data, filetype="pdf")
        text = doc[0].get_text() if len(doc) else ""
    except Exception:  # noqa: BLE001
        return False
    return "入庫通知單" in text and bool(NO_RE.search(text))


def _cells(page):
    """表格的格子：畫出來的矩形（寬高都超過 8pt 的）。"""
    out = set()
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "re":
                r = it[1]
                if r.width > 8 and r.height > 8:
                    out.add((round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1)))
    return sorted(out, key=lambda r: (r[1], r[0]))


def parse_notice(data, filename=""):
    """→ {inbound_no, rows:[{page, product_no, item_no, qty, qty_box, band, x0, x1, qty_col, exp_col}], warnings}"""
    try:
        doc = pymupdf.open(stream=data, filetype="pdf")
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"{filename}：無法開啟此 PDF 檔。") from exc
    out = {"inbound_no": "", "rows": [], "warnings": [], "pages": len(doc)}
    for pno, page in enumerate(doc):
        text = page.get_text()
        m = NO_RE.search(text)
        if m and not out["inbound_no"]:
            out["inbound_no"] = m.group(1)
        words = page.get_text("words")
        head = next((w for w in words if w[4].startswith("商品編碼")), None)
        if head is None:
            continue
        cells = _cells(page)
        hx, hy = (head[0] + head[2]) / 2, (head[1] + head[3]) / 2
        hcell = next((c for c in cells if c[0] <= hx <= c[2] and c[1] <= hy <= c[3]), None)
        if hcell is None:
            out["warnings"].append(f"第 {pno + 1} 頁找不到表頭的格子"); continue
        hband = [c for c in cells if abs(c[1] - hcell[1]) < 1 and abs(c[3] - hcell[3]) < 1]
        tx0, tx1 = min(c[0] for c in hband), max(c[2] for c in hband)

        def col_of(prefix):
            hw = [w for w in words if hcell[1] <= w[1] <= hcell[3]]
            w = next((w for w in hw if w[4].startswith(prefix)), None) or next((w for w in hw if prefix in w[4]), None)
            if w is None:
                return None
            cx = (w[0] + w[2]) / 2
            c = next((c for c in hband if c[0] <= cx <= c[2]), None)
            return (c[0], c[2]) if c else None
        qcol, ecol = col_of("數量"), col_of("商品有效")
        if not qcol or not ecol:
            out["warnings"].append(f"第 {pno + 1} 頁表頭找不到「數量」或「商品有效日期」欄"); continue
        bands = sorted({(c[1], c[3]) for c in cells if abs(c[0] - tx0) < 1 and c[1] >= hcell[3] - 0.5 and c[3] - c[1] > 20})
        for y0, y1 in bands:
            inside = [w for w in words if y0 <= (w[1] + w[3]) / 2 <= y1]
            if any(w[4].startswith("總合") for w in inside):
                continue
            first = [w for w in inside if tx0 <= w[0] <= hband[0][2] and re.fullmatch(r"\d{8}", w[4])]
            second = [w for w in inside if hband[0][2] <= w[0] <= hband[1][2] and re.fullmatch(r"\d{3}", w[4])] if len(hband) > 1 else []
            qw = [w for w in inside if qcol[0] <= (w[0] + w[2]) / 2 <= qcol[1] and re.fullmatch(r"\d[\d,]*", w[4])]
            if not first:
                continue
            out["rows"].append({"page": pno, "product_no": first[0][4], "item_no": second[0][4] if second else "",
                                "qty": int(qw[0][4].replace(",", "")) if qw else None, "qty_box": [round(v, 1) for v in qw[0][:4]] if qw else None,
                                "band": [y0, y1], "x0": tx0, "x1": tx1, "qty_col": list(qcol), "exp_col": list(ecol)})
    if not out["inbound_no"]:
        raise ValueError(f"{filename}：這份 PDF 不是 MOMO 入庫通知單（找不到通知單號）。")
    if not out["rows"]:
        raise ValueError(f"{filename}：入庫通知單讀不到品項的格子，表格畫法可能跟之前不一樣，請人工畫。")
    return out


# ── 印章 ────────────────────────────────────────────────────────────────────
def load_stamp(conn):
    row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (STAMP_KEY,)).fetchone()
    if not row or not row["value"]:
        return None
    try:
        d = json.loads(row["value"])
        return {"png": base64.b64decode(d["png"]), "updated_by": d.get("updated_by", ""), "updated_at": d.get("updated_at", "")}
    except (ValueError, KeyError):
        return None


def _clean_stamp(data):
    """白底去掉（印章掃描圖都是白紙紅章），存成 PNG。"""
    from PIL import Image
    im = Image.open(io.BytesIO(data)).convert("RGBA")
    px = im.load()
    w, h = im.size
    for y in range(h):
        for x in range(w):
            r, g, b, _a = px[x, y]
            if r > 200 and g > 200 and b > 200:
                px[x, y] = (255, 255, 255, 0)
    buf = io.BytesIO(); im.save(buf, "PNG")
    return buf.getvalue()


@momo_bp.route("/api/momo/stamp", methods=["POST"])
def api_stamp_upload():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify({"error": "請選一張印章的圖片（jpg 或 png）。"}), 400
    try:
        png = _clean_stamp(f.read())
    except Exception:  # noqa: BLE001
        return jsonify({"error": "讀不了這張圖，請用 jpg 或 png。"}), 400
    operator, stamp = _operator(), now()
    value = json.dumps({"png": base64.b64encode(png).decode("ascii"), "updated_by": operator, "updated_at": stamp})
    conn = get_conn()
    try:
        if conn.execute("SELECT 1 FROM mst_meta WHERE key = ?", (STAMP_KEY,)).fetchone():
            conn.execute("UPDATE mst_meta SET value = ? WHERE key = ?", (value, STAMP_KEY))
        else:
            conn.execute("INSERT INTO mst_meta (key, value) VALUES (?, ?)", (STAMP_KEY, value))
        _log(conn, MARS, "", "", "", "momo_stamp", "業務專用章", "", "更新圖片", operator, "manual", f.filename)
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "updated_at": stamp})


@momo_bp.route("/api/momo/stamp")
def api_stamp_info():
    conn = get_conn()
    try:
        st = load_stamp(conn)
    finally:
        conn.close()
    return jsonify({"has": bool(st), "updated_by": st["updated_by"] if st else "", "updated_at": st["updated_at"] if st else ""})


@momo_bp.route("/api/momo/stamp.png")
def api_stamp_png():
    conn = get_conn()
    try:
        st = load_stamp(conn)
    finally:
        conn.close()
    if not st:
        return jsonify({"error": "還沒上傳印章。"}), 404
    return send_file(io.BytesIO(st["png"]), mimetype="image/png")


# ── 畫 ──────────────────────────────────────────────────────────────────────
def marks_for(ck):
    """③ 的結果 → 每品要畫什麼：{(商品編號, 單品編號): {"kind": noship|short, "actual": n, "qty": n, "name"}}；還不能畫的回在 problems。"""
    marks, problems = {}, []
    for it in ck["items"]:
        key = (it["product_no"], it["item_no"])
        if it["decision"] in ("wait", "pending"):
            problems.append(f"{it['factory_code']}：{it['why']}")
        elif it["decision"] == "noship":
            marks[key] = {"kind": "noship", "actual": 0, "qty": it["qty"], "name": it["name"], "factory_code": it["factory_code"]}
        elif it["actual"] is not None and it["actual"] != (it["qty"] or 0):
            marks[key] = {"kind": "short", "actual": it["actual"], "qty": it["qty"], "name": it["name"], "factory_code": it["factory_code"]}
    return marks, problems


def _stamp_rect(cx, cy, w, h, png_size):
    """印章放進 w×h 的框，置中、不變形。"""
    pw, ph = png_size
    k = min(w / pw, h / ph)
    sw, sh = pw * k, ph * k
    return pymupdf.Rect(cx - sw / 2, cy - sh / 2, cx + sw / 2, cy + sh / 2)


def draw_notice(data, rows, marks, stamp_png=None):
    """原 PDF ＋ 要畫的 → 畫好的 PDF；回 (bytes, 畫了幾列, 對不到的品項)。"""
    doc = pymupdf.open(stream=data, filetype="pdf")
    size = None
    if stamp_png:
        pm = pymupdf.Pixmap(stamp_png); size = (pm.width, pm.height)
    drawn, seen = 0, set()

    def stamp(page, cx, cy, w, h):
        if stamp_png:
            page.insert_image(_stamp_rect(cx, cy, w, h, size), stream=stamp_png, keep_proportion=True, overlay=True)
        else:                                            # 還沒上傳印章：紅框寫字，先能用
            r = pymupdf.Rect(cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
            page.draw_rect(r, color=(0.85, 0, 0), width=1)
            page.insert_textbox(r, "業務專用章", fontsize=6, fontname="china-t", color=(0.85, 0, 0), align=1)
    for row in rows:
        mk = marks.get((row["product_no"], row["item_no"]))
        if not mk:
            continue
        seen.add((row["product_no"], row["item_no"]))
        page = doc[row["page"]]
        y0, y1 = row["band"]; ym = (y0 + y1) / 2
        if mk["kind"] == "noship":
            page.draw_line((row["x0"] + 3, ym), (row["x1"] - 3, ym), color=(0, 0, 0), width=2.2)
            ex0, ex1 = row["exp_col"]
            stamp(page, (ex0 + ex1) / 2, ym, min(ex1 - ex0 + 10, 48), 30)
        else:
            qx0, qx1 = row["qty_col"]
            if row["qty_box"]:
                bx0, by0, bx1, by1 = row["qty_box"]
                page.draw_line((bx0 - 1, (by0 + by1) / 2), (bx1 + 1, (by0 + by1) / 2), color=(0, 0, 0), width=1.3)
                page.insert_text((bx0, by0 - 2.5), f"{mk['actual']:g}", fontsize=max(7, min(10, by1 - by0)), fontname="helv", color=(0, 0, 0))
                stamp(page, (qx0 + qx1) / 2, min(by1 + 12, y1 - 10), 36, 20)
            else:
                page.insert_text((qx0 + 2, ym), f"{mk['actual']:g}", fontsize=9, fontname="helv", color=(0, 0, 0))
                stamp(page, (qx0 + qx1) / 2, ym + 12, 36, 20)
        drawn += 1
    missing = [v["factory_code"] for k, v in marks.items() if k not in seen]
    return doc.tobytes(), drawn, missing


def _desc_marks(marks):
    parts = []
    for v in marks.values():
        what = "不出貨" if v["kind"] == "noship" else f"{v['qty']}→{v['actual']:g}"
        parts.append(f"{v['factory_code']} {what}")
    return "、".join(parts) or "沒有要畫的"


@momo_bp.route("/api/momo/notice", methods=["POST"])
def api_notice_upload():
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "請上傳 MOMO 入庫通知單（PDF）。"}), 400
    operator, stamp = _operator(), now()
    out, errs = [], []
    conn = get_conn()
    try:
        for f in files:
            data = f.read()
            try:
                p = parse_notice(data, f.filename)
            except ValueError as exc:
                errs.append(str(exc)); continue
            no = p["inbound_no"]
            ck = check_orders(conn, [no]).get(no)
            if ck is None:
                errs.append(f"{f.filename}：{no} 不在「待核缺貨與效期」（不是瑪氏、還沒回填 EIP 單號，或系統裡沒有這張單），先不收。"); continue
            conn.execute("DELETE FROM momo_notice WHERE inbound_no = ?", (no,))
            conn.execute("INSERT INTO momo_notice (inbound_no, filename, pdf_b64, rows_json, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?)",
                         (no, f.filename, base64.b64encode(data).decode("ascii"), json.dumps(p["rows"], ensure_ascii=False), operator, stamp))
            _log(conn, MARS, no, "", "", "momo_notice", "入庫通知單", "", f"{len(p['rows'])} 列", operator, "import", f.filename)
            marks, problems = marks_for(ck)
            keys = {(r["product_no"], r["item_no"]) for r in p["rows"]}
            unmatched = [v["factory_code"] for k, v in marks.items() if k not in keys]
            out.append({"inbound_no": no, "warehouse": ck["warehouse"], "filename": f.filename, "rows": len(p["rows"]),
                        "noship": sum(1 for v in marks.values() if v["kind"] == "noship"), "short": sum(1 for v in marks.values() if v["kind"] == "short"),
                        "marks": _desc_marks(marks), "problems": problems, "unmatched": unmatched, "ready": not problems and not unmatched,
                        "need": bool(marks), "warnings": p["warnings"]})
        conn.commit()
    finally:
        conn.close()
    if not out and errs:
        return jsonify({"error": "入庫通知單沒有收進來。", "details": errs}), 400
    return jsonify({"ok": True, "orders": out, "errors": errs})


@momo_bp.route("/api/momo/notice/<no>/file")
def api_notice_file(no):
    conn = get_conn()
    try:
        nt = _row(conn.execute("SELECT * FROM momo_notice WHERE inbound_no = ?", (no,)))
        if nt is None:
            return jsonify({"error": f"{no} 的入庫通知單還沒上傳，請先把 MOMO 下載的 PDF 丟進上傳區。"}), 404
        ck = check_orders(conn, [no]).get(no)
        if ck is None:
            return jsonify({"error": f"{no} 不在「待核缺貨與效期」，沒有核對結果可以畫。"}), 400
        marks, problems = marks_for(ck)
        if problems:
            return jsonify({"error": f"還有 {len(problems)} 個品項要先決定，還不能畫。", "details": problems}), 400
        st = load_stamp(conn)
        pdf, drawn, missing = draw_notice(base64.b64decode(nt["pdf_b64"]), json.loads(nt["rows_json"] or "[]"), marks, st["png"] if st else None)
        if missing:
            return jsonify({"error": f"入庫通知單上找不到這幾個品項，請確認是同一張單：{'、'.join(missing)}"}), 400
        operator, stamp = _operator(), now()
        conn.execute("UPDATE momo_notice SET drawn_by = ?, drawn_at = ? WHERE inbound_no = ?", (operator, stamp, no))
        _log(conn, MARS, no, "", "", "momo_notice_draw", "畫入庫通知單", "", _desc_marks(marks), operator, "manual", nt["filename"])
        conn.commit()
    finally:
        conn.close()
    base = nt["filename"].rsplit(".", 1)[0] if "." in nt["filename"] else nt["filename"]
    resp = send_file(io.BytesIO(pdf), as_attachment=True, download_name=f"{base}_已畫.pdf", mimetype="application/pdf")
    resp.headers["X-Momo-Drawn"] = str(drawn)
    resp.headers["X-Momo-Stamp"] = "1" if st else "0"
    return resp


__all__ = ["is_notice", "parse_notice", "draw_notice", "marks_for", "load_stamp", "api_notice_upload", "api_notice_file",
           "api_stamp_upload", "api_stamp_info", "api_stamp_png"]
