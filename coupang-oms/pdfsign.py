"""驗收單 PDF 批次簽名。

取代原本要開 Google Colab 貼程式碼跑的作業方式（SOP-CP-CPG-001）。
做的事情跟那支 Colab 程式一樣：逐頁搜尋「出貨確認（廠商簽名）」這行字，
找到就在它正下方蓋上簽名圖。差別在於：

  - 簽名圖存在系統裡（每人一張），不用每次上傳，也就沒有「不小心傳了
    兩張簽名檔」這種要人自己記得避開的地雷。
  - 尺寸、位移、關鍵字都從設定讀，酷澎哪天改了版面或字，不用改程式。
  - 每份檔案各自回報結果，搜不到欄位的會明確說出來——Colab 版只印一行
    「共簽 N 個欄位」，N 是 0 的時候很容易被忽略，人就以為簽好了。
"""

import io
import re

import pymupdf

# 酷澎 PO 單號：13 開頭的 14 位數字。驗收單 PDF 內文會出現，用來把
# 簽好的檔案掛回對應的那張單，事後查得到「這張單是誰簽的」。
PO_PATTERN = re.compile(r"\b1\d{13}\b")

DEFAULT_GEOMETRY = {
    "keyword": "出貨確認（廠商簽名）",
    "width": 65,
    "height": 22,
    "offset_x": 0,
    "offset_y": 2,
}


class SignError(Exception):
    """這份檔案沒辦法處理（壞檔、加密、不是 PDF…）。"""


def extract_po_number(text, filename=""):
    """先從 PDF 內文找 PO 單號，找不到再退回檔名找。

    兩邊都找不到就回空字串——沒有單號還是要能簽、能下載，只是歸檔時
    對不回哪張單而已，不該因此整份檔案失敗。
    """
    match = PO_PATTERN.search(text or "")
    if match:
        return match.group(0)
    match = PO_PATTERN.search(filename or "")
    return match.group(0) if match else ""


def sign_pdf(pdf_bytes, signature_bytes, geometry=None):
    """在一份 PDF 上蓋章，回傳 (簽好的 bytes, 蓋了幾處, PO 單號)。

    蓋章位置沿用原本 Colab 版的算法：以關鍵字方框的左邊界為 x 起點、
    下緣往下 offset_y 為 y 起點，往右下畫出 width×height 的框。這是
    同事已經在實際驗收單上調好的數字，不要自作聰明改掉。
    """
    geo = {**DEFAULT_GEOMETRY, **(geometry or {})}
    keyword = geo["keyword"]

    try:
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:
        raise SignError(f"打不開這個 PDF（{exc}）") from exc

    try:
        if doc.needs_pass:
            raise SignError("這份 PDF 有密碼保護，請先解除密碼再上傳。")

        total = 0
        text_parts = []
        for page in doc:
            text_parts.append(page.get_text())
            for area in page.search_for(keyword):
                x0, _y0, _x1, y1 = area
                rect = pymupdf.Rect(
                    x0 + geo["offset_x"],
                    y1 + geo["offset_y"],
                    x0 + geo["offset_x"] + geo["width"],
                    y1 + geo["offset_y"] + geo["height"],
                )
                page.insert_image(rect, stream=signature_bytes)
                total += 1

        full_text = "\n".join(text_parts)
        po_number = extract_po_number(full_text)

        if total == 0:
            # 分辨「這份根本沒有文字層」跟「有文字但沒有這個關鍵字」——
            # 前者幾乎都是掃描檔，後者通常是關鍵字設定跟實際版面對不上。
            # 兩種的處理方式完全不同，訊息要講清楚使用者才知道怎麼辦。
            if not full_text.strip():
                raise SignError(
                    "這份 PDF 抓不到任何文字，應該是掃描檔（圖片），"
                    "沒辦法自動找簽名欄位。")
            raise SignError(f"這份 PDF 裡找不到「{keyword}」這個欄位。")

        out = io.BytesIO()
        doc.save(out)
        return out.getvalue(), total, po_number
    finally:
        doc.close()



# ── 驗收單缺貨劃線（Jerry 2026-10-08，先只做瑪氏）──────────────────────
# 送勇信的驗收單，缺貨的品項原本是人手在 PDF 上改：整筆不出就整列畫一條粗線，部分下修就把原本的出貨數量塗黑、
# 旁邊寫新的數量（兩聯都改，合計不改）。系統裡的出貨數量在「瑪氏出貨 ⑤ 勇信缺貨」按確認後就是對的，簽名時照它畫。

SKU_HEADER = "SKU No."
QTY_HEADER = "出貨數量"
NAME_HEADER = "品名"


def _digits(text):
    return re.sub(r"\D", "", text or "")


def find_item_rows(doc):
    """找出驗收單「3.商品資訊」每一列：[{page, no, sku, name, qty, row_rect, qty_rect, table_rect}]。
    兩聯都會各列一次（同一個 no 出現兩次）。表格跨頁時，續頁沒有表頭，用上一頁表格的欄位位置認。
    品名太長時一列會被表格線切成兩段（第二段 No 是空的），接回上一列。"""
    out, cols = [], None
    for pno, page in enumerate(doc):
        try:
            tables = page.find_tables().tables
        except Exception:  # noqa: BLE001 — 這頁認不出表格就跳過，不影響簽名
            continue
        for t in tables:
            data = t.extract()
            if not data:
                continue
            head = [(c or "").replace("\n", "") for c in data[0]]
            xs = [round(c[0]) for c in t.rows[0].cells if c]
            if SKU_HEADER in head and QTY_HEADER in head:
                cols = {"sku": head.index(SKU_HEADER), "qty": head.index(QTY_HEADER),
                        "name": head.index(NAME_HEADER) if NAME_HEADER in head else None, "n": len(head), "xs": xs}
                start = 1
            elif cols and len(head) == cols["n"] and len(xs) == len(cols["xs"]) and all(abs(a - b) <= 3 for a, b in zip(xs, cols["xs"])):
                start = 0                                            # 續頁：欄位位置跟上一頁一樣
            else:
                continue
            cur = None
            for ri in range(start, len(data)):
                cells, rects = data[ri], t.rows[ri].cells
                first = (cells[0] or "").strip()
                if any((c_ or "").strip().startswith("合計") for c_ in cells):   # 合計那列（不一定在第一格）
                    break
                rect = pymupdf.Rect(t.bbox[0], min(c[1] for c in rects if c), t.bbox[2], max(c[3] for c in rects if c))
                if first.isdigit():
                    cur = {"page": pno, "no": int(first), "sku": "", "name": "", "qty_text": "", "row_rect": rect,
                           "qty_cell": None, "table_rect": pymupdf.Rect(t.bbox)}
                    out.append(cur)
                elif cur is None:
                    continue
                else:
                    cur["row_rect"] = cur["row_rect"] | rect           # 被切成兩段的同一列
                cur["sku"] += _digits(cells[cols["sku"]])
                if cols["name"] is not None:
                    cur["name"] += (cells[cols["name"]] or "").replace("\n", "")
                q = (cells[cols["qty"]] or "").strip()
                if q:
                    cur["qty_text"] += q
                    qc = pymupdf.Rect(rects[cols["qty"]])
                    cur["qty_cell"] = qc if cur["qty_cell"] is None else cur["qty_cell"] | qc
    for r in out:
        r["qty"] = int(_digits(r["qty_text"])) if _digits(r["qty_text"]) else None
        r["qty_rect"] = None
        if r["qty_cell"] is not None:                                   # 數字本身的範圍（塗黑用）
            words = [w for w in doc[r["page"]].get_text("words", clip=r["qty_cell"]) if _digits(w[4])]
            if words:
                r["qty_rect"] = pymupdf.Rect(min(w[0] for w in words), min(w[1] for w in words), max(w[2] for w in words), max(w[3] for w in words))
    return out


def shortage_plan(rows, new_qty):
    """哪幾列要改：new_qty＝{SKU: 系統裡現在的出貨數量}。系統比驗收單少才改；0＝整列劃掉，其他＝塗掉原數量寫新的。
    回傳 (要畫的列, 一品一筆的清單, 驗收單上有、系統裡找不到的 SKU)。"""
    marks, items, missing, seen = [], [], [], set()
    for r in rows:
        new = new_qty.get(r["sku"])
        if new is None:
            if r["sku"] and r["sku"] not in seen:
                missing.append(r["sku"])
        elif r["qty"] is not None and new < r["qty"]:
            kind = "strike" if new == 0 else "reduce"
            marks.append(dict(r, new_qty=new, kind=kind))
            if (r["no"], r["sku"]) not in seen:
                items.append({"no": r["no"], "sku": r["sku"], "name": r["name"], "pdf_qty": r["qty"], "new_qty": new, "kind": kind})
        seen.add((r["no"], r["sku"])); seen.add(r["sku"])
    return marks, items, missing


def draw_marks(doc, marks):
    """照 Jerry 給的手改範例畫：整筆不出＝整列一條 3pt 黑線；部分下修＝原數量塗黑、右邊寫新數量。"""
    for m in marks:
        page = doc[m["page"]]
        if m["kind"] == "strike":
            y = (m["row_rect"].y0 + m["row_rect"].y1) / 2
            page.draw_line((m["table_rect"].x0 + 6, y), (m["table_rect"].x1 - 30, y), color=(0, 0, 0), width=3)
        else:
            box = m["qty_rect"] or m["qty_cell"]
            if box is None:
                continue
            box = pymupdf.Rect(box.x0 - 1.5, box.y0 - 1, box.x1 + 1.5, box.y1 + 1)
            page.draw_rect(box, color=(0, 0, 0), fill=(0, 0, 0), width=0)
            page.insert_text((box.x1 + 2, box.y1 - 1.5), str(m["new_qty"]), fontsize=11, fontname="helv", color=(0, 0, 0))


def mark_shortage(pdf_bytes, lookup):
    """在驗收單上劃缺貨。lookup(PO 單號) → {SKU: 系統出貨數量}，不是要劃的單（例如不是瑪氏）回 None。
    回傳 (新的 bytes, 結果)。結果：po_number、applies（這張有沒有要比對）、items、missing。"""
    try:
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:  # noqa: BLE001
        raise SignError(f"打不開這個 PDF（{exc}）") from exc
    try:
        po = extract_po_number("\n".join(p.get_text() for p in doc))
        new_qty = lookup(po) if po else None
        info = {"po_number": po, "applies": new_qty is not None, "items": [], "missing": [], "rows": 0}
        if new_qty is None:
            return pdf_bytes, info
        rows = find_item_rows(doc)
        marks, info["items"], info["missing"] = shortage_plan(rows, new_qty)
        info["rows"] = len({(r["no"], r["sku"]) for r in rows})
        if not marks:
            return pdf_bytes, info
        draw_marks(doc, marks)
        out = io.BytesIO()
        doc.save(out)
        return out.getvalue(), info
    finally:
        doc.close()

# DPI 用來把 PDF 座標（點，pt）跟畫面上顯示的圖片像素互換。校正時
# 前端會秀出這張渲染圖，管理員拖曳簽名框、存檔前再依這個 DPI 換算回
# pt——渲染跟簽名蓋章用的是同一套點/像素換算比例，拖出來的位置才會
# 跟實際簽名時完全對得上。
CALIBRATE_DPI = 150


def render_for_calibration(pdf_bytes, keyword):
    """找出關鍵字所在的那一頁，渲染成圖片，回傳給前端讓人用拖曳的
    方式校正簽名要蓋在哪裡——比要求非技術同事去猜四個數字直觀得多。
    """
    try:
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
    except Exception as exc:
        raise SignError(f"打不開這個 PDF（{exc}）") from exc

    try:
        if doc.needs_pass:
            raise SignError("這份 PDF 有密碼保護，請先解除密碼再試一次。")

        page_index, rect = None, None
        for i, page in enumerate(doc):
            areas = page.search_for(keyword)
            if areas:
                page_index, rect = i, areas[0]
                break

        if page_index is None:
            raise SignError(
                f"這份 PDF 裡找不到「{keyword}」這個欄位，換一份範例，"
                "或先把上面的關鍵字改對再試一次。")

        page = doc[page_index]
        pix = page.get_pixmap(dpi=CALIBRATE_DPI)
        scale = CALIBRATE_DPI / 72  # PDF 座標單位是 72 dpi 的「點」
        return {
            "image_bytes": pix.tobytes("png"),
            "dpi": CALIBRATE_DPI,
            "image_width": pix.width,
            "image_height": pix.height,
            "keyword_rect_px": [rect.x0 * scale, rect.y0 * scale,
                                rect.x1 * scale, rect.y1 * scale],
        }
    finally:
        doc.close()


def validate_signature(image_bytes):
    """確認上傳的簽名圖真的是張圖，順便回報尺寸給畫面顯示。

    擋在存進資料庫之前——不然壞檔會等到真的要簽名時才爆，那時候已經
    有人排隊等著用了。
    """
    try:
        pix = pymupdf.Pixmap(image_bytes)
        return {"width": pix.width, "height": pix.height}
    except Exception as exc:
        raise SignError(f"這不是能讀取的圖片檔（{exc}）") from exc
