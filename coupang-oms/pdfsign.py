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
