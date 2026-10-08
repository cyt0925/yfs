"""瑪氏出貨（mars/）的端到端驗證：假的瑪氏商品總表＋假訂單彙總表裡線別是瑪氏的單，
走「商品總表上傳 → ② 匯訂單 → 拆單 → 產出 zip（EMMA 匯入檔＋EIP 採購表）→ 回填 EIP 採購單號／約倉時間 → 訂單改了之後
→ 採購單設定（倉庫資料、固定文字、假日）→ 產出瑪氏採購單（V2 範本）」。

執行：python test_mars.py（設了 DATABASE_URL 就跑 PostgreSQL）
"""
import datetime
import io
import json
import os
import sys
import tempfile
import zipfile
from urllib.parse import unquote

import openpyxl
import xlrd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FAKE_M = os.path.join(BASE_DIR, "samples", "master", "fake")
FAKE = os.path.join(BASE_DIR, "samples", "mars", "fake")
MASTER = os.path.join(FAKE, "假_瑪氏商品總表.xlsx")

_tmp = tempfile.mkdtemp(prefix="oms_mars_test_")
sys.path.insert(0, BASE_DIR)
import db  # noqa: E402
db.DATA_DIR = os.path.join(_tmp, "資料與設定")
db.DB_PATH = os.path.join(db.DATA_DIR, "test.db")
db.BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")
import app as app_module  # noqa: E402
import mars.common as mc  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(f"  {'✓' if cond else '✗'} {name}" + (f"  — {detail}" if detail else ""))


def up(client, url, name_or_bytes, filename=None, **form):
    data = open(name_or_bytes, "rb").read() if isinstance(name_or_bytes, str) else name_or_bytes
    d = dict(form); d["file"] = (io.BytesIO(data), filename or os.path.basename(name_or_bytes))
    return client.post(url, data=d, content_type="multipart/form-data")


def import_orders(client, data, filename):
    pv = up(client, "/api/master/import/preview", data, filename).get_json()
    return client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"], "add_missing_products": True})


def splits(client, d1, d2=None):
    return client.get(f"/api/mars/splits?from={d1}&to={d2 or d1}").get_json()


def special_sheet(rows):
    """照整合表的表頭做一份只有瑪氏幾列的訂單彙總表。"""
    wb = openpyxl.load_workbook(os.path.join(FAKE_M, "假_訂單彙總表_9月_第一次.xlsx")); ws = wb.worksheets[0]
    hdr = [c.value for c in ws[1]]
    ws.delete_rows(2, ws.max_row)
    for r in rows:
        ws.append([r.get(h) for h in hdr])
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def yx_pdf(pages):
    """假的勇信配送明細表：pages = [(EIP單號, 倉, 指送日期 yyyy/mm/dd, [(瑪氏貨號, 箱數, 效期), ...])]，文字排法照真檔。"""
    import pymupdf
    doc = pymupdf.open()
    for eip, wh, date, items in pages:
        page = doc.new_page()
        lines = ["桃園市大園區高鐵北路三段100號", "客戶單號：7816797526", f"BMS100222-永豐商店酷澎-{wh}", "出貨廠商名稱：A5888-台灣瑪氏股份有限公司",
                 f"指送日期：{date}", "配 送 明 細 表(預排鮮度)", "產品編號", "產品名稱", "入數", "配送數量", f"收貨單號：{eip}-酷澎-T"]
        for n, (code, qty, exp) in enumerate(items, start=1):
            lines += [f"58880{code}", str(n), f"假品名 {code} 1:6:10", f"{qty}C  ", "625D1TAP01", exp]
        lines += ["備註:箱麥/ 對點/酷澎嘜頭+驗收單", "單據金額:", f"合計箱數:{sum(q for _, q, _ in items)}", "客戶簽章:"]
        y = 40
        for ln in lines:
            page.insert_text((40, y), ln, fontname="china-t", fontsize=9); y += 14
    return doc.tobytes()


def base_row(**kw):
    r = {"訂單類型": "一般", "線別": "瑪氏", "PO單號": "13000000699901", "到貨倉別": "TAO3", "地址": "桃園市大園區中山南路472號",
         "交付日期": datetime.datetime(2026, 9, 30), "NO": 1, "品牌": "測試", "酷澎下單單價(含稅)": 100}
    r.update(kw)
    return r


def main():
    db.init_db(); app_module.app.config["TESTING"] = True
    c = app_module.app.test_client()

    print("\n【0】登入、頁面、選單")
    check("沒登入打 API → 401", c.get("/api/mars/status").status_code == 401)
    c.post("/login", data={"username": "小真", "password": "changeme123"})
    # CI 把幾套測試跑在同一個 PostgreSQL 裡，前一套會留下訂單、匯入批次；這裡先清空，不依賴資料庫是空的
    app_module._write_users(app_module.get_users(), {"Jerry", "小真"})
    r = c.post("/api/master/reset", json={"confirm": "清空資料", "orders": True, "products": True})
    check("開始前先清空訂單與主檔（含瑪氏的表）", r.status_code == 200, r.get_data(as_text=True)[:120])
    html = c.get("/mars").get_data(as_text=True)
    check("/mars 頁面打得開、有拆單與 JS 版本號", "瑪氏出貨" in html and "btn-gen" in html and "mars.js?v=" in html)
    check("線別工具的瑪氏欄有「瑪氏出貨」，不再是準備中", "/mars" in c.get("/master").get_data(as_text=True))

    print("\n【1】瑪氏商品總表")
    check("還沒上傳時狀態是空的", c.get("/api/mars/status").get_json()["last_upload"] is None)
    r = up(c, "/api/mars/products/import", MASTER); d = r.get_json()
    check("上傳：7 個料號、21 列（箱盒包各一）、品類分得出來", r.status_code == 200 and d["codes"] == 7 and d["rows"] == 21
          and d["by_category"] == {"CHO": 2, "GUM": 2, "PET": 3} and not d["warnings"], str(d)[:200])
    st = c.get("/api/mars/status").get_json()
    check("狀態：上次上傳、料號數、② 裡有瑪氏訂單的日子", st["last_upload"]["filename"] == "假_瑪氏商品總表.xlsx" and st["codes"] == 7 and st["last_order_import"] is None)
    check("不是商品總表的檔 → 400 講清楚要哪幾欄", up(c, "/api/mars/products/import", os.path.join(FAKE_M, "假_酷澎主檔_瑪氏.xlsx")).status_code == 400)
    check("壞掉的檔 → 400", up(c, "/api/mars/products/import", b"not excel", "x.xlsx").status_code == 400)
    wb = openpyxl.load_workbook(MASTER); ws = wb["商品表"]
    ws.append([1, "Candy", "x", "M00000001", "怪品類", 1, 1, 1, 1, "", "", "箱", "R", 1, "", 1, "", ""])
    ws.append([2, "Gum", "x", "M00000002", "怪單位", 1, 1, 1, 1, "", "", "打", "", 1, "", 1, "", ""])
    buf = io.BytesIO(); wb.save(buf)
    d = up(c, "/api/mars/products/import", buf.getvalue(), "改過的.xlsx").get_json()
    check("單位不是箱盒包的列跳過、Category 認不得的列標出來", d["codes"] == 8 and any("打" in w for w in d["warnings"]) and any("Candy" in w for w in d["warnings"]), str(d["warnings"]))
    d = up(c, "/api/mars/products/import", MASTER).get_json()
    check("再傳原本那份＝整份覆蓋：跟上一版比拿掉 1 個", d["codes"] == 7 and d["removed"] == 1 and d["removed_examples"] == ["M00000001"], str(d)[:160])

    print("\n【2】② 匯訂單：地址、報價備註也存下來")
    for n in ("假_訂單彙總表_9月_第一次.xlsx", "假_訂單彙總表_9月_第二次.xlsx"):
        check(f"匯 {n}", import_orders(c, os.path.join(FAKE_M, n), None).status_code == 200)
    conn = db.get_conn()
    o = dict(conn.execute("SELECT * FROM mst_orders WHERE po_number = '13000000600028' AND yf_sku = 'M36752197'").fetchone()); conn.close()
    check("mst_orders 有存地址", o["address"] != "", o["address"])
    st = c.get("/api/mars/status").get_json()
    check("狀態：最近一次有瑪氏單的匯入是第二次那份", st["last_order_import"]["filename"] == "假_訂單彙總表_9月_第二次.xlsx" and "2026-09-18" in st["dates"], str(st["last_order_import"]))

    print("\n【3】拆單：9/18 那張 PO 拆成 5 份")
    v = splits(c, "2026-09-18")
    names = [s["filename"] for s in v["splits"]]
    check("一張 PO、5 份（品類 × 單位 × 中標各不同）", v["summary"]["pos"] == 1 and v["summary"]["files"] == 5, str(names))
    check("檔名照 Alice 的規則：訂單系統拆單表_PO_到貨日_倉_品類_單位_中標_序號", "訂單系統拆單表_13000000600028_20260918_TAO1_CHO_盒_需貼中標_01.xlsx" in names
          and "訂單系統拆單表_13000000600028_20260918_TAO1_PET_箱_不貼中標_01.xlsx" in names, str(names))
    s_cho = next(s for s in v["splits"] if s["category"] == "CHO" and s["label"] == "")
    it = s_cho["items"][0]
    check("箱數 = 出貨數量 ÷ 整合表箱入數（36 ÷ 12 = 3）", it["cases"] == 3 and s_cho["cases_total"] == 3, str(it["cases"]))
    check("整合表箱入數跟商品總表不一樣 → 提醒、箱數照整合表", any("訂單彙總表 12、商品總表 6" in x for x in it["issues"]), str(it["issues"]))
    check("帶進商品總表的欄（瑪氏貨號、品名、價格、每箱中盒數…）", it["mars_code"] == "36752197" and it["mars_name"].startswith("【巧克棒】") and it["price"] == 2410.0 and it["inner_per_case"] == 12)
    check("全部都還沒產出", all(s["status"] == "new" and s["id"] is None for s in v["splits"]))
    total_file = sum(r["cases_total"] for r in v["splits"])
    check("拆完的箱數加總 = 這張 PO 的箱數加總（3+1+7+1+3）", total_file == 15, str(total_file))
    v2 = splits(c, "2026-09-11")
    pv = [s for s in v2["splits"] if s["po_number"] == "13000000600020" and s["category"] == "PET"]
    check("9/11 TAO5：同品類、同單位、同中標的兩個品項放同一份", len(pv) == 1 and pv[0]["item_count"] == 2, str([(s["filename"], s["item_count"]) for s in pv]))
    mc.SPLIT_BY_UNIT = False
    try:
        import mars.split as ms
        ms.SPLIT_BY_UNIT = False
        vn = splits(c, "2026-09-05")
    finally:
        mc.SPLIT_BY_UNIT = True; ms.SPLIT_BY_UNIT = True
    vy = splits(c, "2026-09-05")
    check("關掉「依單位拆」：檔名單位那段固定寫箱", all("_箱_" in s["filename"] for s in vn["splits"]) and vn["split_by_unit"] is False)

    cal = c.get("/api/mars/calendar?month=2026-09").get_json()
    d18 = cal["days"].get("2026-09-18")
    check("月曆：9 月每天有幾張 PO、幾箱、幾份；9/18 是 1 張 PO、15 箱、5 份、0 份填了單號", d18 and d18["pos"] == 1 and d18["cases"] == 15 and d18["files"] == 5 and d18["filled"] == 0 and "2026-09-11" in cal["days"], str(d18))
    check("月曆：月份亂填 → 400", c.get("/api/mars/calendar?month=2026-9").status_code == 400)

    print("\n【4】特殊狀況：組出商品、對不到、箱數不是整數")
    sheet = special_sheet([
        base_row(**{"SKU ID": "900000000000001", "條碼(國條)": "4700000000001", "永豐料號": "M60055599", "報價備註": "M55500001", "品名": "[組出] 口香糖王 超值包",
                    "下單數量(酷澎單位)": 60, "出貨數量": 60, "單位": "盒", "箱入數": 6}),
        base_row(**{"SKU ID": "900000000000002", "條碼(國條)": "4700000000002", "永豐料號": "M99999999", "報價備註": "M99999999", "品名": "商品總表沒有的",
                    "下單數量(酷澎單位)": 10, "出貨數量": 10, "單位": "盒", "箱入數": 10}),
        base_row(**{"SKU ID": "900000000000003", "條碼(國條)": "4712227918056", "永豐料號": "M81232885", "品名": "口香糖王 薄荷",
                    "下單數量(酷澎單位)": 25, "出貨數量": 25, "單位": "盒", "箱入數": 20}),
        base_row(**{"SKU ID": "900000000000004", "條碼(國條)": "4717986828288", "永豐料號": "M69072565", "品名": "喵愛餡 盒",
                    "下單數量(酷澎單位)": 24, "出貨數量": 24, "單位": "盒", "箱入數": 12}),
        base_row(**{"SKU ID": "900000000000005", "條碼(國條)": "4717986828289", "永豐料號": "M69072565", "品名": "喵愛餡 包",
                    "下單數量(酷澎單位)": 144, "出貨數量": 144, "單位": "包", "箱入數": 144}),
    ])
    check("匯特殊狀況那份", import_orders(c, sheet, "特殊狀況.xlsx").status_code == 200)
    conn = db.get_conn()
    q = dict(conn.execute("SELECT quote_note FROM mst_orders WHERE sku_id = '900000000000001'").fetchone()); conn.close()
    check("報價備註存進 mst_orders", q["quote_note"] == "M55500001", str(q))
    v = splits(c, "2026-09-30")
    combo = next(i for s in v["splits"] for i in s["items"] if i["sku_id"] == "900000000000001")
    check("下採料號＝報價備註：組出商品下採 M55500001（永豐料號 M60055599 只是酷澎編號）", combo["via"] == "下採料號" and combo["purchase_code"] == "M55500001" and combo["yf_sku"] == "M60055599" and combo["cases"] == 10, str(combo)[:200])
    plain = next(i for s in v["splits"] for i in s["items"] if i["sku_id"] == "900000000000004")
    check("報價備註空白 → 下採料號退回永豐料號", plain["purchase_code"] == "M69072565" and plain["via"] == "下採料號")
    check("對不到商品總表的列出來、不進任何一份", [u["yf_sku"] for u in v["unmatched"]] == ["M99999999"]
          and all(i["yf_sku"] != "M99999999" for s in v["splits"] for i in s["items"]))
    frac = next(s for s in v["splits"] if any(i["sku_id"] == "900000000000003" for i in s["items"]))
    check("箱數 25÷20 不是整數 → 那份標成擋下來", frac["blocking"] and "不是整數" in frac["blocking"][0], str(frac["blocking"]))
    pet = [s for s in v["splits"] if s["category"] == "PET"]
    check("同料號、盒跟包：依單位拆成兩份", len(pet) == 2 and {s["unit"] for s in pet} == {"盒", "包"}, str([s["filename"] for s in pet]))
    check("箱那列才有的採購單箱備註，盒的品項不會亂帶", all(not i["po_case_note"] for s in pet for i in s["items"]))

    print("\n【5】下載 EIP 採購單（之前叫產出 EMMA 與 EIP 檔、產出拆單表）")
    r = c.post("/api/mars/splits/generate", json={"from": "2026-09-30", "to": "2026-09-30"})
    check("有箱數不是整數的 → 400、點名是哪個、什麼都沒存", r.status_code == 400 and "M81232885" in json.dumps(r.get_json(), ensure_ascii=False)
          and splits(c, "2026-09-30")["splits"][0]["id"] is None, str(r.get_json())[:200])
    r = c.post("/api/mars/splits/ensure", json={"split_key": frac["split_key"]})
    check("單列自動存：箱數不是整數的那份 → 400，不存", r.status_code == 400 and "不是整數" in json.dumps(r.get_json(), ensure_ascii=False))
    orow = c.get("/api/master/orders?month=2026-09&q=900000000000003").get_json()["rows"][0]
    r = c.put(f"/api/master/orders/{orow['id']}", json={"version": orow["version"], "qty_ship": 40, "reason": "缺貨"})
    check("到 ② 把出貨數量改成 40（2 箱）", r.status_code == 200, str(r.get_json())[:120])
    r = c.post("/api/mars/splits/generate", json={"from": "2026-09-30", "to": "2026-09-30"})
    check("還有對不到的 → 409 要人確認，列出是哪個", r.status_code == 409 and r.get_json()["needs_ack"] and "M99999999" in r.get_json()["details"][0])
    # 單列自動存（Jerry 2026-10-06：還沒產出的列也要直接有按鈕，按了先存那一份）
    one = splits(c, "2026-09-30")["splits"][0]
    r = c.post("/api/mars/splits/ensure", json={"split_key": one["split_key"]})
    check("單列自動存：同一張 PO 有對不到商品總表的 → 409 先問", r.status_code == 409 and r.get_json()["needs_ack"] and "M99999999" in r.get_json()["details"][0])
    r = c.post("/api/mars/splits/ensure", json={"split_key": one["split_key"], "ack_unmatched": True}); d1 = r.get_json()
    r2 = c.post("/api/mars/splits/ensure", json={"split_key": one["split_key"]}); d2 = r2.get_json()
    v1 = {x["split_key"]: x for x in splits(c, "2026-09-30")["splits"]}
    check("確認後只存那一份（狀態變已產出、其他還是還沒產出），再按一次回同一個、不重存",
          d1["created"] and d1["id"] and d2["id"] == d1["id"] and not d2["created"] and v1[one["split_key"]]["status"] == "generated"
          and all(x["status"] == "new" for k, x in v1.items() if k != one["split_key"]), str(d1))
    check("自動存的那份可以直接下載 EIP 上傳用採購表、填 EIP 單號",
          c.get(f"/api/mars/splits/{d1['id']}/file?kind=eip").status_code == 200 and c.put(f"/api/mars/splits/{d1['id']}", json={"slot_time": "09:00~12:00"}).status_code == 200)
    check("單列自動存：拆單鍵亂填 → 400；訂單已經沒有的那份 → 404",
          c.post("/api/mars/splits/ensure", json={"split_key": "亂填"}).status_code == 400
          and c.post("/api/mars/splits/ensure", json={"split_key": "P-NONE|2026-09-30|TAO3|GUM|箱|"}).status_code == 404)
    c.put(f"/api/mars/splits/{d1['id']}", json={"slot_time": ""})
    r = c.post("/api/mars/splits/generate", json={"from": "2026-09-30", "to": "2026-09-30", "ack_unmatched": True})
    cd = unquote(r.headers.get("Content-Disposition", ""))
    check("確認後下載 zip，檔名「瑪氏EIP採購單_到貨日」", r.status_code == 200 and "瑪氏EIP採購單_20260930.zip" in cd, cd)
    z = zipfile.ZipFile(io.BytesIO(r.data)); names = z.namelist()
    v = splits(c, "2026-09-30")
    check("zip 裡只有每份的 EIP 上傳用採購表（_EIP上傳.xls）；EMMA 匯入檔不放（這時還沒有 EIP 單號，Jerry 2026-10-06）",
          len(names) == len(v["splits"]) and all(f"{s['filename'][:-5]}_EIP上傳.xls" in names for s in v["splits"])
          and all(n.endswith("_EIP上傳.xls") for n in names) and not any(n.startswith("酷澎訂單匯入") for n in names), str(names))
    check("產出後狀態：已產出、等 EIP 單號", all(s["status"] == "generated" and s["id"] for s in v["splits"]))
    gum_combo = next(s for s in v["splits"] if any(i["sku_id"] == "900000000000001" for i in s["items"]))
    wsx = openpyxl.load_workbook(io.BytesIO(c.get(f"/api/mars/splits/{gum_combo['id']}/file?kind=split").data)).active   # 拆單表還能單獨下載
    hdr = [x.value for x in wsx[1]]
    check("拆單表帶了 Alice 指定的商品總表欄（瑪氏貨號、Category、品名、價格、每箱中盒數、每箱產品數、單位類別、中盒貼標、備註、採購單箱備註）",
          all(h in hdr for h in ("瑪氏貨號", "Category", "品名(永豐建檔)", "系統價格(未稅)", "每箱中盒數", "每箱產品數(最小單位)", "單位類別", "中盒貼標", "備註", "採購單箱備註")), str(hdr))
    check("拆單表最後兩欄是 EIP 採購單號、約倉時間（還沒填 → 空）", hdr[-2:] == ["EIP採購單號", "約倉時間"] and wsx.cell(2, len(hdr) - 1).value in (None, ""))
    tot = [x.value for x in wsx[wsx.max_row]]
    check("最後一列合計：箱數加總", tot[0] == "合計" and tot[hdr.index("箱數")] == gum_combo["cases_total"], str(tot[:15]))
    sh = xlrd.open_workbook(file_contents=z.read(gum_combo["filename"][:-5] + "_EIP上傳.xls")).sheet_by_index(0)
    eip = [sh.row_values(i) for i in range(1, sh.nrows) if sh.cell_value(i, 1)]
    check("EIP 採購表：同一個下採料號加總、單位箱、數量是整數、備註照採購表轉換的瑪氏格式",
          {e[1]: e[6] for e in eip} == {"M55500001": 10.0, "M81232885": 2.0} and all(e[3] == "箱" for e in eip) and eip[0][8] == "MARS入倉9/30瑪氏送酷澎-TAO3倉", str(eip))
    check("對不到的 M99999999 不在任何一個檔", all("M99999999" not in json.dumps([x.value for row in openpyxl.load_workbook(io.BytesIO(z.read(n))).active.iter_rows() for x in row]) for n in names if n.endswith(".xlsx")))

    print("\n【6】回填 EIP 採購單號、約倉時間")
    s1, s2 = v["splits"][0], v["splits"][1]
    put = lambda sid, body: c.put(f"/api/mars/splits/{sid}", json=body)  # noqa: E731
    r = put(s1["id"], {"eip_po": "PO12345"})
    check("格式不對 → 400，講要長怎樣", r.status_code == 400 and "PO202609004" in r.get_json()["error"])
    r = put(s1["id"], {"eip_po": " po202609300 "})
    check("前後空白、小寫都接受，存成 PO202609300", r.status_code == 200 and r.get_json()["split"]["eip_po"] == "PO202609300")
    r = put(s2["id"], {"eip_po": "PO202609300"})
    check("同一張 PO 但品類不同的另一份填同一個號碼 → 可以存，但提醒依規則應分開、說用在哪一份",
          r.status_code == 200 and "應分開" in r.get_json()["warning"] and s1["filename"] in r.get_json()["warning"], str(r.get_json()))
    put(s2["id"], {"eip_po": ""})
    r = put(s1["id"], {"slot_time": "12:30~15:30（1台車）"})
    check("約倉時間存得進去", r.status_code == 200 and r.get_json()["split"]["slot_time"] == "12:30~15:30（1台車）")
    v = splits(c, "2026-09-30"); f1 = next(s for s in v["splits"] if s["id"] == s1["id"])
    check("狀態變成已回填", f1["status"] == "filled" and f1["eip_po"] == "PO202609300")
    wsx = openpyxl.load_workbook(io.BytesIO(c.get(f"/api/mars/splits/{s1['id']}/file?kind=split").data)).active
    check("單份重新下載：拆單表裡有 EIP 採購單號與約倉時間", wsx.cell(2, wsx.max_column - 1).value == "PO202609300" and wsx.cell(2, wsx.max_column).value == "12:30~15:30（1台車）")
    r = c.get(f"/api/mars/splits/{s1['id']}/file?kind=eip")
    check("單份 EIP 採購表也能重新下載", r.status_code == 200 and "_EIP上傳.xls" in unquote(r.headers.get("Content-Disposition", "")))
    logs = c.get("/api/master/logs?q=PO202609300&limit=1000").get_json()["logs"]
    check("回填有記修改歷程", any(l["field"] == "mars_eip_po" and l["new_value"] == "PO202609300" for l in logs))

    print("\n【7】回填之後訂單又改了：已送 EIP 的那份不偷偷改")
    target = next(i for i in f1["items"])
    orow = next(o for o in c.get("/api/master/orders?month=2026-09&q=" + target["sku_id"]).get_json()["rows"])
    r = c.put(f"/api/master/orders/{orow['id']}", json={"version": orow["version"], "qty_ship": orow["qty_ship"] + orow["box_size_file"], "reason": "酷澎要求"})
    check("到 ② 多加一箱", r.status_code == 200, str(r.get_json())[:100])
    v = splits(c, "2026-09-30"); f1 = next(s for s in v["splits"] if s["id"] == s1["id"])
    check("狀態：EIP 送出後訂單有變、寫出差在哪", f1["status"] == "changed_after_eip" and "出貨" in f1["diff"], f1.get("diff"))
    check("畫面上顯示的還是送出去的那份（舊數量）", next(i for i in f1["items"] if i["sku_id"] == target["sku_id"])["qty_ship"] == target["qty_ship"])
    c.post("/api/mars/splits/generate", json={"from": "2026-09-30", "to": "2026-09-30", "ack_unmatched": True})
    wsx = openpyxl.load_workbook(io.BytesIO(c.get(f"/api/mars/splits/{s1['id']}/file?kind=split").data)).active
    hdr = [x.value for x in wsx[1]]
    qcol = hdr.index("出貨數量") + 1
    scol = hdr.index("SKU ID") + 1
    got = [wsx.cell(rr, qcol).value for rr in range(2, wsx.max_row + 1) if wsx.cell(rr, scol).value == target["sku_id"]]
    check("再按產出：已回填的那份內容不動（還是舊數量）", got == [target["qty_ship"]], f"{got} vs {target['qty_ship']}")
    put(s1["id"], {"eip_po": ""})
    c.post("/api/mars/splits/generate", json={"from": "2026-09-30", "to": "2026-09-30", "ack_unmatched": True})
    v = splits(c, "2026-09-30"); f1 = next(s for s in v["splits"] if s["id"] == s1["id"])
    check("把單號清掉再產出：更新成新數量、狀態回到已產出", f1["status"] == "generated" and next(i for i in f1["items"] if i["sku_id"] == target["sku_id"])["qty_ship"] == target["qty_ship"] + orow["box_size_file"])

    print("\n【8】訂單改期：拆不出來的舊拆單")
    put(s2["id"], {"eip_po": "PO202609301"})
    moved = next(i for i in next(s for s in v["splits"] if s["id"] == s2["id"])["items"])
    r = c.put("/api/master/pos/date", json={"po_number": "13000000699901", "delivery_date": "2026-09-29", "reason": "沒車"})
    check("整張 PO 改期到 9/29", r.status_code == 200, str(r.get_json())[:120])
    v = splits(c, "2026-09-30")
    gone = [s for s in v["splits"] if s["status"] == "gone"]
    check("9/30 存過的三份都標「訂單已沒有這份」（有人可能已經拿著檔），不會拆出新的", len(gone) == 3 and s2["id"] in [s["id"] for s in gone]
          and not [s for s in v["splits"] if s["status"] == "new"], str([(s['status'], s['filename']) for s in v['splits']]))
    v29 = splits(c, "2026-09-29")
    check("9/29 拆出新的（還沒產出）", v29["splits"] and all(s["status"] == "new" for s in v29["splits"]), str([s["filename"] for s in v29["splits"]]))
    r = c.post("/api/mars/splits/generate", json={"from": "2026-09-29", "to": "2026-09-30", "ack_unmatched": True})
    v = splits(c, "2026-09-29", "2026-09-30")
    check("兩天一起產出：9/30 沒回填的舊拆單刪掉，只剩已回填的那份", r.status_code == 200 and [s["id"] for s in v["splits"] if s["delivery_date"] == "2026-09-30"] == [s2["id"]]
          and all(s["status"] == "generated" for s in v["splits"] if s["delivery_date"] == "2026-09-29"), str([(s["delivery_date"], s["status"]) for s in v["splits"]]))

    print("\n【8b】採購單設定：倉庫資料、固定文字、假日")
    ps = c.get("/api/mars/po/settings").get_json()
    codes = [w["code"] for w in ps["warehouses"]]
    check("倉庫清單從瑪氏訂單自動列出來（TAO1、TAO3、TAO5…），預設名稱是 永豐商店酷澎-倉", "TAO3" in codes and "TAO1" in codes
          and next(w for w in ps["warehouses"] if w["code"] == "TAO3")["name"] == "永豐商店酷澎-TAO3", str(codes))
    w3 = next(w for w in ps["warehouses"] if w["code"] == "TAO3")
    check("地址沒填時用訂單上的地址、標出電話與 ship-to 還缺", w3["address"] == "桃園市大園區中山南路472號" and w3["address_from_orders"] and w3["missing"] == ["電話", "ship-to"], str(w3))
    check("預設設定：提前 2 個工作天、效期依品類（巧、糖 3/5，寵 1/2）、EMMA 固定字、特殊需求有進倉時間那行", ps["settings"]["lead_days"] == 2 and ps["settings"]["shelf_req_cat"] == {"CHO": "3/5效期以上", "GUM": "3/5效期以上", "PET": "1/2效期以上"}
          and ps["settings"]["emma_customer"] == "N71178" and ps["settings"]["emma_payment"] == "貨到不付款"
          and "{約倉時間}" in ps["settings"]["special_text"] and ps["template"] == "mars_po_v2.xlsx")
    r = c.put("/api/mars/warehouses/TAO3", json={"phone": "03-1234567", "ship_to": "17600001", "contact": "王小姐"})
    w3 = next(w for w in r.get_json()["warehouses"] if w["code"] == "TAO3")
    check("填 TAO3 的電話、ship-to、聯絡人 → 存起來、不再缺", r.status_code == 200 and w3["saved"] and w3["missing"] == [] and w3["contact"] == "王小姐" and w3["address"] == "桃園市大園區中山南路472號", str(w3))
    logs = c.get("/api/master/logs?q=TAO3&limit=1000").get_json()["logs"]
    check("倉庫資料有記修改歷程", any(l["field"] == "mars_warehouse" for l in logs))
    check("提前天數亂填 → 400", c.put("/api/mars/po/settings", json={"lead_days": 20}).status_code == 400)
    r = c.put("/api/mars/po/settings", json={"holidays": "2026-09-28\n9/29x"})
    check("假日看不懂 → 400 點名是哪個", r.status_code == 400 and "9/29x" in r.get_json()["error"], str(r.get_json()))
    r = c.put("/api/mars/po/settings", json={"holidays": "2026-09-28, 2026/10/9", "lead_days": "2"})
    check("假日接受幾種寫法、存成 ISO 且排序", r.status_code == 200 and r.get_json()["settings"]["holidays"] == ["2026-09-28", "2026-10-09"], str(r.get_json()))
    check("再讀一次還在", c.get("/api/mars/po/settings").get_json()["settings"]["holidays"] == ["2026-09-28", "2026-10-09"])
    r = c.put("/api/mars/po/settings", json={"shelf_req_cat": {"GUM": "2/3效期以上", "XXX": "亂填"}, "emma_shipto": "0002"})
    check("效期依品類：只改 GUM，其他不動；EMMA 固定字可改", r.status_code == 200 and r.get_json()["settings"]["shelf_req_cat"] == {"CHO": "3/5效期以上", "GUM": "2/3效期以上", "PET": "1/2效期以上"} and r.get_json()["settings"]["emma_shipto"] == "0002", str(r.get_json()["settings"]["shelf_req_cat"]))
    c.put("/api/mars/po/settings", json={"shelf_req_cat": {"GUM": "3/5效期以上"}})
    c.put("/api/mars/po/settings", json={"emma_shipto": "0001"})
    r = c.put("/api/mars/warehouses/TAO3", json={"special_note": "司機需加入TAO3 line領取排隊號碼"})
    check("倉庫資料多一格特殊需求加註", r.status_code == 200 and next(w for w in r.get_json()["warehouses"] if w["code"] == "TAO3")["special_note"] == "司機需加入TAO3 line領取排隊號碼")
    # 上傳倉庫資料表（Jerry 給的格式：倉別／中文地址／電話），檔案有值就蓋、空白不動、電話轉成 02-5592-7598
    wbw = openpyxl.Workbook(); wsw = wbw.active
    wsw.append(["倉別", "中文地址", "電話"])
    wsw.append(["TAO1", "桃園市大園區建國路102號4樓", "+886-02-55927598\t\t"])
    wsw.append(["TAO3", "", "+886-0911556291"])
    wsw.append(["tao9", "桃園市大園區建國路102號3樓", None])
    wsw.append([None, "沒倉別的列", "123"])
    bufw = io.BytesIO(); wbw.save(bufw)
    r = up(c, "/api/mars/warehouses/import", bufw.getvalue(), "倉庫.xlsx"); d = r.get_json()
    check("上傳倉庫資料表：3 個倉，TAO3 更新、TAO1／TAO9 新增，沒倉別的列跳過", r.status_code == 200 and d["rows"] == 3 and d["added"] == 2 and d["updated"] == 1, str(d)[:200])
    whs = {w["code"]: w for w in d["warehouses"]}
    check("電話轉成台灣寫法：+886-02-55927598 → 02-5592-7598、+886-0911556291 → 0911-556-291", whs["TAO1"]["phone"] == "02-5592-7598" and whs["TAO3"]["phone"] == "0911-556-291", str([whs["TAO1"]["phone"], whs["TAO3"]["phone"]]))
    check("檔案空白的不動：TAO3 地址還是訂單的、ship-to 還在；倉別小寫轉大寫", whs["TAO3"]["address"] == "桃園市大園區中山南路472號" and whs["TAO3"]["ship_to"] == "17600001" and "TAO9" in whs and whs["TAO9"]["missing"] == ["電話", "ship-to"], str(whs.get("TAO9")))
    check("再傳一次同樣的檔 → 全部沒變", up(c, "/api/mars/warehouses/import", bufw.getvalue(), "倉庫.xlsx").get_json()["same"] == 3)
    check("不是倉庫資料表 → 400 講清楚", up(c, "/api/mars/warehouses/import", MASTER).status_code == 400)
    # 網頁上直接加一個倉、刪一個倉
    r = c.post("/api/mars/warehouses", json={"code": " tao8 ", "address": "桃園市某處", "phone": "+886-03-1234567", "ship_to": "17600008"})
    whs = {w["code"]: w for w in r.get_json()["warehouses"]}
    check("直接加一個倉：倉別轉大寫、電話 +886 也轉、列在清單裡", r.status_code == 200 and r.get_json()["code"] == "TAO8" and whs["TAO8"]["phone"] == "03-123-4567" and whs["TAO8"]["missing"] == [] and not whs["TAO8"]["in_orders"], str(whs.get("TAO8")))
    check("倉別已經有 → 400", c.post("/api/mars/warehouses", json={"code": "TAO3"}).status_code == 400)
    check("倉別亂填 → 400", c.post("/api/mars/warehouses", json={"code": "倉庫 8"}).status_code == 400)
    check("訂單裡有的倉不能刪", c.delete("/api/mars/warehouses/TAO3").status_code == 400 and whs["TAO3"]["in_orders"])
    r = c.delete("/api/mars/warehouses/TAO8")
    check("手動加的倉可以刪、清單裡不見", r.status_code == 200 and all(w["code"] != "TAO8" for w in r.get_json()["warehouses"]))
    check("刪沒有的倉 → 404", c.delete("/api/mars/warehouses/TAO8").status_code == 404)

    print("\n【8c】產出瑪氏採購單（V2 範本）")
    v = splits(c, "2026-09-29", "2026-09-30")
    f2 = next(x for x in v["splits"] if x["id"] == s2["id"])          # 9/30、已填 PO202609301、還沒填約倉時間
    check("清單帶每份差什麼才能產採購單：這份填了單號、約倉時間不強制 → 可以產", f2["po_missing"] == [], str(f2["po_missing"]))
    check("9/29 那幾份：沒填 EIP 單號", all("尚未填 EIP 採購單號" in x["po_missing"] for x in v["splits"] if x["delivery_date"] == "2026-09-29"))
    check("提示（不擋）：約倉時間沒填", any("約倉時間" in h for h in f2["po_hints"]), str(f2["po_hints"]))
    import mars.po as mp
    st0 = c.get("/api/mars/po/settings").get_json()["settings"]
    fake_wh = mp._empty_wh("TAO9", st0)
    check("倉庫資料沒填：不擋，只提示會留空", mp.po_missing(f2, f2["items"], fake_wh) == [] and any("TAO9" in h and "留空" in h for h in mp.po_hints(dict(f2, warehouse="TAO9"), f2["items"], fake_wh)), str(mp.po_hints(dict(f2, warehouse="TAO9"), f2["items"], fake_wh)))
    ws9 = openpyxl.load_workbook(io.BytesIO(mp.po_file(dict(f2, warehouse="TAO9"), f2["items"], fake_wh, st0))).active
    check("倉庫資料沒填時產出的採購單：倉別名還是帶、地址／電話／ship-to 留空", ws9["C6"].value == "永豐商店酷澎-TAO9" and ws9["C7"].value is None and ws9["C8"].value is None and ws9["F7"].value is None, str([ws9[x].value for x in ("C6", "C7", "C8", "F7")]))
    tail = f"({f2['unit']})" if f2["unit"] != "箱" else ""
    check("檔名照實際在用的格式：EIP單號永豐Mars採購單(箱單位)-品類貼中標_倉_酷澎PO；盒、包那份最後加(原單位)", f2["po_filename"] == f"PO202609301永豐Mars採購單(箱單位)-{f2['category']}{'貼中標' if f2['label'] == 'V' else '不貼中標'}_TAO3_13000000699901{tail}.xlsx", f2["po_filename"])
    check("本來就是箱的那份不加括弧；沒填 EIP 單號的檔名從「永豐Mars採購單」開始", all(x["po_filename"].endswith(f"({x['unit']}).xlsx") == (x["unit"] != "箱") for x in v["splits"])
          and all(x["po_filename"].startswith("永豐Mars採購單(箱單位)-") for x in v["splits"] if not x["eip_po"]), str([x["po_filename"] for x in v["splits"]]))
    r = c.get(f"/api/mars/splits/{s2['id']}/po")
    f5_noslot = openpyxl.load_workbook(io.BytesIO(r.data)).active["F5"].value.split("\n")
    check("約倉時間沒填也能下載：特殊需求裡進倉時間、幾點前抵達那兩行整行省略", r.status_code == 200 and not any("進倉時間" in x or "抵達" in x for x in f5_noslot)
          and "箱嘜，需當面對點數量" in f5_noslot, str(f5_noslot))
    r9 = next(x for x in v["splits"] if x["delivery_date"] == "2026-09-29")
    check("沒填 EIP 單號還是擋：400 講清楚", c.get(f"/api/mars/splits/{r9['id']}/po").status_code == 400)
    put(s2["id"], {"slot_time": "12:30~15:30（1台車）"})
    r = c.get(f"/api/mars/splits/{s2['id']}/po")
    check("填齊了 → 下載 .xlsx、檔名對", r.status_code == 200 and f2["po_filename"] in unquote(r.headers.get("Content-Disposition", "")), unquote(r.headers.get("Content-Disposition", "")))
    ws = openpyxl.load_workbook(io.BytesIO(r.data)).active
    check("C3 下單日＝到貨日 9/30 往前 2 個工作天、跳過 9/28 教師節、9/26～27 週末、9/25 中秋 → 9/24；F3 配送日＝9/30",
          ws["C3"].value == datetime.datetime(2026, 9, 24) and ws["F3"].value == datetime.datetime(2026, 9, 30), f"{ws['C3'].value} / {ws['F3'].value}")
    check("C5 永豐PO單號＝EIP 採購單號", ws["C5"].value == "PO202609301")
    check("C6～F8 倉庫資料：入倉倉別、地址（訂單的）、電話、ship-to（數字）、聯絡人", ws["C6"].value == "永豐商店酷澎-TAO3" and ws["C7"].value == "桃園市大園區中山南路472號"
          and ws["C8"].value == "0911-556-291" and ws["F7"].value == 17600001 and ws["F8"].value == "王小姐", str([ws[x].value for x in ("C6", "C7", "C8", "F7", "F8")]))
    check("F6 效期要求照這份的品類（巧、糖 3/5，寵 1/2）、B9 品類全名", ws["F6"].value == ("1/2效期以上" if f2["category"] == "PET" else "3/5效期以上") and ws["B9"].value == mc.CAT_FULL[f2["category"]], f"{ws['F6'].value} / {ws['B9'].value}")
    f5 = ws["F5"].value.split("\n")
    exp_first = ["需貼中盒標"] if f2["label"] == "V" else []
    check("F5 特殊需求：（要貼中標才有）需貼中盒標 → 固定文字第一行 → 這個倉的加註 → 固定文字 → 進倉時間＝約倉時間 → 請在 12:30 前抵達",
          f5[:len(exp_first)] == exp_first and f5[len(exp_first):len(exp_first) + 3] == ["箱嘜，需當面對點數量", "司機需加入TAO3 line領取排隊號碼", "酷澎嘜頭+驗收單"]
          and f5[len(exp_first) + 3] == "進倉時間12:30~15:30（1台車）" and f5[len(exp_first) + 4] == "*請在12:30前抵達，以免被算遲到，謝謝", str(f5))
    rows = [[ws.cell(rr, cc).value for cc in range(1, 14)] for rr in range(12, 12 + len(f2["items"]))]
    codes_in = {r_[0] for r_ in rows}
    check("第 12 列起每列一個下採料號（Jerry 2026-10-07 改）：A 料號、B 瑪氏貨號、D 單位需求＝出貨數量、E 箱數、F 價格、J 每箱產品數、K 每箱中盒數",
          codes_in == {i["purchase_code"] for i in f2["items"]} and all(r_[3] == sum(i["qty_ship"] for i in f2["items"] if i["purchase_code"] == r_[0])
          and r_[4] == sum(i["cases"] for i in f2["items"] if i["purchase_code"] == r_[0]) and r_[1] and r_[5] and r_[9] and r_[10] for r_ in rows), str(rows))
    from mars.po import po_rows
    combo_rows = po_rows([i for i in gum_combo["items"] if i["sku_id"] == "900000000000001"])
    check("組出商品：瑪氏採購單 A 欄放下採料號 M55500001，不是永豐料號 M60055599", [r_["code"] for r_ in combo_rows] == ["M55500001"], str(combo_rows))
    check("H 中盒需貼標照這份的中標；I 指定效期留空", all((r_[7] == "V") == (f2["label"] == "V") and r_[8] is None for r_ in rows))
    last = 12 + len(f2["items"]) - 1
    check("E10／F10 加總公式蓋到最後一列（範本原本只到 28）", ws["E10"].value == f"=SUM(E12:E{last})" and ws["F10"].value == f"=SUM(G12:G{last})", f"{ws['E10'].value} {ws['F10'].value}")
    check("G 欄小計公式還在", ws["G12"].value == "=IFERROR(E12*F12,0)")
    r = c.post("/api/mars/po/zip", json={"from": "2026-09-29", "to": "2026-09-30"})
    z = zipfile.ZipFile(io.BytesIO(r.data)); names = z.namelist()
    check("整段期間打包：只有填齊的那 1 份進 zip，其他列在說明檔，標頭寫幾份產了幾份沒", r.status_code == 200 and f2["po_filename"] in names and "未產出清單.txt" in names
          and r.headers.get("X-Mars-Po-Count") == "1" and int(r.headers.get("X-Mars-Po-Skipped")) == len(v["splits"]) - 1
          and "尚未填 EIP 採購單號" in z.read("未產出清單.txt").decode("utf-8"), str(names))
    check("zip 檔名帶到貨日區間", "瑪氏採購單_20260929-20260930.zip" in unquote(r.headers.get("Content-Disposition", "")))
    check("沒有一份能產 → 400", c.post("/api/mars/po/zip", json={"from": "2026-09-01", "to": "2026-09-02"}).status_code == 400)
    # 特殊需求：有採購單箱備註的品項 → 最後多一行；組出商品備註寫訂單料號；同料號合併
    fake_s = {"delivery_date": "2026-09-30", "eip_po": "PO202609999", "slot_time": "下午2點", "label": "", "category": "PET", "unit": "箱", "warehouse": "TAO3", "po_number": "1"}
    fake_items = [{"purchase_code": "M1", "yf_sku": "M9", "unit": "箱", "qty_ship": 10, "cases": 1, "mars_code": "1", "mars_name": "a", "price": 1, "pcs_per_case": 1, "inner_per_case": 1, "note": "", "po_case_note": "小白標"},
                  {"purchase_code": "M1", "yf_sku": "M9", "unit": "箱", "qty_ship": 20, "cases": 2, "mars_code": "1", "mars_name": "a", "price": 1, "pcs_per_case": 1, "inner_per_case": 1, "note": "舊備註", "po_case_note": ""},
                  {"purchase_code": "M1", "yf_sku": "M8", "unit": "箱", "qty_ship": 5, "cases": 1, "mars_code": "1", "mars_name": "a", "price": 1, "pcs_per_case": 1, "inner_per_case": 1, "note": "", "po_case_note": ""}]
    st = c.get("/api/mars/po/settings").get_json()["settings"]
    txt = mp.special_text(st, fake_s, fake_items)
    check("有採購單箱備註 → 特殊需求最後一行「指定品需加工貼小白標」；約倉時間裡沒有 HH:MM 就整段代入；倉庫沒加註那行不放", txt.endswith("指定品需加工貼小白標") and "進倉時間下午2點" in txt and "*請在下午2點前抵達" in txt and "{倉庫加註}" not in txt and "\n\n" not in txt, txt)
    txt2 = mp.special_text(dict(st, special_text="A\nB"), fake_s, fake_items, {"special_note": "加註"})
    check("自己改過固定文字沒放記號：加註插在第一行後面", txt2.split("\n")[:3] == ["A", "加註", "B"], txt2)

    print("\n【8d】EMMA 匯入檔（酷澎訂單匯入）")
    import mars.emma as me
    check("電話轉 +886：02-5592-7598 → +886-02-55927598、0911-556-291 → +886-0911556291", me.intl_phone("02-5592-7598") == "+886-02-55927598" and me.intl_phone("0911-556-291") == "+886-0911556291" and me.intl_phone("") == "")
    r = c.get(f"/api/mars/splits/{f2['id']}/emma")
    check("單份下載，檔名 酷澎訂單匯入_0930交貨-TAO3_PO_品類_中標(單位).xlsx", r.status_code == 200 and unquote(r.headers.get("Content-Disposition", "")).split("''")[-1].startswith(f"酷澎訂單匯入_0930交貨-TAO3_13000000699901_{f2['category']}_"), unquote(r.headers.get("Content-Disposition", "")))
    we = openpyxl.load_workbook(io.BytesIO(r.data)).active
    hdr_e = [x.value for x in we[1]]
    check("表頭照 Kate 系統的 Coupang_PO_Order", hdr_e == ["出貨備註", "訂單編號", "收件人", "收件人手機", "收件地址", "客編", "BillTo", "ShipTo", "倉庫別", "付款方式", "料號", "單位", "數量", "單價", "金額小計"], str(hdr_e))
    row2 = [x.value for x in we[2]]
    exp_qty = sum(i["cases"] for i in f2["items"] if i["purchase_code"] == row2[10])
    exp_price = 100 * next(i["box_file"] for i in f2["items"] if i["purchase_code"] == row2[10])
    check("A 出貨備註＝EIP 單號、B 酷澎 PO、C 收件人、D 手機 +886（TAO3 的 0911）、E 訂單地址、F～J 固定字、K 下採料號、L 箱、M 箱數、N 單價＝酷澎單價×箱入數、O＝M×N",
          row2[0] == "PO202609301" and row2[1] == "13000000699901" and row2[2] == "酷澎股份有限公司" and row2[3] == "+886-0911556291" and row2[4] == "桃園市大園區中山南路472號"
          and row2[5:10] == ["N71178", "B001", "0001", None, "貨到不付款"] and row2[10] in {i["purchase_code"] for i in f2["items"]} and row2[11] == "箱"
          and row2[12] == exp_qty and row2[13] == exp_price and row2[14] == exp_qty * exp_price, str(row2))
    check("最後一列 M 欄 SUBTOTAL", we.cell(we.max_row, 13).value == f"=SUBTOTAL(9,M2:M{we.max_row - 1})", str(we.cell(we.max_row, 13).value))
    others = [x["id"] for x in v["splits"] if x["id"] != f2["id"]]
    r = c.post("/api/mars/emma", json={"ids": [f2["id"]] + others[:1]})
    check("合併：有一份沒填 EIP 單號 → 409 先問、列出是哪份", r.status_code == 409 and r.get_json()["needs_ack"] and len(r.get_json()["details"]) == 1, str(r.get_json()))
    r = c.post("/api/mars/emma", json={"ids": [f2["id"]] + others[:1], "ack_missing_eip": True})
    cd = unquote(r.headers.get("Content-Disposition", "")).split("''")[-1]
    we2 = openpyxl.load_workbook(io.BytesIO(r.data)).active
    check("確認後合併：檔名 酷澎訂單匯入_0929、0930交貨-TAO3.xlsx（兩天、同倉），列數＝兩份的料號數，沒單號那份 A 欄空、標頭寫有幾個留空",
          r.status_code == 200 and cd == "酷澎訂單匯入_0929、0930交貨-TAO3.xlsx" and we2.max_row - 2 == len({(i["po_number"], i["purchase_code"]) for x in v["splits"] if x["id"] in ([f2["id"]] + others[:1]) for i in x["items"]})
          and any(we2.cell(rr, 1).value is None for rr in range(2, we2.max_row)) and r.headers.get("X-Mars-Emma-Warnings") == "1", f"{cd} rows={we2.max_row}")
    r = c.post("/api/mars/emma", json={"from": "2026-09-29", "to": "2026-09-30", "ack_missing_eip": True})
    check("整段期間全部合併：列數＝所有拆單表的料號數", r.status_code == 200 and r.headers.get("X-Mars-Emma-Rows") == str(len(v["splits"])), str(r.headers.get("X-Mars-Emma-Rows")))
    check("沒有拆單表的期間 → 400", c.post("/api/mars/emma", json={"from": "2026-09-01", "to": "2026-09-02"}).status_code == 400)
    pr = mp.po_rows(fake_items)
    check("採購單一列一個下採料號（Jerry 2026-10-07：永豐料號 M9、M8 都下採 M1 → 寫 M1 一列）：同下採料號合成一列、數量箱數加總、備註照商品總表",
          [r_["code"] for r_ in pr] == ["M1"] and pr[0]["qty"] == 35 and pr[0]["cases"] == 4 and pr[0]["note"] == "舊備註", str(pr))
    check("下單日：9/21（一）往前 2 個工作天 → 9/17（四）", mp.order_date(datetime.date(2026, 9, 21), 2, set()) == datetime.date(2026, 9, 17))
    check("內建國定假日自動跳過：9/29（二）到貨，9/28 教師節、9/25 中秋、週末都跳 → 下單日 9/23（三）", mp.order_date(datetime.date(2026, 9, 29), 2) == datetime.date(2026, 9, 23))
    check("春節：2026-02-23（一）到貨，2/16～2/20 全放、2/14～15 週末 → 下單日 2/12（四）", mp.order_date(datetime.date(2026, 2, 23), 2) == datetime.date(2026, 2, 12))
    check("2027 也有：2027-02-11（四）到貨，2/4～2/10 春節 → 下單日 2/2（二）", mp.order_date(datetime.date(2027, 2, 11), 2) == datetime.date(2027, 2, 2))
    check("額外假日還是會加上去", mp.order_date(datetime.date(2026, 9, 21), 2, {"2026-09-18"}) == datetime.date(2026, 9, 16))
    ty = datetime.date.today().year
    check(f"內建假日涵蓋今年＋明年（{ty}、{ty + 1}）；不夠了就去人事行政總處抄新的一年進 TW_HOLIDAYS", ty in mp.TW_HOLIDAYS and (ty + 1) in mp.TW_HOLIDAYS, str(sorted(mp.TW_HOLIDAYS)))
    check("假日名字查得到", mp.holiday_name("2026-10-09") == "國慶日補假" and mp.holiday_name("2026-10-08") == "")
    ps2 = c.get("/api/mars/po/settings").get_json()
    check("設定 API 回內建假日的年份與清單", ps2["builtin_holiday_years"] == sorted(mp.TW_HOLIDAYS) and "2026-09-28" in ps2["builtin_holidays"])

    print("\n【8e】勇信缺貨：PDF 比對、下修檔、改系統數量")
    import mars.shortage as msx
    v = splits(c, "2026-09-29", "2026-09-30")
    tgt = max((x for x in v["splits"] if x["delivery_date"] == "2026-09-29"), key=lambda x: max(i["cases"] for i in x["items"]))
    put(tgt["id"], {"eip_po": "PO202609305"})
    v = splits(c, "2026-09-29", "2026-09-30"); tgt = next(x for x in v["splits"] if x["id"] == tgt["id"])
    it0 = max(tgt["items"], key=lambda i: i["cases"])
    others = [i for i in tgt["items"] if i is not it0]
    pdf = yx_pdf([("PO202609305", "TAO3", "2026/09/29", [(it0["mars_code"], int(it0["cases"]) - 1, "2027/06/18")] + [(i["mars_code"], int(i["cases"]), "2027/08/11") for i in others]),
                  ("PO202609999", "TAO3", "2026/09/29", [("99999999", 3, "2027/01/01")])])
    pages = msx.parse_yx_pdf(pdf, "假.pdf")
    check("假 PDF 讀得出：2 頁、收貨單號、倉、日期、箱數、效期", len(pages) == 2 and pages[0]["eip_po"] == "PO202609305" and pages[0]["warehouse"] == "TAO3"
          and pages[0]["ship_date"] == "2026-09-29" and pages[0]["items"][0]["qty"] == int(it0["cases"]) - 1 and pages[0]["items"][0]["expiry"] == "2027-06-18" and not pages[0]["warnings"], str(pages[0])[:300])
    check("產品編號 58880＋貨號、配送數量 4C 的讀法", msx.QTY_RE.match("4C  ") and msx._mars_code("060019810") == "60019810" and msx._mars_code("10266398") == "10266398")
    r = up(c, "/api/mars/shortage/compare", pdf, "假_勇信.pdf"); res = r.get_json()
    check("比對：對到 1 張 PO，一個品項少 1 箱＝部分缺，其他沒缺；PO202609999 系統沒有 → 列未知", r.status_code == 200 and res["summary"]["pos"] == 1 and res["summary"]["partial"] == 1
          and res["summary"]["ok"] == len(others) and [u["eip_po"] for u in res["unknown"]] == ["PO202609999"], str(res["summary"]) + str(res["unknown"]))
    row0 = next(it for it in res["pos"][0]["splits"][0]["items"] if it["sku_id"] == it0["sku_id"])
    check("部分缺的那筆：缺 1 箱、改成的出貨數量＝勇信箱數×箱入數", row0["status"] == "partial" and row0["short"] == 1 and row0["new_qty"] == (int(it0["cases"]) - 1) * it0["box_file"] and not res["pos"][0]["full"], str(row0))
    check("這張 PO 還有別份沒在表裡 → 列出來、不算整張不出", len(res["pos"][0]["missing_splits"]) == len(v["splits"]) - 1 and res["downgrade_rows"] == 1, str(res["pos"][0]["missing_splits"]))
    r = c.post("/api/mars/shortage/apply", json={"result": res, "only_file": True})
    wd = openpyxl.load_workbook(io.BytesIO(r.data))["issueReport"]
    check("只下載下修檔：套酷澎範本，第 1 列代碼不動、第 2 列表頭、第 3 列起一列一個缺的品項（SKU、酷澎 PO、變更數量、原因 4）",
          r.status_code == 200 and str(wd["A1"].value).startswith("oNJf0ob5l") and [x.value for x in wd[2]] == ["SKU ID", "採購訂單ID", "變更數量", "請求原因", "附件", "評論"]
          and wd.max_row == 3 and wd["A3"].value == it0["sku_id"] and wd["B3"].value == "13000000699901" and wd["C3"].value == str(row0["new_qty"]) and wd["D3"].value == "4. 供應商庫存不足" and wd["F3"].value is None, str([x.value for x in wd[3]]))
    check("下修檔檔名帶指送日期與倉", "酷澎下修_0929交貨-TAO3.xlsx" in unquote(r.headers.get("Content-Disposition", "")), unquote(r.headers.get("Content-Disposition", "")))
    before = next(o for o in c.get("/api/master/orders?month=2026-09&q=13000000699901").get_json()["rows"] if o["sku_id"] == it0["sku_id"])["qty_ship"]
    r = c.post("/api/mars/shortage/apply", json={"result": res})
    check("確認：改了 1 筆、下修檔一起回來", r.status_code == 200 and r.headers.get("X-Mars-Changed") == "1" and r.headers.get("X-Mars-Rows") == "1", str(r.headers.get("X-Mars-Changed")))
    after = next(o for o in c.get("/api/master/orders?month=2026-09&q=13000000699901").get_json()["rows"] if o["sku_id"] == it0["sku_id"])
    check("系統裡那筆訂單的出貨數量改成勇信實際出的、標人改過", after["qty_ship"] == row0["new_qty"] and after["qty_ship"] != before and after["qty_ship_overridden"], f"{before} → {after['qty_ship']}")
    res_again = up(c, "/api/mars/shortage/compare", pdf, "假_勇信.pdf").get_json()
    row_again = next(it for it in res_again["pos"][0]["splits"][0]["items"] if it["sku_id"] == it0["sku_id"])
    check("確認過再跑一次：那筆照原本訂的箱數比、還是部分缺、下修檔還是有它（不會變成沒缺）", row_again["ordered"] == row0["ordered"] and row_again["qty_ship"] == before
          and row_again["status"] == "partial" and row_again["new_qty"] == row0["new_qty"] and res_again["downgrade_rows"] == 1, str(row_again))
    n_logs = len([l for l in c.get("/api/master/logs?q=13000000699901&limit=1000").get_json()["logs"] if l["field"] == "qty_ship"])
    r = c.post("/api/mars/shortage/apply", json={"result": res_again})
    n_logs2 = len([l for l in c.get("/api/master/logs?q=13000000699901&limit=1000").get_json()["logs"] if l["field"] == "qty_ship"])
    check("再按一次確認：系統數量已經是對的所以改 0 筆，但下修檔照樣回來、不重複記歷程", r.status_code == 200 and r.headers.get("X-Mars-Changed") == "0" and r.headers.get("X-Mars-Rows") == "1"
          and n_logs2 == n_logs, f"改 {r.headers.get('X-Mars-Changed')} 筆、{r.headers.get('X-Mars-Rows')} 列、歷程 {n_logs}→{n_logs2}")
    logs = c.get("/api/master/logs?q=13000000699901&limit=1000").get_json()["logs"]
    check("有記歷程：出貨數量、原因缺貨、說明寫勇信出幾箱", any(l["field"] == "qty_ship" and "勇信出" in json.dumps(l, ensure_ascii=False) and "缺貨" in json.dumps(l, ensure_ascii=False) for l in logs), str([l for l in logs if l["field"] == "qty_ship"][:1]))
    v = splits(c, "2026-09-29", "2026-09-30"); tb = next(x for x in v["splits"] if x["id"] == tgt["id"])
    check("已送 EIP 的那份快照跟著改：狀態還是已回填（不是 EIP 送出後訂單有變）、數量是新的", tb["status"] == "filled" and next(i for i in tb["items"] if i["sku_id"] == it0["sku_id"])["qty_ship"] == row0["new_qty"], f"{tb['status']} {tb.get('diff')}")
    # 整張不出：勇信表上根本不會有那張，要人勾「勇信沒出這份」
    for i, x in enumerate([x for x in v["splits"] if not x["eip_po"]]):
        put(x["id"], {"eip_po": f"PO20260931{i}"})
    v = splits(c, "2026-09-29", "2026-09-30")
    ids_all = [x["id"] for x in v["splits"]]
    pdf2 = yx_pdf([("PO202609999", "TAO3", "2026/09/29", [("99999999", 3, "2027/01/01")])])
    r = up(c, "/api/mars/shortage/compare", pdf2, "假2.pdf"); res2 = r.get_json()
    check("表裡沒有我們的單：同倉同日有 EIP 單號的幾份列在「沒在表裡」（一個單號一條）、不判缺貨", res2["summary"]["pos"] == 0 and {i for x in res2["not_in_pdf"] for i in x["split_ids"]} == {x["id"] for x in v["splits"] if x["delivery_date"] == "2026-09-29"}
          and all(not x["declared"] for x in res2["not_in_pdf"]) and len(res2["not_in_pdf"]) == len({x["eip_po"] for x in v["splits"] if x["delivery_date"] == "2026-09-29"}), str(res2["not_in_pdf"]))
    r = up(c, "/api/mars/shortage/compare", pdf2, "假2.pdf", none_ids=",".join(str(i) for i in ids_all)); res3 = r.get_json()
    p0 = res3["pos"][0]
    check("勾過的還留在「沒在表裡」清單、標已勾（才能取消），不會整個框消失", len(res3["not_in_pdf"]) == len(res2["not_in_pdf"]) and all(x["declared"] for x in res3["not_in_pdf"]), str(res3["not_in_pdf"]))
    check("勾「勇信沒出」那張 PO 的每一份 → 整張不出、有訂的品項全缺", res3["summary"]["pos"] == 1 and p0["full"] and all(it["status"] == "none" for y in p0["splits"] for it in y["items"] if it["ordered"]) and not p0["missing_splits"], str(res3["summary"]) + str(p0["missing_splits"]))
    rows = msx.downgrade_rows(res3)
    check("整張不出的下修檔：每個品項都列、第一支 1 其他 0、評論寫「PO此單不出，如有不便，請見諒。」", len(rows) == sum(len(y["items"]) for y in p0["splits"]) and rows[0]["qty"] == 1 and all(r_["qty"] == 0 for r_ in rows[1:])
          and all(r_["comment"] == "13000000699901此單不出，如有不便，請見諒。" for r_ in rows), str(rows[:2]))
    r = c.post("/api/mars/shortage/apply", json={"result": res3})
    in_splits = {i["sku_id"] for x in v["splits"] for i in x["items"]}
    left = [(o["sku_id"], o["qty_ship"]) for o in c.get("/api/master/orders?month=2026-09&q=13000000699901").get_json()["rows"] if o["qty_ship"] and o["sku_id"] in in_splits]
    check("確認整張不出 → 下修檔回來、系統那張 PO 有拆到的品項出貨數量全部改 0（對不到商品總表、沒拆進來的 M99999999 不動）", r.status_code == 200 and not left, f"{r.status_code} {left}")
    check("掃描檔／不是勇信表 → 400 講清楚", up(c, "/api/mars/shortage/compare", b"%PDF-1.4 nothing", "x.pdf").status_code == 400)

    print("\n【8f】盒、包同一張 EIP 採購單（Shanin 2026-10-07：不同料號箱跟盒包分開；同料號箱、盒、包各一張）")
    def row1012(**kw):
        return base_row(**{"PO單號": "13000000583162", "交付日期": datetime.datetime(2026, 10, 12), **kw})
    sheet = special_sheet([
        row1012(**{"SKU ID": "910000000000001", "條碼(國條)": "4710000000001", "永豐料號": "M69072565", "品名": "喵愛餡 盒",
                   "下單數量(酷澎單位)": 24, "出貨數量": 24, "單位": "盒", "箱入數": 12}),
        row1012(**{"SKU ID": "910000000000002", "條碼(國條)": "4710000000002", "永豐料號": "M80676511", "品名": "喵喵鮮 包",
                   "下單數量(酷澎單位)": 48, "出貨數量": 48, "單位": "包", "箱入數": 24}),
        row1012(**{"SKU ID": "910000000000003", "條碼(國條)": "4710000000003", "永豐料號": "M80676511", "品名": "喵喵鮮 箱",
                   "下單數量(酷澎單位)": 3, "出貨數量": 3, "單位": "箱", "箱入數": 1}),
    ])
    check("匯 10/12 那張（PET、都要貼中標：盒、包不同料號，箱跟包同料號）", import_orders(c, sheet, "盒包同單.xlsx").status_code == 200)
    v = splits(c, "2026-10-12"); by_u = {x["unit"]: x for x in v["splits"]}
    check("拆單照舊依單位拆成 3 份", set(by_u) == {"盒", "包", "箱"}, str([x["filename"] for x in v["splits"]]))
    check("畫面標出盒、包同一張 EIP 採購單，箱自己一張；統計 2 張 EIP 採購單",
          [m["unit"] for m in by_u["盒"]["eip_mates"]] == ["包"] and [m["unit"] for m in by_u["包"]["eip_mates"]] == ["盒"]
          and by_u["箱"]["eip_mates"] == [] and v["summary"]["eip_files"] == 2, str(v["summary"]))
    r = c.post("/api/mars/splits/generate", json={"from": "2026-10-12", "to": "2026-10-12"})
    z = zipfile.ZipFile(io.BytesIO(r.data)); names = z.namelist()
    merged = next((n for n in names if "盒+包" in n), None)
    check("下載 EIP 採購單：2 個檔，盒、包合成一個（檔名單位寫「盒+包」）", r.status_code == 200 and len(names) == 2 and merged, str(names))
    eip_of = lambda data: {row[1]: row[6] for row in (xlrd.open_workbook(file_contents=data).sheet_by_index(0).row_values(i) for i in range(1, xlrd.open_workbook(file_contents=data).sheet_by_index(0).nrows)) if row[1]}  # noqa: E731
    check("合併那個檔：盒、包兩個料號各自一列、箱數各 2", eip_of(z.read(merged)) == {"M69072565": 2.0, "M80676511": 2.0}, str(eip_of(z.read(merged))))
    v = splits(c, "2026-10-12"); by_u = {x["unit"]: x for x in v["splits"]}
    r = c.get(f"/api/mars/splits/{by_u['包']['id']}/file?kind=eip")
    check("在包那列單獨下載 EIP：拿到的也是盒+包合在一起的那張", r.status_code == 200 and "盒+包" in unquote(r.headers.get("Content-Disposition", ""))
          and eip_of(r.data) == {"M69072565": 2.0, "M80676511": 2.0})
    r = put(by_u["盒"]["id"], {"eip_po": "PO202610128"}); d = r.get_json()
    v = splits(c, "2026-10-12"); by_u = {x["unit"]: x for x in v["splits"]}
    check("在盒那列填 PO202610128 → 包一起填上、箱不動、不跳提醒", r.status_code == 200 and len(d["mates_updated"]) == 1 and not d["warning"]
          and by_u["包"]["eip_po"] == "PO202610128" and by_u["箱"]["eip_po"] == "", str(d))
    r = put(by_u["箱"]["id"], {"eip_po": "PO202610128"}); d = r.get_json()
    check("箱那份也填同一個號碼 → 可以存，但提醒依規則應分開", r.status_code == 200 and "應分開" in d["warning"], str(d))
    put(by_u["箱"]["id"], {"eip_po": ""})
    v = splits(c, "2026-10-12"); by_u = {x["unit"]: x for x in v["splits"]}
    check("清掉箱的號碼，盒、包不受影響", by_u["箱"]["eip_po"] == "" and by_u["盒"]["eip_po"] == by_u["包"]["eip_po"] == "PO202610128")
    other = next(x for x in splits(c, "2026-09-30")["splits"] if x["id"])
    r = put(other["id"], {"eip_po": "PO202610128"})
    check("別張 PO 填到同一個號碼 → 400，說已經用在哪一份", r.status_code == 400 and "13000000583162" in r.get_json()["error"], str(r.get_json()))
    pdf = yx_pdf([("PO202610128", "TAO3", "2026/10/12", [("69072565", 2, "2027/06/18"), ("80676511", 1, "2027/06/18")])])
    res = up(c, "/api/mars/shortage/compare", pdf, "假_勇信_盒包.pdf").get_json()
    sp = {y["unit"]: y for p_ in res["pos"] for y in p_["splits"]}
    check("勇信比對：同一個號碼的盒、包兩份合起來比 → 盒沒缺、包少 1 箱", set(sp) == {"盒", "包"} and sp["盒"]["items"][0]["status"] == "ok"
          and sp["包"]["items"][0]["status"] == "partial" and sp["包"]["items"][0]["short"] == 1 and not res["unknown"], str(res["summary"]))
    r = put(by_u["包"]["id"], {"eip_po": ""}); v = splits(c, "2026-10-12"); by_u = {x["unit"]: x for x in v["splits"]}
    check("在包那列清掉號碼 → 盒一起清掉", r.status_code == 200 and by_u["盒"]["eip_po"] == "" and by_u["包"]["eip_po"] == "")

    print("\n【9】清除資料")
    app_module._write_users(app_module.get_users(), {"Jerry", "小真"})
    r = c.post("/api/master/reset", json={"confirm": "清空資料", "orders": True, "products": True})
    conn = db.get_conn()
    left = {t: conn.execute(f"SELECT COUNT(*) AS n FROM {t}").fetchone()["n"] for t in ("mst_mars_splits", "mst_mars_products", "mst_mars_uploads")}
    wh_left = conn.execute("SELECT COUNT(*) AS n FROM mst_mars_warehouses").fetchone()["n"]; conn.close()
    check("清訂單＋主檔：瑪氏拆單、商品總表一起清", r.status_code == 200 and not any(left.values()), str(left))
    check("倉庫資料、採購單設定是設定，清資料不清", wh_left == 3 and c.get("/api/mars/po/settings").get_json()["settings"]["holidays"] == ["2026-09-28", "2026-10-09"], str(wh_left))
    r = c.post("/api/mars/splits/generate", json={"from": "2026-09-29", "to": "2026-09-30"})
    check("沒有商品總表時產出 → 400 講清楚", r.status_code == 400 and "商品總表" in r.get_json()["error"])

    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]; sys.exit(1)


if __name__ == "__main__":
    main()
