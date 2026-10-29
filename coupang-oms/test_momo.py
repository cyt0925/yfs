"""MOMO 訂單管理系統（/momo）端到端測試（第一批，2026-10-28）。
讀 MOMO 下載檔（單據首列往下帶、單位代碼、自動分線別、錯誤清單）→ 一次好幾個檔預覽、確認匯入 → 換成箱數、除不盡改數量、
贈品改 0 → 產 EIP 採購單（瑪氏依倉／到貨日／箱或盒包／品類／中標分檔）→ 回填 EIP 單號 → 瑪氏採購單、直送採購單 →
重傳差異、出貨方式、取消產出、右側面板 → ③ 勇信配送明細表核缺貨與效期、填 MOMO 效期匯入檔 → ④ 入庫通知單畫線蓋章 → 首頁卡片、歷程。

測試檔照 MOMO 下載檔的表頭用程式造（真的業務檔不放進 repo）；瑪氏商品總表用 samples/mars/fake 那份假的。
執行：python test_momo.py（設了 DATABASE_URL 就跑 PostgreSQL）
"""
import datetime as _dt
import io
import os
import sys
import tempfile
import zipfile

import openpyxl
import xlrd
import xlwt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

_tmp = tempfile.mkdtemp(prefix="oms_momo_test_")
import db  # noqa: E402
db.DATA_DIR = os.path.join(_tmp, "資料與設定")
db.DB_PATH = os.path.join(db.DATA_DIR, "test.db")
db.BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")

import app as app_module  # noqa: E402
from momo.common import line_of, split_code, wh_canon  # noqa: E402
from momo.parse import ParseError, parse_file  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  ✓ " if cond else "  ✗ ") + name + ("" if cond or not detail else f"  — {detail}"))


HDR = ["項次", "倉別", "到倉日", "採購日", "入庫單號", "入庫預定日期", "確認日期", "採購總量", "延遲訊息", "momo採購備註", "原入庫單號", "列印狀態",
       "課別", "配送溫層", "請採單號", "商品編號", "商品名稱", "單品編號", "單品詳細", "保存天數", "商品有效日期", "入庫最低有效期限", "採購數量",
       "長", "寬", "高", "重量", "商品原廠編號"]


def unsend_rows(docs):
    """docs = [(倉別, 入庫單號, 到貨日 yyyy/mm/dd, [(商品編號, 單品編號, 品名, 數量, 商品原廠編號)], 採購總量 or None)]
    照真檔：單據層欄位只寫在每張單的第一列。數量、編號都存成文字（真檔就是這樣）。"""
    rows = [HDR]
    for n, (wh, no, date, items, total) in enumerate(docs, start=1):
        tot = total if total is not None else sum(int(i[3]) for i in items if str(i[3]).isdigit())
        for k, (prod, item, name, qty, code) in enumerate(items):
            r = dict.fromkeys(HDR, "")
            if k == 0:
                r.update({"項次": str(n), "倉別": wh, "採購日": "2026/10/16", "入庫單號": no, "入庫預定日期": date, "採購總量": str(tot), "原入庫單號": no})
            r.update({"商品編號": prod, "商品名稱": name, "單品編號": item, "單品詳細": "無", "保存天數": "540", "入庫最低有效期限": "2027/06/20",
                      "採購數量": str(qty), "商品原廠編號": code})
            rows.append([r[h] for h in HDR])
    return rows


def unsend_xls(docs):
    wb = xlwt.Workbook(); ws = wb.add_sheet("Sheet1")
    for i, row in enumerate(unsend_rows(docs)):
        for j, v in enumerate(row):
            ws.write(i, j, v)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def unsend_xlsx(docs, extra_first_col=False, sheet="Sheet1"):
    """extra_first_col：照紙品「紙箱購mo回單確認」那樣，前面多一欄查詢公式。"""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = sheet
    for row in unsend_rows(docs):
        ws.append((["查詢區(公式)" if row[0] == "項次" else ""] if extra_first_col else []) + row)
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def price_xlsx(rows):
    """紙潔價格本：[(料號3, 單位, 品名, 箱入數)]，表頭照紙潔價格本整合。"""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "紙品價格本"; ws.append(["料號", "單位", "品名", "料號3", "箱入數"])
    for code3, unit, name, box in rows:
        ws.append([code3[:-1], unit, name, code3, box])
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def yx_pdf(pages):
    """假的勇信配送明細表（文字排法照真檔，跟 test_shopee.yx_pdf 一樣）：pages = [(EIP單號, [(瑪氏貨號, 箱數, 效期) 或 (瑪氏貨號, [(箱數, 效期), ...])])]。"""
    import pymupdf
    doc = pymupdf.open()
    for eip, items in pages:
        lines = ["桃園市觀音區", "BMS100230-永豐商店-momo", "出貨廠商名稱：A5888-台灣瑪氏股份有限公司", "指送日期：2026/10/23",
                 "配 送 明 細 表(預排鮮度)", "產品編號", "產品名稱", "入數", "配送數量", f"收貨單號：{eip}-momo-DC6"]
        total = 0
        for n, it in enumerate(items, start=1):
            batches = it[1] if isinstance(it[1], list) else [(it[1], it[2])]
            lines += [f"58880{it[0]}", str(n), f"假品名 {it[0]} 1:6:10"]
            for q, exp in batches:
                lines += [f"{q}C  ", "625D1TAP01"] + ([exp] if exp else []); total += q
        lines += ["備註:momo嘜頭+入倉單", "單據金額:", f"合計箱數:{total}", "客戶簽章:"]
        page = doc.new_page(); y = 40
        for ln in lines:
            page.insert_text((40, y), ln, fontname="china-t", fontsize=9); y += 14
    return doc.tobytes()


def notice_pdf(no, items):
    """假的 MOMO 入庫通知單，表格畫法照真檔（一格一格的矩形、同樣的欄寬）：items = [(商品編號, 單品編號, 數量)]。"""
    import pymupdf
    doc = pymupdf.open(); page = doc.new_page(width=595, height=842)
    page.insert_text((226, 40), "入庫通知單", fontname="china-t", fontsize=12)
    page.insert_text((63, 75), f"通知單號: {no}", fontname="china-t", fontsize=9)
    xs = [43.5, 109.3, 135.6, 188.3, 254.1, 301.5, 348.8, 383.0, 417.3, 443.6, 469.9, 483.1, 496.2, 509.4, 551.5]
    heads = ["商品編碼", "單品", "原廠編號", "商品名稱", "單品詳細", "保存天數", "最低效期", "分類", "單位", "數量", "註", "備", "檢", "商品有效"]
    def row(y0, y1, texts):
        for k in range(len(xs) - 1):
            page.draw_rect(pymupdf.Rect(xs[k], y0, xs[k + 1], y1), color=(0, 0, 0), width=0.5)
            if texts[k]:
                page.insert_text((xs[k] + 3, y0 + 14), texts[k], fontname="china-t", fontsize=7)
    row(160, 200, heads)
    y = 200
    for prod, item, qty in items:
        row(y, y + 49, [prod, item, "M1", "品名", "", "365", "2027/03/18", "零嘴", "組", str(qty), "", "", "一般", "2027/08/17"]); y += 49
    row(y, y + 16, ["", "", "", "總合", "", "", "", "", "", str(sum(q for _p, _i, q in items)), "", "", "", ""])
    return doc.tobytes()


def up(c, url, files):
    return c.post(url, data={"file": [(io.BytesIO(b), n) for n, b in files]}, content_type="multipart/form-data")


# 瑪氏（假商品總表）：彩豆豆 M89297484 盒 24／箱、需貼中標、Chocolate；口香糖王 M81232885 包 200／箱、不貼中標、Gum；
# 喵愛餡 M69072565 箱 1、需貼中標、Petcare
MARS_A = ("DC6", "26101657-0001", "2026/10/23", [("11000001", "001", "【彩豆豆】牛奶巧克力 40g*12", "48", "M89297484_QMN"),
                                                 ("11000002", "002", "【口香糖王】薄荷", "300", "M81232885_CMN"),
                                                 ("11000003", "001", "【喵愛餡】貓點心 化毛 箱購", "5", "M69072565_RMN"),
                                                 ("15763571", "001", "贈【假】造型地墊", "100", "MARS260905-GIF_TMN")], None)
MARS_B = ("DC1(自)", "26101657-0002", "2026/10/23", [("11000001", "001", "【彩豆豆】牛奶巧克力 40g*12", "24", "M89297484_QMN"),
                                                    ("11000009", "001", "沒有數量的品項", "0", "M89297484_RMN")], None)
PAPER = ("新竹物流幼獅倉", "26102158-0003", "2026/10/24", [("3252331", "001", "【五月花】新柔韌抽取式衛生紙(110抽x72包)", "120", "0102365_RNN"),
                                                       ("15492099", "001", "【五月花】摩力抽廚房紙巾(100抽x1包)", "48", "0104074_CNV")], None)
CLEAN = ("中衛四", "26101657-0004", "2026/10/24", [("11301148", "001", "【橘子工坊】洗衣膠囊 54顆/盒", "60", "1021357_CNN")], None)


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    c = app_module.app.test_client()
    with c.session_transaction() as s:
        s["user"] = "Jerry"

    print("\n【1】讀 MOMO 下載檔（需求二、三、八）")
    check("料號分類：M＝瑪氏、7 碼 0 開頭＝紙品、1 開頭＝潔品、8 碼以上＝寶僑", [line_of(x) for x in ("M89297484", "0102365", "1021357", "4987176232878", "AAAQ1")]
          == ["瑪氏", "紙品", "潔品", "寶僑", ""])
    check("商品原廠編號拆料號和單位代碼（含 -1 的料號）", split_code("M10201240-1_RMN") == ("M10201240-1", "R") and split_code("0104074_CNV") == ("0104074", "C"))
    check("倉別別名只認列出來的：DC1→DC1(自)、CS3→中衛三、北衛二→北衛二(K)，DC99 對不到", [wh_canon(x) for x in ("DC1", "CS3", "北衛二", "DC6", "DC99")]
          == ["DC1(自)", "中衛三", "北衛二(K)", "DC6", ""])
    d = parse_file(unsend_xls([MARS_A, MARS_B]), "UnsendList_1.xls")
    r0 = d["rows"]
    check("單據層欄位（倉別、入庫單號、到貨日）往下帶到同一張單的每一列", [(r["warehouse"], r["inbound_no"], r["expected_date"]) for r in r0[:4]]
          == [("DC6", "26101657-0001", "2026-10-23")] * 4 and r0[4]["warehouse"] == "DC1(自)", str(r0[:2]))
    check("採購數量 0 的不算補貨明細；單品編號保留前面的 0；單位代碼換成中文", len(r0) == 5 and r0[0]["item_no"] == "001" and r0[0]["unit_name"] == "盒"
          and r0[1]["unit_name"] == "包" and r0[3]["unit_name"] == "組", str([r["unit_name"] for r in r0]))
    check("採購總量跟明細加總對得上就沒有警告", d["warnings"] == [] and d["errors"] == [], str(d))
    bad = ("SDC", "26101657-0099", "2026/10/23", [("1", "001", "沒有底線", "5", "M89297484"), ("2", "001", "怪單位", "5", "M89297484_ZMN"),
                                                  ("3", "001", "正常", "6", "M89297484_QMN")], 99)
    d = parse_file(unsend_xls([bad]), "壞.xls")
    check("沒底線、單位代碼不在清單 → 錯誤清單寫檔名、工作表、列號、倉別、原因，不猜", len(d["errors"]) == 2 and d["errors"][0]["row"] == 2
          and d["errors"][0]["warehouse"] == "SDC" and "沒有底線" in d["errors"][0]["reason"] and "Z" in d["errors"][1]["reason"] and len(d["rows"]) == 1, str(d["errors"]))
    check("採購總量跟明細加總對不上 → 警告", d["warnings"] and "對不上" in d["warnings"][0], str(d["warnings"]))
    d = parse_file(unsend_xlsx([PAPER], extra_first_col=True, sheet="紙箱購mo回單確認"), "momo紙品.xlsx")
    check("紙品那份（前面多一欄查詢公式、工作表叫紙箱購mo回單確認）也讀得到：表頭用文字認", len(d["rows"]) == 2 and d["sheet"] == "紙箱購mo回單確認"
          and d["rows"][0]["base_code"] == "0102365" and d["rows"][0]["line"] == "紙品", str(d["rows"][:1]))
    try:
        parse_file(price_xlsx([("0102365R", "箱", "x", 1)]), "價格本.xlsx"); ok = False
    except ParseError as exc:
        ok = "找不到 MOMO 下載檔的表頭" in str(exc)
    check("不是 MOMO 下載檔 → 講清楚", ok)

    print("\n【2】上傳入口認檔、一次好幾個檔預覽、確認匯入")
    r = up(c, "/api/momo/kinds", [("a.xls", unsend_xls([MARS_A])), ("紙潔價格本.xlsx", price_xlsx([("0102365R", "箱", "x", 1)])), ("x.pdf", b"%PDF-1.4")])
    check("認檔：MOMO 下載檔 orders、價格本 conv、PDF 當勇信配送明細表", [f["kind"] for f in r.get_json()["files"]] == ["orders", "conv", "yx"], str(r.get_json()))
    master = os.path.join(BASE_DIR, "samples", "mars", "fake", "假_瑪氏商品總表.xlsx")
    up(c, "/api/mars/products/import", [("假_瑪氏商品總表.xlsx", open(master, "rb").read())])
    up(c, "/api/shopee/conv/import", [("紙潔價格本整合.xlsx", price_xlsx([("0102365R", "箱", "新柔韌", 1), ("0104074C", "包", "摩力抽", 24),
                                                                         ("1021357C", "包", "洗衣膠囊", 6)]))])
    r = up(c, "/api/momo/import/preview", [("UnsendList_1.xls", unsend_xls([MARS_A, MARS_B])), ("UnsendList_2.xlsx", unsend_xlsx([PAPER, CLEAN])),
                                           ("壞.xls", b"not an excel")])
    pv = r.get_json()
    check("三個檔一起丟：兩份讀得到的合在一起預覽、讀不了的那份列在錯誤裡，其他照常", r.status_code == 200 and pv["filenames"] == ["UnsendList_1.xls", "UnsendList_2.xlsx"]
          and pv["counts"]["new"] == 8 and len(pv["orders"]) == 4 and any(e["file"] == "壞.xls" for e in pv["errors"]), str(pv)[:400])
    check("預覽寫各線別幾品、倉別、到貨日", pv["lines"] == {"瑪氏": 5, "紙品": 2, "潔品": 1} and pv["dates"] == ["2026-10-23", "2026-10-24"], str(pv["lines"]))
    pend = c.get("/api/momo/import/pending").get_json()["pending"]
    check("還沒確認：重新整理頁面後預覽放回來", len(pend) == 1 and pend[0]["upload_id"] == pv["upload_id"])
    r = c.post("/api/momo/import/commit", json={"upload_id": pv["upload_id"]})
    check("確認匯入：新增 8 筆", r.status_code == 200 and r.get_json()["counts"]["new"] == 8, str(r.get_json()))
    check("同一份預覽不能匯兩次", c.post("/api/momo/import/commit", json={"upload_id": pv["upload_id"]}).status_code == 404)
    r = up(c, "/api/momo/import/preview", [("again.xls", unsend_xls([MARS_A]))])
    check("同一份再傳一次：全部沒有變，不會變兩筆", r.get_json()["counts"] == {"new": 0, "update": 0, "same": 4, "remove": 0, "changed": 0}, str(r.get_json()["counts"]))
    c.post("/api/momo/import/cancel", json={"upload_id": r.get_json()["upload_id"]})
    r = up(c, "/api/momo/import/preview", [("a.xls", unsend_xls([MARS_B])), ("b.xls", unsend_xls([MARS_B]))])
    check("同一筆在兩個檔裡都有：以後面的為準，預覽講重複幾筆", r.get_json()["dup"] == 1, str(r.get_json()))
    c.post("/api/momo/import/cancel", json={"upload_id": r.get_json()["upload_id"]})

    print("\n【3】訂單總覽、搜尋、篩選")
    o = c.get("/api/momo/orders").get_json()
    st = {s["key"]: s["count"] for s in o["stages"]}
    check("作業流程格子：待產 EIP 4 張、待回填 0、全部 4", st == {"diff": 0, "buy": 4, "eip": 0, "yx": 0, "notice": 0, "all": 4}, str(st))
    g = {x["inbound_no"]: x for x in o["orders"]}
    check("一張入庫單一列：線別、倉別、品項數、MOMO 數量、上傳日", g["26101657-0001"]["lines"] == ["瑪氏"] and g["26101657-0001"]["item_count"] == 4
          and g["26101657-0001"]["qty"] == 453 and g["26101657-0001"]["upload_date"] == db.now()[:10], str(g["26101657-0001"])[:300])
    check("倉別主檔有的倉 wh_known；DC1(自) 也認得", g["26101657-0002"]["wh_known"] and g["26102158-0003"]["wh_known"])
    nos = lambda q: sorted(x["inbound_no"] for x in c.get("/api/momo/orders?" + q).get_json()["orders"])  # noqa: E731
    check("搜尋：入庫單號、商品編號、料號、品名都找得到", nos("q=26102158") == ["26102158-0003"] and nos("q=15492099") == ["26102158-0003"]
          and nos("q=M81232885") == ["26101657-0001"] and nos("q=洗衣膠囊") == ["26101657-0004"])
    check("線別、倉別可以複選", nos("line=紙品,潔品") == ["26101657-0004", "26102158-0003"] and nos("warehouse=DC6,中衛四") == ["26101657-0001", "26101657-0004"])
    check("到貨日篩選", nos("from=2026-10-24") == ["26101657-0004", "26102158-0003"] and nos("to=2026-10-23") == ["26101657-0001", "26101657-0002"])
    check("上傳日篩選", len(nos(f"up_from={db.now()[:10]}")) == 4 and nos("up_to=2020-01-01") == [])

    print("\n【4】換成箱數、除不盡、贈品")
    plan = lambda: {x["inbound_no"]: x for x in c.get("/api/momo/purchase/plan").get_json()["orders"]}  # noqa: E731
    p = plan()
    it = {x["factory_code"]: x for x in p["26101657-0001"]["items"]}
    check("瑪氏 48 盒 ÷ 箱入數 24 ＝ 2 箱；分檔條件：盒包、品類、中標", it["M89297484_QMN"]["out_qty"] == 2 and it["M89297484_QMN"]["unit_group"] == "盒包_CHO_需貼中標")
    check("瑪氏 300 包 ÷ 200 除不盡 → 自動捨去成 1 箱（200 包），不擋、寫一行說明（Jerry 2026-10-29「平常是怎樣那就是怎樣」）",
          not it["M81232885_CMN"]["blocking"] and it["M81232885_CMN"]["out_qty"] == 1 and it["M81232885_CMN"]["use_qty"] == 200 and "捨去" in it["M81232885_CMN"]["note"],
          str(it["M81232885_CMN"]))
    check("贈品商品總表查不到 → 擋下來，提示改成 0", it["MARS260905-GIF_TMN"]["blocking"] and "改成 0" in it["MARS260905-GIF_TMN"]["problems"][-1])
    check("紙品查紙潔價格本：48 包 ÷ 24 ＝ 2 箱；潔品 60 包 ÷ 6 ＝ 10 箱", {x["factory_code"]: x["out_qty"] for x in p["26102158-0003"]["items"]}
          == {"0102365_RNN": 120, "0104074_CNV": 2} and p["26101657-0004"]["items"][0]["out_qty"] == 10)
    gum, gift = it["M81232885_CMN"]["order_id"], it["MARS260905-GIF_TMN"]["order_id"]
    adj = lambda oid, q, reason="測試": c.post("/api/momo/purchase/adjust", json={"order_id": oid, "qty": q, "reason": reason})  # noqa: E731
    check("改的數量還是除不盡 → 擋、告訴可以改成幾", adj(gum, 350).status_code == 400 and "200 的倍數" in adj(gum, 350).get_json()["error"])
    check("沒寫原因 → 擋", adj(gum, 400, "").status_code == 400)
    check("改成 400（2 箱）", adj(gum, 400, "跟業務確認改整箱").status_code == 200)
    check("贈品查不到箱入數也能改成 0（不下採）", adj(gift, 0, "贈品不下採").status_code == 200)
    check("查不到箱入數又要改成 0 以外 → 擋", adj(gift, 5).status_code == 400)
    p = plan()
    check("改完：這張單不再擋；贈品標不下採", p["26101657-0001"]["blocking"] == 0 and next(x for x in p["26101657-0001"]["items"] if x["order_id"] == gift)["skip"], str(p["26101657-0001"]["items"]))

    print("\n【5】產 EIP 採購單")
    r = c.post("/api/momo/purchase/generate", json={"inbound_nos": list(p)})
    check("產出：壓縮檔、幾份 EIP", r.status_code == 200 and r.headers["X-Momo-Files"] == "6", r.get_json() if r.status_code != 200 else r.headers["X-Momo-Files"])
    z = zipfile.ZipFile(io.BytesIO(r.data))
    names = sorted(z.namelist())
    check("瑪氏依倉、到貨日、箱／盒包、品類、中標分檔；紙品、潔品一天一倉一份", names == sorted([
        "產品採購表上傳-MOMO瑪氏_1023到DC6_盒包_CHO_需貼中標.xls", "產品採購表上傳-MOMO瑪氏_1023到DC6_盒包_GUM_不貼中標.xls",
        "產品採購表上傳-MOMO瑪氏_1023到DC6_箱_PET_需貼中標.xls", "產品採購表上傳-MOMO瑪氏_1023到DC1(自)_盒包_CHO_需貼中標.xls",
        "產品採購表上傳-MOMO紙品_1024到新竹物流幼獅倉.xls", "產品採購表上傳-MOMO潔品_1024到中衛四.xls"]), str(names))
    sh = xlrd.open_workbook(file_contents=z.read("產品採購表上傳-MOMO瑪氏_1023到DC6_盒包_GUM_不貼中標.xls")).sheet_by_index(0)
    check("瑪氏 EIP：單位一律箱、數量是箱數（改過的 400 包＝2 箱）、備註 MARS入倉日期送 momo-倉", [sh.cell_value(1, k) for k in (1, 3, 6, 8)]
          == ["M81232885", "箱", 2.0, "MARS入倉20261023台灣瑪氏送momo-DC6"], str([sh.cell_value(1, k) for k in range(10)]))
    sh = xlrd.open_workbook(file_contents=z.read("產品採購表上傳-MOMO紙品_1024到新竹物流幼獅倉.xls")).sheet_by_index(0)
    check("紙品 EIP 用產品領用單、標題「採購數量」接總數、料號前面的 0 保留", sh.name == "產品領用單" and sh.cell_value(0, 6) == "採購數量122"
          and [sh.cell_value(1, k) for k in (1, 3, 6)] == ["0102365", "箱", 120.0] and sh.cell_value(1, 8) == "momo-新竹物流幼獅倉 10/24到貨", str(sh.row_values(0)))
    st = {s["key"]: s["count"] for s in c.get("/api/momo/orders").get_json()["stages"]}
    check("產完：待產 0、待回填 4（贈品不下採不算待產）", st["buy"] == 0 and st["eip"] == 4, str(st))

    print("\n【6】回填 EIP 單號、瑪氏採購單、直送採購單")
    eips = {e["filename"]: e for e in c.get("/api/momo/eips").get_json()["eips"]}
    e1 = eips["產品採購表上傳-MOMO瑪氏_1023到DC6_盒包_CHO_需貼中標.xls"]
    put = lambda e, v: c.put(f"/api/momo/eips/{e['id']}", json={"eip_po": v})  # noqa: E731
    check("單號格式不對 → 擋", put(e1, "PO12").status_code == 400)
    db_conn = db.get_conn()
    db_conn.execute("INSERT INTO shp_eip (line, filename, eip_po, rows_json) VALUES (?,?,?,?)", ("紙潔", "蝦皮的.xls", "PO202610999", "[]")); db_conn.commit(); db_conn.close()
    check("蝦皮特選用過的單號 → 擋（全公司不重複）", put(e1, "PO202610999").status_code == 400)
    check("回填前不能下載瑪氏採購單", c.get(f"/api/momo/eips/{e1['id']}/mars_po").status_code == 400)
    for i, e in enumerate(eips.values()):
        put(e, f"PO2026105{i:02d}")
    st = {s["key"]: s["count"] for s in c.get("/api/momo/orders").get_json()["stages"]}
    check("全部回填：待回填 0", st["eip"] == 0, str(st))
    e1 = next(e for e in c.get("/api/momo/eips").get_json()["eips"] if e["id"] == e1["id"])
    r = c.get(f"/api/momo/eips/{e1['id']}/mars_po")
    ws = openpyxl.load_workbook(io.BytesIO(r.data)).active
    cd = r.headers["Content-Disposition"]
    check("瑪氏採購單檔名：EIP 單號＋永豐Mars採購單_箱_品類_中標_倉_到貨日，盒包那份後面加 (盒包)", "_CHO_" in cd and "DC6_1023" in cd and "%28%E7%9B%92%E5%8C%85%29" in cd, cd)
    check("C3 下單日＝到貨日往前 2 個工作天（10/23 四 → 10/21 二）、F3 配送日、C5 EIP 單號", ws["C3"].value.date() == _dt.date(2026, 10, 21)
          and ws["F3"].value.date() == _dt.date(2026, 10, 23) and ws["C5"].value == e1["eip_po"])
    check("C6 DC6 顯示「momo進貨組-DC6」（Alice 2026-10-29：不加幼獅倉）、F7 ship-to 17617038、C7 地址從倉別主檔、F6 1/2效期以上、B9 品類全名",
          ws["C6"].value == "momo進貨組-DC6" and ws["F7"].value == 17617038 and "高獅路589號" in ws["C7"].value and ws["F6"].value == "1/2效期以上"
          and ws["B9"].value == "Chocolate巧克力", str([ws[k].value for k in ("C6", "F7", "C7", "F6", "B9")]))
    check("F5 作業備註：momo嘜頭+入倉單、需貼中盒標（盒包不貼板嘜）", ws["F5"].value == "momo嘜頭+入倉單\n需貼中盒標", repr(ws["F5"].value))
    check("明細：A 永豐料號、D 單位需求＝MOMO 原數量 48、E 採購數量＝2 箱、H 中標 V", [ws.cell(12, k).value for k in (1, 4, 5, 8)] == ["M89297484", 48, 2, "V"],
          str([ws.cell(12, k).value for k in range(1, 14)]))
    e2 = next(e for e in c.get("/api/momo/eips").get_json()["eips"] if e["filename"].endswith("箱_PET_需貼中標.xls"))
    ws = openpyxl.load_workbook(io.BytesIO(c.get(f"/api/momo/eips/{e2['id']}/mars_po").data)).active
    check("箱單位那份 F5 加貼板嘜；商品總表有採購單箱備註的加「指定品需加工貼小白標」", ws["F5"].value == "momo嘜頭+入倉單\n需貼中盒標\n貼板嘜\n指定品需加工貼小白標", repr(ws["F5"].value))
    ep = next(e for e in c.get("/api/momo/eips").get_json()["eips"] if e["line"] == "紙品")
    check("紙品沒有瑪氏採購單", c.get(f"/api/momo/eips/{ep['id']}/mars_po").status_code == 404)
    r = c.get(f"/api/momo/eips/{ep['id']}/direct_po")
    ws = openpyxl.load_workbook(io.BytesIO(r.data)).active
    check("直送採購單：A1 富邦媒體科技(momo)、D1 EIP 單號、H2 到貨日", ws["A1"].value == "富邦媒體科技(momo)" and ws["D1"].value == ep["eip_po"]
          and ws["H2"].value.date() == _dt.date(2026, 10, 24))
    check("一品一列：採購單ID＝MOMO 入庫單號、供應商料號＝商品原廠編號、數量照 MOMO、轉換箱數公式", [ws.cell(4, k).value for k in (2, 3, 5, 6, 7, 8)]
          == ["26102158-0003", "0102365_RNN", 120, "箱", 1, "=E4/G4"] and [ws.cell(5, k).value for k in (2, 3, 5, 6, 7)] == ["26102158-0003", "0104074_CNV", 48, "包", 24],
          str([ws.cell(4, k).value for k in range(1, 10)]))
    texts = [ws.cell(r_, k).value for r_ in range(18, 26) for k in (1, 2, 3)]
    check("B21 請安排 10/24 到貨、A22 momo-倉、C22 倉別主檔的地址＋入庫時間＋駐倉廠服電話", "請安排10/24 到貨" in texts and "momo-新竹物流幼獅倉" in texts
          and any(isinstance(t, str) and "梅獅路359-1號" in t and "momo進貨組" in t and "商品入庫時間" in t and "02-6600-7600" in t for t in texts), str(texts))
    r = c.get("/api/momo/po.zip?kind=mars")
    check("已回填的瑪氏採購單一次下載", r.status_code == 200 and r.headers["X-Momo-Files"] == "4")
    check("已回填的直送採購單一次下載（紙品、潔品）", c.get("/api/momo/po.zip?kind=direct").headers.get("X-Momo-Files") == "2")

    print("\n【7】取消產出、出貨方式")
    check("已回填的 EIP 不能取消產出", c.delete(f"/api/momo/eips/{ep['id']}").status_code == 400)
    put(ep, "")
    r = c.delete(f"/api/momo/eips/{ep['id']}")
    st = {s["key"]: s["count"] for s in c.get("/api/momo/orders").get_json()["stages"]}
    check("清掉單號後取消產出：那張入庫單退回待產 EIP", r.status_code == 200 and r.get_json()["orders"] == ["26102158-0003"] and st["buy"] == 1, str(st))
    r = c.post("/api/momo/ship_method", json={"inbound_nos": ["26101657-0004"], "ship_method": "竹運出貨"})
    check("已產出 EIP 的單不能改出貨方式（要先取消產出）", r.status_code == 400 and "取消產出" in r.get_json()["error"])
    r = c.post("/api/momo/ship_method", json={"inbound_nos": ["26102158-0003"], "ship_method": "竹運出貨"})
    o = {x["inbound_no"]: x for x in c.get("/api/momo/orders").get_json()["orders"]}
    check("改成竹運出貨：不用產 EIP（待產 0）", r.status_code == 200 and o["26102158-0003"]["ship_method"] == "竹運出貨"
          and o["26102158-0003"]["po_status"] == "不需產出" and not o["26102158-0003"]["todo"], str(o["26102158-0003"]["todo"]))
    c.post("/api/momo/ship_method", json={"inbound_nos": ["26102158-0003"], "ship_method": "原廠直出"})

    print("\n【8】MOMO 重新下載、內容有變")
    changed = ("DC1(自)", "26101657-0002", "2026/10/25", [("11000001", "001", "【彩豆豆】牛奶巧克力 40g*12", "48", "M89297484_QMN")], None)
    r = up(c, "/api/momo/import/preview", [("UnsendList_3.xls", unsend_xls([changed]))])
    pv = r.get_json()
    check("預覽列出改了什麼：數量 24→48、到貨日", pv["counts"]["changed"] == 1 and {d_["field"] for d_ in pv["diffs"]} == {"qty", "expected_date"}, str(pv["diffs"]))
    c.post("/api/momo/import/commit", json={"upload_id": pv["upload_id"]})
    o = c.get("/api/momo/orders").get_json()
    st = {s["key"]: s["count"] for s in o["stages"]}
    check("那張單差異待確認", st["diff"] == 1 and next(x for x in o["orders"] if x["inbound_no"] == "26101657-0002")["diff_pending"], str(st))
    side = c.get("/api/momo/order?no=26101657-0002").get_json()
    check("右側面板：差異、品項計畫、EIP、歷程、倉別資料都在", {"order", "plan", "eips", "diffs", "logs", "warehouse"} <= set(side)
          and len([x for x in side["diffs"] if not x["confirmed"]]) == 2 and "高鐵站前西路" in side["warehouse"]["note"] and side["logs"], str(list(side)))
    check("差異沒確認：品項擋下來", any("MOMO 改過" in t for x in side["plan"]["items"] for t in x["problems"]))
    c.post("/api/momo/diffs/confirm", json={"inbound_nos": ["26101657-0002"]})
    st = {s["key"]: s["count"] for s in c.get("/api/momo/orders").get_json()["stages"]}
    check("確認後差異 0", st["diff"] == 0, str(st))
    gone = ("中衛四", "26101657-0004", "2026/10/24", [("99999999", "001", "換了別的", "6", "1021357_CNN")], None)
    pv = up(c, "/api/momo/import/preview", [("UnsendList_4.xls", unsend_xls([gone]))]).get_json()
    check("同一張入庫單少掉的品項算移除", pv["counts"]["remove"] == 1 and pv["counts"]["new"] == 1, str(pv["counts"]))
    c.post("/api/momo/import/cancel", json={"upload_id": pv["upload_id"]})

    print("\n【9】③ 待核缺貨與效期：勇信配送明細表、效期匯入檔")
    conn = db.get_conn()
    eip_of = {}
    for e in conn.execute("SELECT * FROM momo_eip WHERE line = '瑪氏'").fetchall():
        e = dict(e)
        for oid in [x["order_id"] for x in conn.execute("SELECT order_id FROM momo_eip_items WHERE eip_id = ?", (e["id"],)).fetchall()]:
            o = dict(conn.execute("SELECT inbound_no, base_code FROM momo_orders WHERE id = ?", (oid,)).fetchone())
            eip_of[(o["inbound_no"], o["base_code"])] = e["eip_po"]
    conn.close()
    o = c.get("/api/momo/orders").get_json()
    st = {s["key"]: s["count"] for s in o["stages"]}
    g = {x["inbound_no"]: x for x in o["orders"]}
    check("瑪氏回填過 EIP 的單進「待核缺貨與效期」（勇信表還沒來）；紙品、潔品不走這一步", st["yx"] == 2 and g["26101657-0001"]["check"]["wait"] == 3
          and g["26101657-0004"]["check"] is None and "yx" not in g["26101657-0004"]["todo"], str(st))
    choc, gum, pet = eip_of[("26101657-0001", "M89297484")], eip_of[("26101657-0001", "M81232885")], eip_of[("26101657-0001", "M69072565")]
    choc2 = eip_of[("26101657-0002", "M89297484")]
    pdf1 = yx_pdf([(choc, [("89297484", [(1, "2028/02/08"), (1, "2028/02/09")])]), (gum, [("81232885", 0, "")])])
    pdf2 = yx_pdf([(pet, [("69072565", [(3, "2028/03/11"), (2, "2028/03/12")])]), (choc2, [("89297484", 2, "2028/02/08")]),
                   ("PO202699999", [("89297484", 5, "2028/02/08")])])
    r = up(c, "/api/momo/yx", [("勇信1.pdf", pdf1), ("勇信2.pdf", pdf2)])
    d = r.get_json()
    check("一次丟兩份勇信表：照收貨單號對回 MOMO 的 EIP；不是 MOMO 的單號列出來", r.status_code == 200 and d["unknown"] == ["PO202699999"]
          and sorted(x["inbound_no"] for x in d["orders"]) == ["26101657-0001", "26101657-0002"], str(d)[:400])
    ck = lambda no: {x["factory_code"]: x for x in c.get(f"/api/momo/check?no={no}").get_json()["orders"][0]["items"]}  # noqa: E731
    it = ck("26101657-0001")
    check("勇信 0 箱 → 不出貨（先勾好）", it["M81232885_CMN"]["decision"] == "noship" and it["M81232885_CMN"]["status"] == "none", str(it["M81232885_CMN"]))
    check("兩批效期取數量多的那批（3 箱 2028/03/11 > 2 箱）→ 寫入效期", it["M69072565_RMN"]["decision"] == "ship" and it["M69072565_RMN"]["expiry"] == "2028-03-11",
          str(it["M69072565_RMN"]))
    check("兩批一樣多 → 要人選", it["M89297484_QMN"]["decision"] == "pending" and it["M89297484_QMN"]["tie"], str(it["M89297484_QMN"]))
    check("贈品（改成 0 不下採）→ 不出貨", it["MARS260905-GIF_TMN"]["decision"] == "noship" and it["MARS260905-GIF_TMN"]["status"] == "skip")
    it2 = ck("26101657-0002")
    check("勇信出的比下的多 → 要人決定", it2["M89297484_QMN"]["decision"] == "pending" and it2["M89297484_QMN"]["status"] == "over", str(it2["M89297484_QMN"]))
    choc_id = it["M89297484_QMN"]["order_id"]
    r = c.post("/api/momo/check", json={"order_id": choc_id, "expiry": "2027-06-20"})
    check("人選的效期沒有晚於 MOMO 最低效期（剛好同一天）→ 還是要人看", r.status_code == 200 and ck("26101657-0001")["M89297484_QMN"]["decision"] == "pending"
          and "沒有晚於" in ck("26101657-0001")["M89297484_QMN"]["why"])
    up_ = [("11000003", "001", "【喵愛餡】貓點心 化毛 箱購", "5", "M69072565_RMN"), ("15763571", "001", "贈【假】造型地墊", "100", "MARS260905-GIF_TMN"),
           ("11000002", "002", "【口香糖王】薄荷", "300", "M81232885_CMN"), ("11000001", "001", "【彩豆豆】牛奶巧克力 40g*12", "48", "M89297484_QMN")]
    exp_xls = unsend_xls([("DC6", "26101657-0001", "2026/10/23", up_, None)])
    r = up(c, "/api/momo/kinds", [("UnsendList_效期.xls", exp_xls)])
    check("上傳入口認得效期匯入檔（裡面的單都核過勇信表）", r.get_json()["files"][0]["kind"] == "expfill", str(r.get_json()))
    r = up(c, "/api/momo/expfill", [("UnsendList_效期.xls", exp_xls)])
    check("還有要人決定的品項 → 不產，列出來", r.status_code == 400 and any("M89297484_QMN" in x for x in r.get_json()["details"]), str(r.get_json()))
    c.post("/api/momo/check", json={"order_id": choc_id, "expiry": "2028/02/09", "reason": "跟勇信確認"})
    check("人選效期 2028/02/09 → 寫入效期", ck("26101657-0001")["M89297484_QMN"]["decision"] == "ship" and ck("26101657-0001")["M89297484_QMN"]["override"])
    r = up(c, "/api/momo/expfill", [("UnsendList_效期.xls", exp_xls)])
    sh = xlrd.open_workbook(file_contents=r.data).sheet_by_index(0) if r.status_code == 200 else None
    u = HDR.index("商品有效日期")
    got = {sh.cell_value(k, HDR.index("商品原廠編號")): sh.cell_value(k, u) for k in range(1, 5)} if sh else {}
    check("填好效期：照入庫單號＋商品編號＋單品編號找（列的順序跟系統不一樣也對）；不出貨的留白", got == {"M69072565_RMN": "2028/03/11", "MARS260905-GIF_TMN": "",
          "M81232885_CMN": "", "M89297484_QMN": "2028/02/09"} and r.headers.get("X-Momo-Filled") == "2", str(got) if sh else str(r.get_json()))
    check("檔名後面加「_已填效期」", "_%E5%B7%B2%E5%A1%AB%E6%95%88%E6%9C%9F.xls" in r.headers.get("Content-Disposition", ""), r.headers.get("Content-Disposition"))
    r = up(c, "/api/momo/expfill", [("UnsendList_效期.xlsx", unsend_xlsx([("DC6", "26101657-0001", "2026/10/23", up_, None)]))])
    ws = openpyxl.load_workbook(io.BytesIO(r.data)).active if r.status_code == 200 else None
    check(".xlsx 的效期匯入檔也能填", ws is not None and ws.cell(2, u + 1).value == "2028/03/11" and ws.cell(5, u + 1).value == "2028/02/09")
    bad = unsend_xls([("DC7", "26101657-0001", "2026/10/23", up_, None)])
    r = up(c, "/api/momo/expfill", [("錯倉.xls", bad)])
    check("倉別跟系統不一樣 → 擋", r.status_code == 400 and any("倉別" in x for x in r.get_json()["details"]), str(r.get_json()))
    o = c.get("/api/momo/orders").get_json()
    st = {s["key"]: s["count"] for s in o["stages"]}
    g = {x["inbound_no"]: x for x in o["orders"]}
    check("填完效期：這張離開「待核缺貨與效期」、有不出貨的進「待畫入庫通知單」", "yx" not in g["26101657-0001"]["todo"] and "notice" in g["26101657-0001"]["todo"]
          and g["26101657-0001"]["check"]["done"] and st["yx"] == 1, str(st))
    oid2 = it2["M89297484_QMN"]["order_id"]
    c.post("/api/momo/check", json={"order_id": oid2, "ship": 1, "reason": "勇信多出的照收"})
    check("要人決定的品項改成出貨", ck("26101657-0002")["M89297484_QMN"]["decision"] == "ship")
    c.post("/api/momo/check", json={"order_id": oid2})
    check("清掉人工判斷 → 回到系統判斷", ck("26101657-0002")["M89297484_QMN"]["decision"] == "pending")
    r = up(c, "/api/momo/yx", [("勇信3.pdf", yx_pdf([(choc2, [("89297484", 1, "2028/02/08")])]))])
    check("同一份 EIP 勇信表再傳：換掉上一次", ck("26101657-0002")["M89297484_QMN"]["decision"] == "ship", str(ck("26101657-0002")["M89297484_QMN"]))
    side = c.get("/api/momo/order?no=26101657-0001").get_json()
    check("右側面板有缺貨與效期、歷程記到勇信表和人工判斷", side["check"] and len(side["check"]["items"]) == 4
          and {"momo_yx", "momo_check", "momo_expfill"} <= {x["field"] for x in side["logs"]}, str([x["field"] for x in side["logs"]]))

    print("\n【9b】④ 待畫入庫通知單：入庫通知單 PDF 照列畫線、畫掉數量寫實際數、蓋章")
    o = c.get("/api/momo/orders").get_json()
    st = {s["key"]: s["count"] for s in o["stages"]}
    g = {x["inbound_no"]: x for x in o["orders"]}
    check("③ 做完、有不出貨的單進「待畫入庫通知單」；26101657-0002 效期還沒填不算", st["notice"] == 1 and "notice" in g["26101657-0001"]["todo"]
          and g["26101657-0001"]["check"]["notice_need"] and not g["26101657-0001"]["check"]["notice"], str(st))
    r = up(c, "/api/momo/kinds", [("入庫通知單.pdf", notice_pdf("26101657-0001", [("11000001", "001", 48)]))])
    check("上傳入口認得入庫通知單（跟勇信表都是 PDF）", r.get_json()["files"][0]["kind"] == "notice", str(r.get_json()))
    check("還沒上傳印章", c.get("/api/momo/stamp").get_json()["has"] is False and c.get("/api/momo/stamp.png").status_code == 404)
    pdf = notice_pdf("26101657-0001", [("11000001", "001", 48), ("11000002", "002", 300), ("11000003", "001", 5), ("15763571", "001", 100)])
    r = up(c, "/api/momo/notice", [("26101657-0001入庫通知單.pdf", pdf), ("別的.pdf", notice_pdf("26101657-9999", [("99999999", "001", 1)]))])
    d = r.get_json()
    check("入庫通知單收進來：對回入庫單、列出要畫什麼（口香糖、贈品不出貨；彩豆豆出滿不用畫）；不在系統的那張列在錯誤", r.status_code == 200
          and d["orders"][0]["inbound_no"] == "26101657-0001" and d["orders"][0]["rows"] == 4 and d["orders"][0]["noship"] == 2 and d["orders"][0]["short"] == 0
          and d["orders"][0]["ready"] and any("26101657-9999" in e for e in d["errors"]), str(d)[:500])
    r = c.get("/api/momo/notice/26101657-0001/file")
    check("沒上傳印章也能畫（紅框代替）：畫了 2 列、檔名 _已畫.pdf", r.status_code == 200 and r.headers["X-Momo-Drawn"] == "2" and r.headers["X-Momo-Stamp"] == "0"
          and "_%E5%B7%B2%E7%95%AB.pdf" in r.headers.get("Content-Disposition", ""), r.get_json() if r.status_code != 200 else r.headers.get("Content-Disposition"))
    import pymupdf
    pg0 = pymupdf.open(stream=r.data, filetype="pdf")[0]
    long_ = [it for dr in pg0.get_drawings() for it in dr["items"] if it[0] == "l" and abs(it[1].x - it[2].x) > 400]
    check("不出貨兩列各畫一條橫跨整列的線、紅框寫業務專用章", len(long_) == 2 and pg0.get_text().count("業務專用章") == 2, f"{len(long_)} {pg0.get_text()[-80:]!r}")
    o = c.get("/api/momo/orders").get_json()
    g = {x["inbound_no"]: x for x in o["orders"]}
    check("下載過就算畫好：離開「待畫入庫通知單」，記誰幾點畫的", "notice" not in g["26101657-0001"]["todo"] and g["26101657-0001"]["check"]["notice_done"]
          and g["26101657-0001"]["check"]["notice"]["drawn_by"] == "Jerry")
    from PIL import Image
    im = Image.new("RGB", (60, 40), "white"); im.paste((200, 0, 0), (10, 10, 50, 30)); bs = io.BytesIO(); im.save(bs, "JPEG")
    r = up(c, "/api/momo/stamp", [("章.jpg", bs.getvalue())])
    check("上傳印章：白底去掉、存起來、能取回", r.status_code == 200 and c.get("/api/momo/stamp").get_json()["has"] and c.get("/api/momo/stamp.png").mimetype == "image/png")
    import time as _time; _time.sleep(1.1)              # 時間戳到秒：勇信表要比剛剛畫的「晚」才算重傳
    r = up(c, "/api/momo/yx", [("勇信4.pdf", yx_pdf([(choc, [("89297484", 1, "2028/02/08")])]))])
    o = c.get("/api/momo/orders").get_json()
    g = {x["inbound_no"]: x for x in o["orders"]}
    check("勇信表重傳（彩豆豆只出 1 箱）→ 畫過的不算了、效期檔要重填、彩豆豆要人決定", not g["26101657-0001"]["check"]["notice_done"]
          and not g["26101657-0001"]["check"]["filled_at"] and g["26101657-0001"]["check"]["pending"] == 1, str(g["26101657-0001"]["check"]))
    check("還有要決定的 → 不能畫", c.get("/api/momo/notice/26101657-0001/file").status_code == 400)
    c.post("/api/momo/check", json={"order_id": choc_id, "ship": 1, "reason": "照勇信出的"})
    r = c.get("/api/momo/notice/26101657-0001/file")
    pg0 = pymupdf.open(stream=r.data, filetype="pdf")[0]
    check("彩豆豆改成出 1 箱＝24 盒 → 短出：畫 3 列、數量畫掉寫 24、蓋章用圖片（3 個）、不再畫紅框", r.status_code == 200 and r.headers["X-Momo-Drawn"] == "3"
          and r.headers["X-Momo-Stamp"] == "1" and "24" in pg0.get_text() and len(pg0.get_images()) >= 3 and "業務專用章" not in pg0.get_text(),
          f"{r.headers.get('X-Momo-Drawn')} {len(pg0.get_images())}")
    check("不在待核的單：沒有東西可以畫", c.get("/api/momo/notice/26102158-0003/file").status_code == 404)
    side = c.get("/api/momo/order?no=26101657-0001").get_json()
    check("右側面板的核對結果有實際出貨數、箱入數、入庫通知單狀態", all("actual" in x and "conv" in x for x in side["check"]["items"]) and side["check"]["summary"]["notice"]
          and {"momo_notice", "momo_notice_draw"} <= {x["field"] for x in side["logs"]}, str([x["field"] for x in side["logs"]]))

    print("\n【10】首頁、歷程、頁面")
    sm = c.get("/api/portal/summary").get_json()
    ch = {x["key"]: x for x in sm["channels"]}
    check("首頁 MOMO 卡片上線：有今日新進、連到 /momo", ch["momo"]["live"] and ch["momo"]["url"] == "/momo" and ch["momo"]["new_today"] == 4, str(ch["momo"])[:200])
    logs = c.get("/api/master/logs?scope=momo&limit=500").get_json()
    rows = logs["logs"] if isinstance(logs, dict) else logs
    check("歷程「本頁相關」只有 MOMO 的紀錄", rows and all(x["field"].startswith("momo_") for x in rows), str(rows[:1]))
    mars_logs = c.get("/api/master/logs?scope=mars&limit=500").get_json()
    mrows = mars_logs["logs"] if isinstance(mars_logs, dict) else mars_logs
    check("瑪氏出貨頁的歷程不會混進 MOMO 的瑪氏單", not any(x["field"].startswith("momo_") for x in mrows))
    html = c.get("/momo").get_data(as_text=True)
    check("頁面：上傳區、作業流程、搜尋、右側面板", 'id="dz"' in html and 'id="stages"' in html and 'id="f-q"' in html and 'id="side"' in html and "momo.js" in html)
    whs = c.get("/api/momo/warehouses").get_json()["warehouses"]
    check("倉別主檔 19 個倉", len(whs) == 19 and any(w["name"] == "DC6" and w["ship_to"] == "17617038" for w in whs))
    r = c.put("/api/momo/warehouses/DC4", json={"contact": "駐倉廠服"})
    check("倉別主檔可以改、記歷程", r.status_code == 200 and next(w for w in c.get("/api/momo/warehouses").get_json()["warehouses"] if w["name"] == "DC4")["contact"] == "駐倉廠服")

    print("\n" + "=" * 62)
    print(f"通過 {len(PASS)} 項／失敗 {len(FAIL)} 項")
    if FAIL:
        for n in FAIL:
            print("  ✗", n)
        return 1
    print("全部通過。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
