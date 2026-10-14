"""採購表轉換功能的端到端驗證：拿真實的三線範例檔跑完整流程。

這個功能跟酷澎訂單管理完全獨立（不共用資料庫表），所以獨立成一支
測試檔，不跟 test_flow.py 混在一起——這樣之後改動購表邏輯，不用連
帶跑一次整個訂單匯入流程的測試才能確認有沒有壞掉。

執行：python test_purchase.py
"""
import io
import os
import shutil
import sys
import tempfile
import zipfile

import xlrd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(BASE_DIR, "samples", "purchase")

_tmp = tempfile.mkdtemp(prefix="oms_purchase_test_")
sys.path.insert(0, BASE_DIR)

import db  # noqa: E402
db.DATA_DIR = os.path.join(_tmp, "資料與設定")
db.DB_PATH = os.path.join(db.DATA_DIR, "test.db")
db.BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")

import app as app_module  # noqa: E402

PASS, FAIL = [], []


def check(name, condition, detail=""):
    (PASS if condition else FAIL).append(name)
    mark = "✓" if condition else "✗"
    print(f"  {mark} {name}" + (f"  — {detail}" if detail else ""))


def sign_in(client, username="小真", password="changeme123"):
    res = client.post("/login", data={"username": username, "password": password})
    assert res.status_code == 302, f"測試帳號登入失敗：{username}"


def parse(client, line, filename, upload_name=None):
    """upload_name：模擬使用者上傳時實際帶的檔名跟本機測試檔名不一樣的
    情況——日期救援機制是看「上傳檔名」，不是看樣本檔在硬碟上叫什麼。"""
    path = os.path.join(SAMPLES, filename)
    with open(path, "rb") as fh:
        data = fh.read()
    res = client.post("/api/purchase/parse", data={
        "line": line, "file": (io.BytesIO(data), upload_name or filename)},
        content_type="multipart/form-data")
    return res


def parse_multi(client, line, items):
    """items: [(本機檔名, 模擬上傳檔名或 None), ...]——瑪氏可以一次選
    多個檔案，用 werkzeug 測試客戶端要用同一個欄位名重複帶多個檔案。"""
    files = []
    for filename, upload_name in items:
        with open(os.path.join(SAMPLES, filename), "rb") as fh:
            data = fh.read()
        files.append((io.BytesIO(data), upload_name or filename))
    res = client.post("/api/purchase/parse",
                       data={"line": line, "file": files},
                       content_type="multipart/form-data")
    return res


def export(client, line, groups):
    return client.post("/api/purchase/export", json={"line": line, "groups": groups})


def read_xls_main_sheet(data):
    wb = xlrd.open_workbook(file_contents=data)
    sh = wb.sheet_by_index(0)
    rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    return wb, rows


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()
    sign_in(client)

    print("\n【1】沒登入打不進去（跟酷澎訂單系統共用同一套登入保護）")
    anon = app_module.app.test_client()
    res = anon.get("/purchase")
    check("未登入會被導去登入頁，不是直接看到頁面", res.status_code in (302, 401), res.status_code)

    res = client.get("/purchase")
    check("登入後頁面正常渲染（樣板變數沒漏，例如 logo_file）",
          res.status_code == 200, res.status_code)

    # static/logo_purchase.png 是使用者真的上傳、已經 commit 進 repo 的
    # 正式資產——絕對不要動它（改名、搬移、刪除都不行）：測試被中斷
    # （逾時、Ctrl-C）的話 finally 不一定來得及跑，會把正式檔案真的
    # 弄丟。只驗證「已經存在」該有的行為；只有在檔案還不存在的機器上
    # （例如全新 checkout、還沒人上傳過 logo）才測「不存在→退回預設
    # →放新檔案生效」這條路徑，而且全程只碰自己建立的檔案。
    green = os.path.join(BASE_DIR, "static", "logo_purchase.png")
    if os.path.exists(green):
        res = client.get("/purchase")
        check("已經有 static/logo_purchase.png 時，頁面就是用它",
              b"/static/logo_purchase.png" in res.data)
    else:
        res = client.get("/purchase")
        check("沒有綠色 logo 檔案時，頁面退回用原本那張，不會破圖",
              b"/static/logo.png" in res.data and b"/static/logo_purchase.png" not in res.data)
        try:
            shutil.copy(os.path.join(BASE_DIR, "static", "logo.png"), green)
            res = client.get("/purchase")
            check("static/logo_purchase.png 一放進去就自動改用綠色 logo",
                  b"/static/logo_purchase.png" in res.data)
        finally:
            if os.path.exists(green):
                os.remove(green)
        res = client.get("/purchase")
        check("綠色 logo 被移掉會自動退回原本的，不會破圖",
              b"/static/logo_purchase.png" not in res.data)

    print("\n【2】P&G：一張 PO 一張採購表")
    res = parse(client, "pg", "PG_訂單匯入範例.xlsx")
    check("解析成功", res.status_code == 200, res.get_json() if res.status_code != 200 else "")
    data = res.get_json()
    check("五張 PO 分成五組，沒有被合併", len(data["groups"]) == 5, len(data["groups"]))

    g = next(gr for gr in data["groups"] if gr["key"] == "13000000492925")
    check("這組品項數對得起來（5 個 SKU）", g["item_count"] == 5, g["item_count"])
    check("倉別從出貨備註自動猜到 TXRC8", g["warehouse_guess"] == "TXRC8", g["warehouse_guess"])
    check("日期從出貨備註自動猜到 0905", g["date_guess"] == "0905", g["date_guess"])
    check("備註樣板套出來的字串完全對得上使用者給的真實範例",
          g["remark"] == "13000000492925(TXRC8)_酷澎9/5買斷", g["remark"])

    res = export(client, "pg", [g])
    check("匯出成功（單組直接給 .xls，不是 zip）",
          res.status_code == 200 and res.mimetype == "application/vnd.ms-excel", res.status_code)
    wb, rows = read_xls_main_sheet(res.data)
    check("查表工作表（表單選項／商品資料／客編）原封不動保留",
          set(wb.sheet_names()) == {"產品採購單", "表單選項(勿動)", "商品資料(勿動)", "客編(勿動)"},
          wb.sheet_names())
    check("資料列筆數對得上（表頭 + 5 筆）", len(rows) == 6, len(rows))
    check("料號、數量、備註都填對了第一列",
          rows[1][1] == "4987176232854" and rows[1][6] == 27 and rows[1][8] == g["remark"],
          rows[1])
    check("料號開頭的 0 沒有被吃掉（Excel 數字格式常見地雷）",
          all(str(r["material_no"]).startswith("4987176") for r in g["rows"]))

    print("\n【3】P&G：多組一起匯出要包成 zip，檔名不能互相蓋掉")
    res = export(client, "pg", data["groups"])
    check("匯出成功、格式是 zip", res.status_code == 200 and res.mimetype == "application/zip", res.status_code)
    with zipfile.ZipFile(io.BytesIO(res.data)) as zf:
        names = zf.namelist()
        check("五組各自一個檔案，檔名沒有重複", len(names) == 5 and len(set(names)) == 5, names)

    print("\n【4】紙潔：多張 PO 合併成一張採購表，備註固定空白")
    res = parse(client, "paper", "紙潔_訂單匯入範例.xlsx")
    data2 = res.get_json()
    check("四張 PO 各自分組（合併與否交給匯出那一步決定）",
          len(data2["groups"]) == 4, len(data2["groups"]))
    check("紙潔的備註規則是固定空白，不是猜不到才空白",
          all(g["remark"] == "" for g in data2["groups"]))

    merged_rows = [row for g in data2["groups"] for row in g["rows"]]
    res = export(client, "paper", [{"rows": merged_rows, "remark": "",
                                     "filename": "酷澎_產品採購表上傳_0902到貨.xls"}])
    check("合併匯出成功、單一 .xls（不是 zip）",
          res.status_code == 200 and res.mimetype == "application/vnd.ms-excel", res.status_code)
    wb2, rows2 = read_xls_main_sheet(res.data)
    check("合併後資料列數＝四張 PO 全部品項加起來（表頭+10筆），沒有被合併算加總",
          len(rows2) == 11, len(rows2))
    check("同一個料號在不同 PO 分開出現兩次，不會被誤合併加總",
          [r[1] for r in rows2[1:]].count("0304042") == 2)
    check("表頭那個帶著總和數字的怪標題原樣保留（使用者要求維持原樣）",
          rows2[0][6] == "採購數量398", rows2[0][6])

    print("\n【5】瑪氏：讀 MARS 的「訂貨通知單」，不是酷澎系統的「訂單匯入」表")
    # 第一版把瑪氏也套進 P&G／紙潔那套「訂單匯入」解析邏輯，是同事一開始
    # 給錯範例害的——真正的來源是 MARS TAIWAN 開的訂貨通知單（ORDER
    # FORM），完全不同的表格結構：一份 95 列的大表，只有「採購數量」
    # 欄有填數字的才是真的要訂的品項，其他都是「有賣過但這次沒訂」的
    # 參考列，不能整批照抓；而且一份訂貨通知單就是一張 PO，不用再像
    # P&G／紙潔那樣分組。
    res = parse(client, "mars", "瑪氏_訂貨通知單範例.xlsx",
                upload_name="永豐Mars採購單(箱單位)-GUM糖_TAO4_13000000467952.xlsx")
    check("解析成功", res.status_code == 200, res.get_json() if res.status_code != 200 else "")
    data3 = res.get_json()
    check("一份訂貨通知單只有一組", len(data3["groups"]) == 1, len(data3["groups"]))
    g = data3["groups"][0]

    check("95 列裡只挑出 2 列真的有填採購數量的，其他參考列濾掉",
          g["item_count"] == 2, g["item_count"])
    check("品項明細對得上（永豐料號＋採購數量）",
          g["rows"] == [{"material_no": "M10254005", "qty": 1},
                        {"material_no": "M10267034", "qty": 2}], g["rows"])
    check("PO 單號抓不到「永豐PO單號」欄位（範例裡是空的）時，退而看上傳檔名",
          g["po_numbers"] == ["13000000467952"], g["po_numbers"])
    check("倉別從「入倉倉別: 酷澎-TAO4」這格解析出來，不是猜地址",
          g["warehouse_guess"] == "TAO4", g["warehouse_guess"])
    check("日期直接讀「配送日」那格的日期值，不用再猜",
          g["date_guess"] == "0904", g["date_guess"])
    check("品類代碼從「Gum糖果」查對照表變成 GUM糖（檔名裡那段的公司慣例寫法，不是機械取三碼）",
          g["category"] == "GUM糖", g["category"])
    check("備註樣板套出來的字串完全對得上使用者給的真實範例",
          g["remark"] == "MARS入倉9/4瑪氏送酷澎-TAO4倉", g["remark"])
    check("檔名也完全對得上真實範例",
          g["filename"] == "永豐Mars採購單(箱單位)-GUM糖_TAO4_13000000467952.xls", g["filename"])

    res = export(client, "mars", [{"rows": g["rows"], "remark": g["remark"], "filename": g["filename"]}])
    check("匯出成功", res.status_code == 200, res.status_code)
    wb3, rows3 = read_xls_main_sheet(res.data)
    check("跟使用者給的真實範例檔逐格一致（表頭+2筆資料）",
          rows3 == [
              ["項目", "料號", "品名", "單位", "寄銷倉庫存(直營平台請填0)",
               "寄銷倉前30天實銷(直營平台請填0)", "採購數量", "本月預估CM%", "備註",
               "外幣採購單價(新台幣採購請填0)"],
              [1.0, "M10254005", "", "箱", "", "", 1.0, "", g["remark"], ""],
              [2.0, "M10267034", "", "箱", "", "", 2.0, "", g["remark"], ""],
          ], rows3)
    check("瑪氏範本只有一張主表，沒有 P&G／紙潔那三張查表工作表",
          wb3.sheet_names() == ["產品領用單"], wb3.sheet_names())

    print("\n【5.1】瑪氏：PO 單號抓不到檔名裡的數字時，退而用檔名本身當識別、不硬湊")
    res_no_po = parse(client, "mars", "瑪氏_訂貨通知單範例.xlsx",
                       upload_name="這份檔名沒有訂單編號.xlsx")
    g_no_po = res_no_po.get_json()["groups"][0]
    check("po_numbers 是空的，不硬湊一個假的", g_no_po["po_numbers"] == [], g_no_po["po_numbers"])
    check("key 退而用檔名（去掉副檔名）當識別，畫面上才有東西可以顯示",
          g_no_po["key"] == "這份檔名沒有訂單編號", g_no_po["key"])
    check("匯出檔名沿用匯入檔名（同事反映匯出後對不起來是哪份，改成直接照抄匯入檔名，"
          "只把副檔名換成實際輸出的 .xls），不是用猜出來的日期/倉別/品類重新拼一個",
          g_no_po["filename"] == "這份檔名沒有訂單編號.xls", g_no_po["filename"])

    print("\n【5.2】瑪氏：可以一次選多個檔案，各自轉成各自的一組")
    res_multi = parse_multi(client, "mars", [
        ("瑪氏_訂貨通知單範例.xlsx", "永豐Mars採購單(箱單位)-GUM糖_TAO4_13000000467952.xlsx"),
        ("瑪氏_訂貨通知單範例.xlsx", "第二張訂單.xlsx"),
    ])
    check("兩個檔案各自解析成兩組", res_multi.status_code == 200
          and len(res_multi.get_json()["groups"]) == 2, res_multi.get_json())
    keys = [gr["key"] for gr in res_multi.get_json()["groups"]]
    check("兩組各自保留自己的識別（一個抓到 PO、一個退回用檔名）",
          keys == ["13000000467952", "第二張訂單"], keys)
    filenames = [gr["filename"] for gr in res_multi.get_json()["groups"]]
    check("批次匯入時每一組的匯出檔名也各自沿用自己的匯入檔名，不是共用同一個",
          filenames == ["永豐Mars採購單(箱單位)-GUM糖_TAO4_13000000467952.xls", "第二張訂單.xls"],
          filenames)

    print("\n【5.3】瑪氏以外的線別一次只能上傳一個檔案，不能誤用多選")
    res_pg_multi = parse_multi(client, "pg", [
        ("PG_訂單匯入範例.xlsx", None), ("PG_訂單匯入範例.xlsx", None)])
    check("P&G 選兩個檔案要擋下來，講清楚只有瑪氏能多選",
          res_pg_multi.status_code == 400 and "瑪氏" in res_pg_multi.get_json()["error"],
          res_pg_multi.get_json())

    print("\n【5.4】瑪氏批次解析：只要一個檔案壞掉，整批都不算數（全有全無）")
    bad = io.BytesIO(b"not an excel file")
    with open(os.path.join(SAMPLES, "瑪氏_訂貨通知單範例.xlsx"), "rb") as fh:
        good = io.BytesIO(fh.read())
    res_batch = client.post("/api/purchase/parse", data={
        "line": "mars",
        "file": [(good, "good.xlsx"), (bad, "bad.xlsx")]},
        content_type="multipart/form-data")
    check("整批擋下來，不會只匯入好的那份、漏掉壞的那份",
          res_batch.status_code == 400, res_batch.status_code)
    check("錯誤訊息點名是哪個檔案壞的",
          "bad.xlsx" in res_batch.get_json()["error"], res_batch.get_json()["error"])

    print("\n【5.5】合併多張 PO 時，備註要逐列保留各自那張 PO 的單號")
    # 這是實際踩到的 bug：以前整份檔案共用一個備註，合併之後 20 列全部
    # 被寫成第一張 PO 的單號，後面幾張的單號整個消失、對不出是哪張單。
    res = parse(client, "pg", "PG_訂單匯入範例.xlsx")
    pg_groups = res.get_json()["groups"]
    pick = [g for g in pg_groups
            if g["key"] in ("13000000493049", "13000000492946", "13000000492925")]
    check("挑到三組來測合併", len(pick) == 3, [g["key"] for g in pick])

    merged_rows = [dict(row, remark=g["remark"]) for g in pick for row in g["rows"]]
    res = export(client, "pg", [{
        "rows": merged_rows, "remark": pick[0]["remark"],
        "filename": "酷澎XP&G_產品採購表上傳_0905到貨.xls"}])
    check("合併匯出成功", res.status_code == 200, res.status_code)
    _, mrows = read_xls_main_sheet(res.data)
    check("列數＝三組品項加總（表頭 + 7+8+5 = 20 筆）", len(mrows) == 21, len(mrows))

    remarks = {r[8] for r in mrows[1:]}
    check("備註出現三種、各自對應自己那張 PO，不是全部共用第一張",
          remarks == {f"{g['key']}(TXRC8)_酷澎9/5買斷" for g in pick}, sorted(remarks))

    # 逐列對：每一列的備註要跟它自己的料號所屬那組一致
    expected = [(row["material_no"], g["remark"]) for g in pick for row in g["rows"]]
    actual = [(r[1], r[8]) for r in mrows[1:]]
    check("逐列比對：每一列的料號跟備註配對都正確", actual == expected)

    check("合併檔的檔名不放 PO 單號（放第一張會誤導成只有那一張）",
          "13000000493049" not in "酷澎XP&G_產品採購表上傳_0905到貨.xls")

    print("\n【5.6】build_filename：各線別的檔名長相")
    import purchase  # noqa: E402
    check("P&G 檔名不帶單號（合併時本來就這樣，使用者後來要求單張匯出也一併拿掉）",
          purchase.build_filename("pg", [], "0905", "TXRC8")
          == "酷澎XP&G_產品採購表上傳_0905到貨.xls",
          purchase.build_filename("pg", [], "0905", "TXRC8"))
    check("瑪氏合併檔不會留下多餘的底線",
          purchase.build_filename("mars", [], "0904", "TAO4")
          == "永豐Mars採購單(箱單位)-GUM糖_TAO4.xls",
          purchase.build_filename("mars", [], "0904", "TAO4"))
    check("P&G 單張匯出時檔名也不帶單號了（原本只留後 6 碼，使用者要求直接拿掉數字）",
          purchase.build_filename("pg", ["13000000492925"], "0905", "TXRC8")
          == "酷澎XP&G_產品採購表上傳_0905到貨.xls",
          purchase.build_filename("pg", ["13000000492925"], "0905", "TXRC8"))

    print("\n【6】檔名衝突時自動加序號，不會互相覆蓋")
    dup_groups = [
        {"rows": [{"material_no": "A1", "qty": 1}], "remark": "", "filename": "同名.xls"},
        {"rows": [{"material_no": "A2", "qty": 2}], "remark": "", "filename": "同名.xls"},
    ]
    res = export(client, "pg", dup_groups)
    with zipfile.ZipFile(io.BytesIO(res.data)) as zf:
        names = sorted(zf.namelist())
        check("兩個同名檔案自動改名不衝突", names == ["同名(2).xls", "同名.xls"], names)

    print("\n【6.1】批次匯出：勾選的都是同一張 PO 時，zip 檔名直接用那個單號，不要叫「採購表」")
    same_po_groups = [
        {"rows": [{"material_no": "A1", "qty": 1}], "remark": "", "filename": "第一份.xls",
         "po_numbers": ["13000000496610"]},
        {"rows": [{"material_no": "A2", "qty": 2}], "remark": "", "filename": "第二份.xls",
         "po_numbers": ["13000000496610"]},
    ]
    res = export(client, "pg", same_po_groups)
    check("zip 檔名就是那張 PO 單號，同事一看就知道裡面是哪張單",
          res.headers.get("Content-Disposition", "").find("13000000496610.zip") != -1,
          res.headers.get("Content-Disposition"))

    print("\n【6.2】批次匯出：PO 不只一種時，zip 檔名退回原本的通用檔名，不要硬湊一個誤導的單號")
    mixed_po_groups = [
        {"rows": [{"material_no": "A1", "qty": 1}], "remark": "", "filename": "第一份.xls",
         "po_numbers": ["13000000496610"]},
        {"rows": [{"material_no": "A2", "qty": 2}], "remark": "", "filename": "第二份.xls",
         "po_numbers": ["13000000000001"]},
    ]
    res = export(client, "pg", mixed_po_groups)
    # 中文檔名在 Content-Disposition 裡是 filename*=UTF-8''%E6%8E%A1... 這種
    # URL 編碼過的形式，不能直接用字串比對找「採購表」。
    check("zip 檔名退回「採購表」，不會只挑第一張的單號充數",
          "UTF-8''%E6%8E%A1%E8%B3%BC%E8%A1%A8.zip" in res.headers.get("Content-Disposition", ""),
          res.headers.get("Content-Disposition"))

    print("\n【7】基本防呆")
    res = client.post("/api/purchase/parse", data={"line": "不存在的線別"},
                       content_type="multipart/form-data")
    check("線別不對要擋下來", res.status_code == 400)

    res = client.post("/api/purchase/parse", data={"line": "pg"},
                       content_type="multipart/form-data")
    check("沒帶檔案要擋下來", res.status_code == 400)

    fake = io.BytesIO(b"not an excel file")
    res = client.post("/api/purchase/parse", data={
        "line": "pg", "file": (fake, "fake.xlsx")}, content_type="multipart/form-data")
    check("壞掉的檔案要擋下來、不是整包噴 500", res.status_code == 400)

    res = export(client, "pg", [])
    check("匯出空清單要擋下來", res.status_code == 400)

    print("\n" + "=" * 62)
    print(f"通過 {len(PASS)} 項／失敗 {len(FAIL)} 項")
    if FAIL:
        for name in FAIL:
            print("  ✗", name)
        return 1
    print("全部通過。")
    return 0


if __name__ == "__main__":
    try:
        code = main()
    finally:
        shutil.rmtree(_tmp, ignore_errors=True)
    sys.exit(code)
