"""商品主檔自動化 前端煙霧測試（Playwright 開真的瀏覽器）。

後端測試（test_master.py）看不到 JavaScript 壞掉——今天就漏過「選回第一條沒反應」和
「誤刪 loadLogs」兩個 bug。這支把人每次手動點的那幾條路寫成腳本：登入、三個分頁、
篩選列、各次匯入的下拉、PO 視窗、就地編輯、上傳預覽、三個匯出連結，任何 JS 錯誤或
4xx/5xx 都算失敗。

執行：python test_master_ui.py
需要：pip install playwright && playwright install chromium（CI 會自動裝）。
資料：samples/master/fake/ 的假資料（商品、條碼、PO 全是編的）。
"""
import datetime
import glob
import io
import openpyxl
import pymupdf
import os
import socket
import sys
import tempfile
import threading
import time
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
FAKE = os.path.join(BASE, "samples", "master", "fake")
os.environ["COUPANG_OMS_DATA"] = tempfile.mkdtemp(prefix="oms_ui_test_")
sys.path.insert(0, BASE)

import db  # noqa: E402
import app as app_module  # noqa: E402
import logging  # noqa: E402
logging.getLogger("werkzeug").setLevel(logging.ERROR)   # 不要讓每個請求的 log 淹掉測試結果

PASS, FAIL = [], []


def check(name, ok, detail=""):
    (PASS if ok else FAIL).append(name)
    print(("  ✓ " if ok else "  ✗ ") + name + (f"  — {detail}" if detail and not ok else ""))


def free_port():
    s = socket.socket(); s.bind(("127.0.0.1", 0)); p = s.getsockname()[1]; s.close(); return p


def seed(client):
    """用 API 把假資料灌進去：三個主檔、寶僑總表、兩次訂單彙總表。"""
    def up(url, name, **form):
        d = dict(form); d["file"] = (io.BytesIO(open(os.path.join(FAKE, name), "rb").read()), name)
        return client.post(url, data=d, content_type="multipart/form-data")
    for n in ["假_酷澎主檔_寶僑.xlsx", "假_酷澎主檔_瑪氏.xlsx", "假_酷澎主檔_CPG.xlsx", "假_寶僑總表.xlsx"]:
        assert up("/api/master/products/import", n).status_code == 200, n
    for n in ["假_訂單彙總表_9月_第一次.xlsx", "假_訂單彙總表_9月_第二次.xlsx"]:
        pv = up("/api/master/import/preview", n).get_json()
        assert client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"], "add_missing_products": True}).status_code == 200
    # 第三批：模擬 nicole 那次——6 張新 PO，1 張 9/23、5 張 10 月，用來驗跨月一次列完
    import datetime, openpyxl
    wb = openpyxl.load_workbook(os.path.join(FAKE, "假_訂單彙總表_9月_第二次.xlsx")); ws = wb.worksheets[0]
    rows = [list(r) for r in ws.iter_rows(min_row=2, values_only=True) if r[1] == "寶僑"][:48]
    ws.delete_rows(2, ws.max_row)
    plan = [("13000000600901", datetime.datetime(2026, 9, 23), 3), ("13000000600902", datetime.datetime(2026, 10, 6), 1), ("13000000600903", datetime.datetime(2026, 10, 7), 12),
            ("13000000600904", datetime.datetime(2026, 10, 13), 10), ("13000000600905", datetime.datetime(2026, 10, 13), 12), ("13000000600906", datetime.datetime(2026, 10, 13), 10)]
    i = 0
    for po, d, n in plan:
        for k in range(n):
            r = list(rows[i % len(rows)]); i += 1
            r[2] = po; r[5] = d; r[6] = k + 1; r[3] = "TXRC8" if d.month == 9 else "TAO5"
            ws.append(r)
    buf = io.BytesIO(); wb.save(buf)
    pv = client.post("/api/master/import/preview", data={"file": (io.BytesIO(buf.getvalue()), "假_訂單彙總表_第三次_跨月.xlsx")}, content_type="multipart/form-data").get_json()
    assert client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"], "add_missing_products": True}).status_code == 200


def main():
    from playwright.sync_api import sync_playwright
    db.init_db(); app_module.app.config["TESTING"] = True
    tc = app_module.app.test_client(); tc.post("/login", data={"username": "Jerry", "password": "changeme123"})
    seed(tc)

    port = free_port()
    threading.Thread(target=lambda: app_module.app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False), daemon=True).start()
    time.sleep(1.5)
    base = f"http://127.0.0.1:{port}"
    errs, bad = [], []
    launch = {}
    local = glob.glob("/opt/pw-browsers/chromium-*/chrome-linux*/chrome")
    if local and not os.environ.get("CI"):
        launch["executable_path"] = local[0]

    with sync_playwright() as p:
        b = p.chromium.launch(**launch); pg = b.new_page(viewport={"width": 1500, "height": 900})
        pg.on("pageerror", lambda e: errs.append(str(e)))
        pg.on("response", lambda r: bad.append((r.status, r.url)) if r.status >= 400 else None)

        print("\n【1】登入與進頁")
        pg.goto(f"{base}/login"); pg.fill("input[name=username]", "Jerry"); pg.fill("input[name=password]", "changeme123"); pg.click("button[type=submit]")
        pg.goto(f"{base}/master"); pg.wait_for_selector("#dd-line-btn", state="attached"); pg.wait_for_timeout(600)
        check("頁面載入、JS 檔有帶版本號", "master.js?v=" in pg.eval_on_selector("script[src*=master]", "e => e.src"))
        pg.evaluate("setMonth('2026-09')"); pg.wait_for_selector("#last-import:not(.hidden)"); pg.wait_for_timeout(600)
        check("訂單明細有資料", "個品項" in pg.inner_text("#orders-count"))

        print("\n【1b】線別工具選單（三頁共用）")
        pg.goto(f"{base}/coupang"); pg.wait_for_selector("#lm-btn"); pg.wait_for_timeout(400)
        pg.click("#f-status"); pg.wait_for_selector("#f-status-pop:not(.hidden) input.st-ck", timeout=5000)
        pg.check("#f-status-pop input.st-ck[value='已建立']"); pg.check("#f-status-pop input.st-ck[value='修改中']")
        check("PO 狀態可以複選：勾兩個 → 按鈕寫「已建立、修改中」", pg.inner_text("#f-status-text") == "已建立、修改中", pg.inner_text("#f-status-text"))
        with pg.expect_response(lambda r: "/api/pos?" in r.url and "po_status=" in r.url, timeout=5000) as rq:
            pg.wait_for_timeout(10)
        check("不用按搜尋：勾完自己帶兩個狀態去查（Jerry 2026-10-22「要即刻跳結果」）", "po_status=%E5%B7%B2%E5%BB%BA%E7%AB%8B%2C%E4%BF%AE%E6%94%B9%E4%B8%AD" in rq.value.url, rq.value.url)
        pg.click("body", position={"x": 5, "y": 5}); check("點旁邊 → 選單收起來", pg.is_hidden("#f-status-pop"))
        pg.click("#f-pulled"); pg.wait_for_selector("#f-pulled-pop:not(.hidden) input.st-ck", timeout=5000)
        check("其他下拉也是勾選清單（Jerry 2026-10-22「其他下拉選單也想要複選」）：點開拉單狀態，PO 狀態那個自己收起來",
              pg.is_hidden("#f-status-pop") and pg.eval_on_selector_all("#f-pulled-pop input.st-ck", "els => els.length") == 2)
        with pg.expect_response(lambda r: "/api/pos?" in r.url and "is_pulled=0" in r.url, timeout=5000):
            pg.check("#f-pulled-pop input.st-ck[value='0']")
        check("勾未拉單 → 不用按搜尋就查、按鈕寫「未拉單」", pg.inner_text("#f-pulled-text") == "未拉單")
        pg.click("#f-line"); pg.wait_for_selector("#f-line-pop:not(.hidden)", timeout=5000)
        n_line = pg.eval_on_selector_all("#f-line-pop input.st-ck", "els => els.length")
        if n_line >= 2:
            with pg.expect_response(lambda r: "/api/pos?" in r.url and "line=" in r.url and "%2C" in r.url, timeout=5000):
                pg.check("#f-line-pop input.st-ck >> nth=0"); pg.check("#f-line-pop input.st-ck >> nth=1")
            check("線別勾兩個 → 帶兩個線別去查、按鈕寫兩個", "、" in pg.inner_text("#f-line-text"), pg.inner_text("#f-line-text"))
        pg.click("body", position={"x": 5, "y": 5})
        with pg.expect_response(lambda r: "/api/pos?" in r.url and "keyword=" in r.url, timeout=5000):
            pg.fill("#f-keyword", "偉嘉")
        check("打關鍵字 → 停一下自己查", True)
        pg.click("#btn-reset"); pg.wait_for_timeout(500)
        check("按重設 → PO 狀態、拉單狀態、線別都回到「全部」", pg.inner_text("#f-status-text") == "全部" and pg.inner_text("#f-pulled-text") == "全部"
              and pg.inner_text("#f-line-text") == "全部")
        check("首頁也是一顆「線別工具」、不放 logo、有三個線別色點", pg.eval_on_selector_all("#lm-btn img", "els => els.length") == 0 and pg.eval_on_selector_all("#lm-btn .lm-dots i", "els => els.length") == 3)
        pg.click("#lm-btn"); pg.wait_for_timeout(200)
        check("首頁點開 → 三欄、寶僑那欄標題是寶僑藍", pg.eval_on_selector_all("#lm-pop .lm-h", "els => els.length") == 3 and pg.eval_on_selector("#lm-pop .lm-h", "e => getComputedStyle(e).color") == "rgb(23, 95, 171)")
        pg.click('#lm-pop a.lm-item'); pg.wait_for_selector("#dd-line-btn", state="attached"); pg.wait_for_timeout(600)
        check("從首頁選單點商品主檔自動化 → 進到那頁", "/master" in pg.url)
        pg.click("#lm-btn"); pg.wait_for_timeout(200)
        check("點「線別工具」展開、三欄：寶僑／瑪氏／紙潔", pg.evaluate("document.getElementById('line-menu').classList.contains('open')") and pg.eval_on_selector_all("#lm-pop .lm-h", "els => els.map(e => e.innerText.trim())") == ["寶僑", "瑪氏", "紙潔"])
        check("寶僑底下有商品主檔自動化、而且標「目前在這」", pg.eval_on_selector("#lm-pop a.lm-item.on", "e => e.innerText").startswith("商品主檔自動化"))
        check("每欄都有 logo 圖", pg.eval_on_selector_all("#lm-pop .lm-h img", "els => els.every(i => i.complete && i.naturalWidth > 0)"))
        pg.mouse.click(700, 500); pg.wait_for_timeout(200)
        check("點外面收起", not pg.evaluate("document.getElementById('line-menu').classList.contains('open')"))
        pg.click("#lm-btn"); pg.keyboard.press("Escape"); pg.wait_for_timeout(200)
        check("按 Esc 收起", not pg.evaluate("document.getElementById('line-menu').classList.contains('open')"))
        pg.evaluate("setMonth('2026-09')"); pg.wait_for_selector("#last-import:not(.hidden)"); pg.wait_for_timeout(600)

        print("\n【2】篩選列")
        pg.click("#dd-line-btn"); pg.click('#dd-line input[data-v="寶僑"]'); pg.wait_for_timeout(600)
        check("勾線別後按鈕寫出選了什麼", pg.inner_text("#dd-line-n") == "寶僑", pg.inner_text("#dd-line-n"))
        check("清除篩選按鈕出現", not pg.evaluate("document.getElementById('btn-clear').classList.contains('hidden')"))
        pg.click("#dd-po-btn"); pg.wait_for_timeout(200); pg.fill("#dd-po .dd-q", "600001"); pg.wait_for_timeout(200)
        check("PO 下拉可以用關鍵字過濾", pg.eval_on_selector_all("#dd-po label", "els => els.filter(e => e.style.display !== 'none').length") == 1)
        pg.click("#btn-clear"); pg.wait_for_timeout(600)
        check("清除後回到全部", pg.inner_text("#dd-line-n") == "全部")

        print("\n【3】最近匯入那一行、依匯入批次篩選、抽屜")
        bar = pg.inner_text("#last-import")
        check("上面四張卡：今天要出貨／本週要出貨／最近匯入／需要處理", pg.eval_on_selector_all("#last-import .kpi", "els => els.length") == 4 and "今天要出貨" in bar and "本週要出貨" in bar and "最近匯入" in bar and "需要處理" in bar and "張 PO" in bar, bar)
        check("表頭不再是深藍底", pg.eval_on_selector("#orders-list table.m th", "e => getComputedStyle(e).backgroundColor !== 'rgb(30, 64, 175)'"))
        check("全部訂單模式下，最近一批動到的列有色條標籤", pg.eval_on_selector_all("#orders-list tr.chg-new, #orders-list tr.chg-upd, #orders-list tr.chg-gone", "els => els.length") > 0)
        check("行事曆角標標的是 PO 張數", pg.eval_on_selector_all("#cal .day .dc", "els => els.length") > 0 and "張 PO" in pg.get_attribute("#cal .day .dc", "title"))
        pg.click("#btn-batch-filter"); pg.wait_for_timeout(400)
        check("按「依匯入批次篩選」拉出抽屜、一次一列", pg.evaluate("document.getElementById('imp-drawer').classList.contains('open')") and pg.eval_on_selector_all("#imp-drawer .imp-row", "els => els.length") >= 2)
        check("抽屜是推開式：主畫面往左縮、匯出按鈕沒被遮", pg.evaluate("document.body.classList.contains('drawer-open')") and pg.evaluate("(() => { const r = document.getElementById('btn-export-daily').getBoundingClientRect(); const d = document.getElementById('imp-drawer').getBoundingClientRect(); return r.right <= d.left + 1; })()"))
        pg.mouse.click(300, 400); pg.wait_for_timeout(400)
        check("點抽屜以外的地方就關", not pg.evaluate("document.getElementById('imp-drawer').classList.contains('open')"))
        pg.click("#btn-batch-filter"); pg.wait_for_timeout(400)
        check("抽屜字放大到 14px", pg.eval_on_selector("#imp-drawer .imp-row", "e => parseFloat(getComputedStyle(e).fontSize)") >= 14)
        pg.fill("#imp-q", "第一次"); pg.wait_for_timeout(300)
        check("抽屜可以搜檔名", pg.eval_on_selector_all("#imp-drawer .imp-row", "els => els.length") == 1 and "第一次" in pg.inner_text("#imp-drawer .imp-row"))
        po_any = pg.eval_on_selector("#orders-list .po", "e => e.dataset.po")
        pg.fill("#imp-q", po_any); pg.wait_for_timeout(700)
        check("抽屜可以貼 PO 單號找它在哪幾批", pg.eval_on_selector_all("#imp-drawer .imp-row", "els => els.length") >= 1, pg.inner_text("#imp-count"))
        pg.fill("#imp-q", ""); pg.wait_for_timeout(300)
        pg.select_option("#imp-period", "month"); pg.wait_for_timeout(300)
        check("抽屜可以篩期間（本月）", pg.eval_on_selector_all("#imp-drawer .imp-row", "els => els.length") >= 2)
        pg.select_option("#imp-period", ""); pg.wait_for_timeout(300)
        pg.click("#imp-drawer .imp-row >> nth=0"); pg.wait_for_timeout(1000)
        bar = pg.inner_text("#last-import")
        check("選最近一批 → 那一行寫「依匯入批次篩選：… 這批」，有「清除批次篩選」", "依匯入批次篩選：" in bar and "這批" in bar and "清除批次篩選" in bar, bar)
        check("表格只剩有標籤的列", pg.eval_on_selector_all("#orders-list table.m tbody tr", "els => els.length") > 0 and pg.eval_on_selector_all("#orders-list table.m tbody tr:not(.chg-new):not(.chg-upd):not(.chg-gone)", "els => els.length") == 0)
        check("批次模式行事曆跟全部訂單同一套：沒動到的天壓淡、動到的天亮著、有角標", not pg.evaluate("document.getElementById('cal').classList.contains('hidden')") and pg.eval_on_selector_all("#cal .day.dim", "els => els.length") >= 1 and pg.eval_on_selector_all("#cal .day.has:not(.dim) .dc", "els => els.length") >= 1)
        check("跨月時行事曆一個月一塊往下排（9 月、10 月）", pg.eval_on_selector_all("#cal .cal-month", "els => els.length") == 2, str(pg.eval_on_selector_all("#cal .cm-title", "els => els.map(e => e.innerText)")))
        hit_date = pg.eval_on_selector("#cal .day.has:not(.dim)", "e => e.dataset.date")
        pg.click("#cal .day.has:not(.dim) >> nth=0"); pg.wait_for_timeout(1200)
        check("批次模式點一天 → 跟全部訂單一樣是選日期（表格只剩那天），並且那塊亮起來、PO 展開", pg.evaluate("[...state.dates]") == [hit_date] and pg.eval_on_selector_all("#orders-list .card[data-date]", "els => els.length") == 1 and pg.eval_on_selector_all(f'#orders-list .card.hl[data-date="{hit_date}"]', "els => els.length") == 1 and pg.eval_on_selector_all(f'#orders-list .card[data-date="{hit_date}"] .grp-po:not(.collapsed)', "els => els.length") >= 1)
        dims = pg.eval_on_selector_all("#cal .day.has.dim", "els => els.map(e => e.dataset.date)")
        if dims:
            pg.click(f'#cal .day.has.dim[data-date="{dims[0]}"]'); pg.wait_for_timeout(500)
            check("壓淡的天（這批沒動到）點了不會被選", dims[0] not in pg.evaluate("[...state.dates]"))
        pg.click("#cal-clear-dates"); pg.wait_for_timeout(700)
        check("月份藥丸寫出跨月範圍", "～" in pg.inner_text("#month-pill-txt"), pg.inner_text("#month-pill-txt"))
        pg.click('#cal-view [data-v="list"]'); pg.wait_for_timeout(400)
        check("切成清單：一列一天、有「這批」欄", not pg.evaluate("document.getElementById('cal-list').classList.contains('hidden')") and pg.eval_on_selector_all("#cal-list tbody tr[data-date]", "els => els.length") >= 1 and "這批" in pg.inner_text("#cal-list thead"))
        pg.click('#cal-view [data-v="cal"]'); pg.wait_for_timeout(300)
        check("匯出按鈕寫「匯出這批的專案報價檔（N 張 PO）」", "這批" in pg.inner_text("#btn-export-daily") and "張 PO" in pg.inner_text("#btn-export-daily"), pg.inner_text("#btn-export-daily"))
        check("月份自動對到那批的交期（9 月）", pg.input_value("#sel-month") == "2026-09", pg.input_value("#sel-month"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        check("Esc 關抽屜", not pg.evaluate("document.getElementById('imp-drawer').classList.contains('open')"))
        pg.click("#btn-batch-clear"); pg.wait_for_timeout(700)
        check("清除批次篩選 → 回最近匯入、行事曆回來、月份藥丸回單月", "最近匯入" in pg.inner_text("#last-import") and "這批" not in pg.inner_text("#btn-export-daily") and not pg.evaluate("document.getElementById('cal').classList.contains('hidden')") and "～" not in pg.inner_text("#month-pill-txt"))
        check("表格還沒捲進來時導覽條先藏著，不蓋行事曆", pg.evaluate("document.getElementById('date-rail').classList.contains('hidden')"))
        pg.evaluate("document.querySelector('#orders-list').scrollIntoView()"); pg.wait_for_timeout(500)
        check("表格捲進畫面 → 右側日期導覽條出來、一天一格", not pg.evaluate("document.getElementById('date-rail').classList.contains('hidden')") and pg.eval_on_selector_all("#date-rail a", "els => els.length") >= 3)
        pg.click("#date-rail a >> nth=2"); pg.wait_for_timeout(600)
        check("點導覽條跳到那天（那天的標題進到畫面上方）", pg.evaluate("(() => { const d = document.querySelector('#date-rail a:nth-child(3)').dataset.date; const c = [...document.querySelectorAll('#orders-list .card[data-date]')].find(x => x.dataset.date === d); const r = c.getBoundingClientRect(); return r.top >= -5 && r.top < 200; })()"))
        check("導覽條跳過去的那天 PO 會展開", pg.evaluate("(() => { const d = document.querySelector('#date-rail a:nth-child(3)').dataset.date; return document.querySelectorAll(`#orders-list .card[data-date=\"${d}\"] .grp-po:not(.collapsed)`).length > 0; })()"))
        pg.click("#btn-collapse-all"); pg.wait_for_timeout(300)   # 收回來，後面「預設收合」的檢查才乾淨
        pg.evaluate("window.scrollTo(0,0)"); pg.wait_for_timeout(900)   # 平滑捲動要等它停，不然下面拖選的座標會抓到捲動中的位置
        pg.click('#density [data-d="dense"]'); pg.wait_for_timeout(200)
        check("密度切緊湊", pg.evaluate("document.getElementById('orders-list').classList.contains('dense')"))
        pg.click('#density [data-d="normal"]'); pg.wait_for_timeout(200)
        check("日期標題是黏頂的", pg.eval_on_selector("#orders-list .grp-date", "e => getComputedStyle(e).position === 'sticky'"))

        print("\n【3b】行事曆按著滑過多選、日期清單")
        pg.evaluate("window.scrollTo(0,0)"); pg.wait_for_timeout(700)
        days = pg.eval_on_selector_all("#cal .day.has", "els => els.map(e => e.dataset.date)")
        b1 = pg.locator(f'.day.has[data-date="{days[0]}"]').bounding_box(); b2 = pg.locator(f'.day.has[data-date="{days[2]}"]').bounding_box()
        pg.mouse.move(b1["x"] + 10, b1["y"] + 10); pg.mouse.down()
        pg.mouse.move(b2["x"] + 10, b2["y"] + 10, steps=8); pg.mouse.up(); pg.wait_for_timeout(900)
        check("按著從第 1 天滑到第 3 天 → 至少選起 2 天", pg.eval_on_selector_all("#cal .day.on", "els => els.length") >= 2, str(pg.eval_on_selector_all("#cal .day.on", "els => els.map(e => e.dataset.date)")))
        check("滑選的那幾天在表格裡一起黃一下（幾天就幾塊）", pg.eval_on_selector_all("#orders-list .card.hl", "els => els.length") == pg.eval_on_selector_all("#cal .day.on", "els => els.length"))
        pg.click("#cal-clear-dates"); pg.wait_for_timeout(700)
        check("全部取消", pg.eval_on_selector_all("#cal .day.on", "els => els.length") == 0)
        check("本週／下週／最近一批那排快捷已拿掉", pg.eval_on_selector_all("#cal-quick", "els => els.length") == 0)
        pg.click('#cal-view [data-v="list"]'); pg.wait_for_timeout(400)
        check("切到日期清單：一列一天", not pg.evaluate("document.getElementById('cal-list').classList.contains('hidden')") and pg.eval_on_selector_all("#cal-list tbody tr", "els => els.length") >= 3)
        pg.click("#cal-list tbody tr >> nth=0"); pg.wait_for_timeout(700)
        pg.click("#cal-list tbody tr >> nth=2", modifiers=["Shift"]); pg.wait_for_timeout(900)
        check("Shift 點兩列 → 中間整段一起勾（3 天）", pg.eval_on_selector_all("#cal-list tbody tr.on", "els => els.length") == 3, str(pg.eval_on_selector_all("#cal-list tbody tr.on", "els => els.length")))
        pg.click("#cal-clear-dates"); pg.wait_for_timeout(600)
        r1 = pg.locator("#cal-list tbody tr >> nth=0").bounding_box(); r4 = pg.locator("#cal-list tbody tr >> nth=3").bounding_box()
        pg.mouse.move(r1["x"] + 200, r1["y"] + 10); pg.mouse.down(); pg.mouse.move(r4["x"] + 200, r4["y"] + 10, steps=8); pg.mouse.up(); pg.wait_for_timeout(900)
        check("日期清單按著往下拖 → 4 天一起勾", pg.eval_on_selector_all("#cal-list tbody tr.on", "els => els.length") == 4, str(pg.eval_on_selector_all("#cal-list tbody tr.on", "els => els.length")))
        pg.click("#cal-clear-dates"); pg.wait_for_timeout(600)
        pg.click('#cal-view [data-v="cal"]'); pg.wait_for_timeout(400)
        check("「算不出箱數」「人工調整過」有事才出現（假資料有算不出箱數的品項 → 顯示筆數）", "個品項無法計算箱數" in pg.inner_text("#chip-missing") or pg.evaluate("document.getElementById('chip-missing').classList.contains('hidden')"))

        print("\n【4】行事曆、倉別、PO 視窗、就地編輯")
        pg.click('.day.has[data-date="2026-09-04"]'); pg.wait_for_timeout(900)
        check("點日期後倉別區塊顯示已勾 1 天", "已勾 1 天" in pg.inner_text("#wh-block"))
        check("只點一天：那天黃一下、底下 PO 順便展開", pg.eval_on_selector_all('#orders-list .card.hl[data-date="2026-09-04"]', "els => els.length") == 1 and pg.eval_on_selector_all("#orders-list .grp-po:not(.collapsed)", "els => els.length") >= 1)
        pg.click("#btn-collapse-all"); pg.wait_for_timeout(300)
        check("PO 明細預設收起來、日期展開（不全開也不全關）", pg.eval_on_selector_all("#orders-list .grp-po.collapsed", "els => els.length") >= 1 and pg.eval_on_selector_all("#orders-list .card.collapsed", "els => els.length") == 0)
        pg.click("#orders-list .grp-po >> nth=0"); pg.wait_for_timeout(400)
        check("點 PO 標題列 → 那張的品項展開", pg.eval_on_selector_all("#orders-list .grp-po:not(.collapsed)", "els => els.length") == 1)
        pg.click("#orders-list .grp-date >> nth=0"); pg.wait_for_timeout(400)
        check("點日期標題 → 整天收合", pg.eval_on_selector_all("#orders-list .card.collapsed", "els => els.length") == 1)
        pg.click("#btn-expand-all"); pg.wait_for_timeout(500)
        check("全部展開", pg.eval_on_selector_all("#orders-list .grp-po.collapsed, #orders-list .card.collapsed", "els => els.length") == 0)
        check("PO 單號是深藍底白字的大標籤", pg.eval_on_selector("#orders-list .grp-po .po", "e => parseFloat(getComputedStyle(e).fontSize) >= 15 && getComputedStyle(e).color === 'rgb(255, 255, 255)'"))
        pg.click(".po >> nth=0"); pg.wait_for_timeout(700)
        check("PO 視窗打開", pg.evaluate("document.getElementById('dlg-po').open")); pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        cell = pg.query_selector(".inline-cell") or pg.query_selector("td[data-field]")
        if cell:
            cell.click(); pg.wait_for_timeout(300); pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
            check("就地編輯按 Esc 會回復", pg.query_selector(".inline-cell input") is None)

        print("\n【4b】改單原因：改出貨數量／交貨日前一定要選")
        qcell = pg.query_selector('#orders-list td[data-field="qty_ship"]')
        tr_id = qcell.evaluate("e => e.closest('tr').dataset.id"); old_q = qcell.evaluate("e => e.querySelector('b').innerText")
        qcell.click(); pg.wait_for_timeout(200); pg.fill('#orders-list td[data-field="qty_ship"] input', "7"); pg.keyboard.press("Enter"); pg.wait_for_timeout(400)
        check("按 Enter 後先跳出「為什麼要改」小視窗", pg.evaluate("document.getElementById('dlg-reason').open"))
        check("四個原因按鈕：缺貨／沒車／酷澎要求／其他", pg.eval_on_selector_all("#rsn-opts button", "els => els.map(e => e.innerText)") == ["缺貨", "沒車", "酷澎要求", "其他"])
        check("沒選之前確定按不下去", pg.is_disabled("#rsn-ok"))
        pg.click('#rsn-opts button[data-r="其他"]'); pg.wait_for_timeout(150)
        check("選「其他」才出現說明欄", not pg.evaluate("document.getElementById('rsn-note').classList.contains('hidden')"))
        pg.click("#rsn-ok"); pg.wait_for_timeout(200)
        check("「其他」沒寫說明按確定會被擋、視窗還開著", pg.evaluate("document.getElementById('dlg-reason').open") and "其他" in pg.inner_text("#rsn-msg"))
        pg.check("#rsn-test"); pg.wait_for_timeout(150)
        check("勾「這是測試」後不選原因也能按確定", not pg.is_disabled("#rsn-ok"))
        pg.uncheck("#rsn-test"); pg.wait_for_timeout(150)
        pg.click('#rsn-opts button[data-r="缺貨"]'); pg.wait_for_timeout(150)
        check("換選缺貨後說明欄收起", pg.evaluate("document.getElementById('rsn-note').classList.contains('hidden')"))
        pg.click("#rsn-ok"); pg.wait_for_timeout(1200)
        saved = pg.evaluate(f"document.querySelector('tr[data-id=\"{tr_id}\"] td[data-field=qty_ship] b').innerText")
        check("選好原因後真的存進去（出貨數量變 7）", saved == "7", saved)
        pg.evaluate(f"openPo(state.rows.find(r => String(r.id) === '{tr_id}').po_number)"); pg.wait_for_timeout(900)
        check("PO 視窗歷程顯示原因標籤「缺貨」", "缺貨" in pg.eval_on_selector_all("#po-logs .bd-reason", "els => els.map(e => e.innerText)"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        qcell = pg.query_selector(f'tr[data-id="{tr_id}"] td[data-field="qty_ship"]'); qcell.click(); pg.wait_for_timeout(200)
        pg.fill(f'tr[data-id="{tr_id}"] td[data-field="qty_ship"] input', "9"); pg.keyboard.press("Enter"); pg.wait_for_timeout(400)
        pg.keyboard.press("Escape"); pg.wait_for_timeout(500)
        check("原因視窗按取消 → 數字不改、格子回復", pg.evaluate(f"document.querySelector('tr[data-id=\"{tr_id}\"] td[data-field=qty_ship] b').innerText") == "7" and pg.query_selector(".inline-cell input") is None)
        pg.click("#btn-logs"); pg.wait_for_timeout(600)
        check("修改歷程表多一欄「原因」", "原因" in pg.eval_on_selector_all("#lg-table th", "els => els.map(e => e.innerText)"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(200)
        pg.click("#btn-clear"); pg.wait_for_timeout(500)

        print("\n【4c】④ 改單統計分頁（2026-10-21 起以修改紀錄的 OK 數為準）")
        pg.click('.m-tab[data-tab="stats"]'); pg.wait_for_timeout(900)
        check("四張卡：PO 張數／有修改紀錄的 PO／處理次數／處理完成後又有改動（估計工時 2026-10-29 拿掉）", pg.eval_on_selector_all("#st-kpi .kpi", "els => els.length") == 4
              and all(w in pg.inner_text("#st-kpi") for w in ("有修改紀錄的 PO", "處理次數", "處理完成後又有改動")) and "估計工時" not in pg.inner_text("#tab-stats")
              and pg.query_selector("#st-min") is None, pg.inner_text("#st-kpi"))
        check("拿掉的：含測試勾選框、酷澎修改／我方修改、改單類型表", pg.query_selector("#st-test") is None and "我方修改" not in pg.inner_text("#st-kpi") and pg.query_selector("#st-kinds") is None)
        check("每月統計（Jerry 2026-10-22 第二版）：9 月那列、表頭分組（訂單／修改紀錄／這輪做了什麼／參考）、單位寫在表頭、10/20 以前寫「—」",
              "2026-09" in pg.inner_text("#st-months")
              and all(w in pg.inner_text("#st-months thead") for w in ("PO", "有修改紀錄", "處理次數", "這輪做了什麼", "系統變動"))
              and all(w in pg.inner_text("#st-months tbody") for w in ("張", "次", "收單", "下修", "改交期／倉別", "拆單"))
              and "—" in pg.inner_text("#st-months tbody tr >> nth=0"), pg.inner_text("#st-months thead"))
        check("可點的數字不畫底線", pg.eval_on_selector("#st-months .st-n", "e => getComputedStyle(e).textDecorationLine") == "none")
        check("主因表、各階段做了什麼（三個階段加合計）都有", pg.query_selector("#st-reasons thead") and all(w in pg.inner_text("#st-stages") for w in ("第一階段", "第二階段", "第三階段", "合計")))
        check("剛剛改的那筆算進系統抓到的變動：灰字可以點", pg.eval_on_selector_all("#st-months .st-n.st-g", "els => els.length") >= 1)
        pg.click('#st-months .st-n.st-g >> nth=0'); pg.wait_for_timeout(900)
        check("點灰字 → 明細視窗打開、標題寫「系統抓到的變動」幾次", pg.evaluate("document.getElementById('dlg-stev').open") and "系統抓到的變動" in pg.inner_text("#stev-title") and "次" in pg.inner_text("#stev-title"))
        check("明細列出時間／PO／誰／改了什麼、不列原因", pg.eval_on_selector_all("#stev-table tbody tr", "els => els.length") >= 1 and "我方修改" in pg.inner_text("#stev-table") and "原因" not in pg.inner_text("#stev-table thead"))
        pg.click("#stev-table .stev-po >> nth=0"); pg.wait_for_timeout(700)
        check("明細裡點 PO 會打開 PO 視窗", pg.evaluate("document.getElementById('dlg-po').open")); pg.keyboard.press("Escape"); pg.wait_for_timeout(200); pg.keyboard.press("Escape"); pg.wait_for_timeout(200)

        print("\n【4d】① 來源檔上傳：三個檔一起拖進主檔匯入口")
        pg.click('.m-tab[data-tab="products"]'); pg.wait_for_timeout(700)
        check("主檔表格有 GIV／NIV 兩欄", "GIV" in pg.inner_text("#prod-table thead") and "NIV" in pg.inner_text("#prod-table thead"))
        check("匯入口下面有四格「上次上傳」、還沒上傳過", pg.eval_on_selector_all("#src-status .src", "els => els.length") == 4 and "尚未匯入" in pg.inner_text("#src-status"))
        pg.set_input_files("#file-prod", [os.path.join(FAKE, n) for n in ("假_寶僑supply表.xlsx", "假_CoupangMaster.xlsx", "假_庫存銷售表.xlsx")]); pg.wait_for_timeout(300)
        pg.click("#prod-files-go"); pg.wait_for_selector("#prod-imp-msg .bi-check-circle", timeout=30000); pg.wait_for_timeout(800)
        msg = pg.inner_text("#prod-imp-msg")
        check("三個檔各自被認出來：supply 表／Coupang Master／庫存銷售表", all(k in msg for k in ["寶僑 supply 表", "Coupang Master", "庫存銷售表"]), msg[:200])
        check("訊息寫了對到幾個商品、更新幾筆、哪幾個月", "對到主檔" in msg and "月份" in msg)
        check("上傳後「上次上傳」三格變成今天", pg.inner_text("#src-status").count("今天") >= 3, pg.inner_text("#src-status")[:200])
        check("主檔表格出現 GIV 數字", pg.eval_on_selector_all("#prod-table tbody tr", "els => els.filter(tr => /\\d/.test(tr.children[9].innerText)).length") >= 1)

        print("\n【5】其他分頁與視窗")
        for tab in ["products", "summary", "orders"]:
            pg.click(f'.m-tab[data-tab="{tab}"]'); pg.wait_for_timeout(700)
        pg.click('.m-tab[data-tab="summary"]'); pg.wait_for_timeout(500); pg.select_option("#sum-line", "寶僑"); pg.wait_for_timeout(900)
        check("③ 總表顯示照哪份底稿匯出", "假_寶僑總表" in pg.inner_text("#tpl-info"))
        pg.click('#bd-view button[data-v="products"]'); pg.wait_for_timeout(300)
        check("③ 看板四張卡：下單、下單金額 GIV、剩餘庫存、需要注意", pg.eval_on_selector_all("#bd-kpi .kpi", "els => els.length") == 4 and all(k in pg.inner_text("#bd-kpi") for k in ["月下單", "GIV", "剩餘庫存", "需要處理"]), pg.inner_text("#bd-kpi")[:160])
        check("③ 商品表有分組表頭（單價／供需／庫存／金額）與列", all(k in pg.inner_text("#bd-prod-table thead") for k in ["單價", "供需", "庫存", "金額", "Supply", "剩餘可供貨"]) and pg.eval_on_selector_all("#bd-prod-table tbody tr", "els => els.length") >= 5)
        n_all = pg.eval_on_selector_all("#bd-prod-table tbody tr", "els => els.length")
        first_brand = pg.eval_on_selector("#bd-prod-table tbody tr .sub", "e => e.innerText.split('·')[0].trim()")
        pg.fill("#bd-q", first_brand); pg.wait_for_timeout(400)
        n_q = pg.eval_on_selector_all("#bd-prod-table tbody tr", "els => els.length")
        check("③ 搜品牌會過濾商品列", 0 < n_q < n_all and all(first_brand in t for t in pg.eval_on_selector_all("#bd-prod-table tbody tr .sub", "els => els.map(e => e.innerText)")), f"{n_q}/{n_all} {first_brand}")
        pg.fill("#bd-q", ""); pg.wait_for_timeout(300)
        pg.click('#bd-view button[data-v="brands"]'); pg.wait_for_timeout(300)
        check("③ 切到品牌：品牌表出現、商品表收起、搜尋框收起", pg.is_visible("#bd-brands") and not pg.is_visible("#bd-products") and not pg.is_visible("#bd-q") and pg.eval_on_selector_all("#bd-brand-table tbody tr", "els => els.length") >= 3)
        pg.click('#bd-brand-table tbody tr:first-child td.ed[data-k="target_giv"]'); pg.wait_for_timeout(200)
        pg.fill('#bd-brand-table td.ed input', "100000"); pg.keyboard.press("Enter"); pg.wait_for_timeout(1200)
        row1 = pg.inner_text("#bd-brand-table tbody tr:first-child")
        check("③ 品牌目標點一下填、存完顯示 100,000 與達成 %", "100,000" in row1 and "%" in row1 and "距目標 " in row1, row1[:160])
        pg.click('#bd-view button[data-v="products"]'); pg.wait_for_timeout(300)
        cell = '#bd-prod-table tbody tr:first-child td.ed[data-f="supply_cs"]'
        bc_e = pg.eval_on_selector("#bd-prod-table tbody tr:first-child", "e => e.dataset.bc")
        pg.click(cell); pg.wait_for_timeout(200); pg.fill(f"{cell} input", "4321"); pg.keyboard.press("Enter"); pg.wait_for_timeout(1200)
        row_e = pg.eval_on_selector(f'#bd-prod-table tbody tr[data-bc="{bc_e}"]', "e => e.innerText")
        check("③ 商品表點 Supply 格子直接改：存完顯示 4,321、剩餘可供貨跟著變、格子有橘色小點", "4,321" in row_e and pg.eval_on_selector_all(f'#bd-prod-table tbody tr[data-bc="{bc_e}"] .ed-mark', "els => els.length") == 1, row_e[:200])
        pg.click('.m-tab[data-tab="products"]'); pg.wait_for_timeout(400)
        pg.set_input_files("#file-prod", os.path.join(FAKE, "假_寶僑總表.xlsx")); pg.wait_for_timeout(300)
        pg.click("#prod-files-go"); pg.wait_for_selector("#dlg-sheet-conf[open]", timeout=30000)
        check("① 重匯舊總表：跳出「總表的數字比系統舊」，列出剛剛手改的 Supply", "4,321" in pg.inner_text("#sc-list") and "系統手動修改" in pg.inner_text("#sc-list"), pg.inner_text("#sc-head")[:120])
        pg.click("#sc-keep"); pg.wait_for_selector("#prod-imp-msg .bi-shield-check", timeout=30000)
        check("① 選「保留系統的」：匯完訊息寫保留了幾格", "保留系統數值" in pg.inner_text("#prod-imp-msg"))
        pg.click('.m-tab[data-tab="summary"]'); pg.wait_for_timeout(900)
        check("③ 回到總表，手改的 4,321 還在", "4,321" in pg.eval_on_selector(f'#bd-prod-table tbody tr[data-bc="{bc_e}"]', "e => e.innerText"))
        pg.click('#bd-view button[data-v="daily"]'); pg.wait_for_timeout(300)
        check("③ 切到每日出貨：原本的寬表", pg.is_visible("#bd-daily") and pg.eval_on_selector_all("#sum-table tbody tr", "els => els.length") >= 5)
        check("③ 總表的「月TTL」表頭字看得到（淺底深字，不是深藍底深字）", pg.evaluate("(() => { const th = [...document.querySelectorAll('#sum-table th')].find(t => t.innerText.includes('TTL')); if (!th) return false; const cs = getComputedStyle(th); const lum = c => { const m = c.match(/\\d+/g).map(Number); return (m[0] * 299 + m[1] * 587 + m[2] * 114) / 1000; }; return lum(cs.backgroundColor) > 180 && lum(cs.color) < 110; })()"))
        pg.click('#bd-view button[data-v="products"]'); pg.wait_for_timeout(200)
        pg.click('.m-tab[data-tab="orders"]'); pg.wait_for_timeout(400)
        pg.click("#btn-logs"); pg.wait_for_timeout(500); check("修改歷程視窗打開", pg.evaluate("document.getElementById('dlg-logs').open")); pg.keyboard.press("Escape")
        pg.click("#btn-reset"); pg.wait_for_timeout(300); check("清除資料視窗打開、按鈕預設鎖住", pg.is_disabled("#rs-go")); pg.keyboard.press("Escape")
        pg.click("#btn-upload"); pg.wait_for_timeout(300)
        pg.set_input_files("#file-import", os.path.join(FAKE, "假_訂單彙總表_9月_第一次.xlsx"))
        pg.wait_for_selector("#btn-commit:not(.hidden)", timeout=30000)
        check("上傳後預覽出現、檔案在暫存清單", pg.eval_on_selector_all("#imp-files-list .chip", "els => els.length") == 1)
        pg.keyboard.press("Escape")

        print("\n【6】匯出連結都能下載")
        cookies = {c["name"]: c["value"] for c in pg.context.cookies()}
        import urllib.request
        for sel in ["#btn-export-daily", "#btn-export-2", "#btn-export-daily-2", "#btn-export-stats"]:
            if sel != "#btn-export-daily":
                pg.click('.m-tab[data-tab="summary"]'); pg.wait_for_timeout(400)
            href = pg.get_attribute(sel, "href")
            req = urllib.request.Request(base + href, headers={"Cookie": "; ".join(f"{k}={v}" for k, v in cookies.items())})
            with urllib.request.urlopen(req) as resp:
                check(f"{sel} 下載 200 且是 Excel", resp.status == 200 and "spreadsheet" in resp.headers.get("Content-Type", ""))
        print("\n【6b】瑪氏出貨：拆單、產出、回填 EIP 採購單號、瑪氏採購單")
        mars_master = os.path.join(BASE, "samples", "mars", "fake", "假_瑪氏商品總表.xlsx")
        assert tc.post("/api/mars/products/import", data={"file": (io.BytesIO(open(mars_master, "rb").read()), "假_瑪氏商品總表.xlsx")},
                       content_type="multipart/form-data").status_code == 200
        pg.goto(f"{base}/mars"); pg.wait_for_selector("#sp-table thead"); pg.wait_for_timeout(500)
        check("瑪氏出貨頁打得開、商品總表狀態寫出 7 個料號", "7" in pg.inner_text("#mp-status") and "假_瑪氏商品總表" in pg.inner_text("#mp-status"), pg.inner_text("#mp-status")[:120])
        pg.click("#btn-fe"); pg.wait_for_selector("#dlg-fe[open]", timeout=5000)
        check("拆單旁「底稿填 EIP 採購單號」→ 打開視窗、有拖曳框", pg.is_visible("#fe-dz") and "單號＋SKU No." in pg.inner_text("#dlg-fe"))
        wbf = openpyxl.Workbook(); wsf = wbf.active; wsf.title = "拆解訂單"
        wsf.append(["交期", "想改交期", "倉庫", "想改倉庫", "單號", "No.", "SKU No.", "條碼", "", "永豐料號", "品名"])
        wsf.append(["2026-10-21", None, "TAO5", None, "99999999999999", 1, "123", "", None, "M00000000", "對不到的品"])
        fp = os.path.join(tempfile.gettempdir(), "ui_拆單底稿.xlsx"); wbf.save(fp)
        pg.set_input_files("#fe-file", fp); pg.wait_for_selector("#fe-out .fe-cnt", timeout=10000)
        check("丟底稿 → 列出幾列、對不到的列出來標黃、沒有可填的就不能下載", "對不到" in pg.inner_text("#fe-out") and "系統裡找不到" in pg.inner_text("#fe-out")
              and pg.is_disabled("#fe-go"), pg.inner_text("#fe-out")[:200])
        pg.click("#fe-x")
        pg.wait_for_selector("#mp-counts .chip", timeout=5000)
        cnt = " ".join(pg.inner_text("#mp-counts").split())          # chip 是 inline-flex，inner_text 會把數字和字拆成兩行
        check("商品總表卡片上一排數字：7 個料號、有換號／停產／對到兩個料號／補登；清單平常收起來", "7 個料號" in cnt and "停產 0" in cnt
              and "對到兩個料號 0 組" in cnt and pg.is_hidden("#mp-res"), cnt)
        pg.fill("#mp-q", "m36752197"); pg.wait_for_selector("#mp-res table", timeout=5000)
        check("搜尋：打小寫料號 → 一個料號一列、單位寫箱 盒 包、共 1 個料號", pg.eval_on_selector_all("#mp-res tbody tr", "els => els.length") == 1
              and "箱 盒 包" in pg.inner_text("#mp-res") and "共 1 個料號" in pg.inner_text("#mp-res"), pg.inner_text("#mp-res")[:200])
        pg.fill("#mp-q", "M60013313"); pg.wait_for_function("document.querySelector('#mp-res').innerText.includes('沒有')", timeout=5000)
        check("查不到 → 寫「商品總表裡沒有」並說要補進 Alice 的總表再上傳", "商品總表裡沒有「M60013313」" in pg.inner_text("#mp-res") and "重新上傳" in pg.inner_text("#mp-res"))
        pg.fill("#mp-q", ""); pg.wait_for_timeout(400)
        pg.click("#mp-counts .chip[data-f=dead]"); pg.wait_for_function("document.querySelector('#mp-res').innerText.includes('這一類沒有料號')", timeout=5000)
        check("點卡片上的「停產」→ 下面展開那一類（假檔沒有停產 → 寫這一類沒有料號）、篩選鈕亮起來", pg.eval_on_selector("#mp-filters .chip.on", "e => e.dataset.f") == "dead")
        pg.click("#mp-filters .chip[data-f=all]"); pg.wait_for_selector("#mp-res table", timeout=5000)
        check("點「全部」→ 7 個料號一列一個、有分頁說明", pg.eval_on_selector_all("#mp-res tbody tr", "els => els.length") == 7 and "第 1／1 頁" in pg.inner_text("#mp-pages"))
        pg.click("#lm-btn"); pg.wait_for_timeout(200)
        check("線別工具：瑪氏那欄有「瑪氏出貨」而且標目前在這", pg.eval_on_selector("#lm-pop a.lm-item.on", "e => e.innerText").startswith("瑪氏出貨"))
        pg.keyboard.press("Escape")
        pg.set_input_files("#od-file", os.path.join(BASE, "samples", "master", "訂單彙總表範例.xlsx")); pg.wait_for_selector("#od-preview:not(.hidden)", timeout=30000)
        check("從瑪氏頁傳只有寶僑的檔 → 擋下來說沒有瑪氏的單，不匯", "沒有瑪氏訂單" in pg.inner_text("#od-preview") and pg.eval_on_selector_all("#od-commit", "els => els.length") == 0)
        pg.set_input_files("#od-file", os.path.join(FAKE, "假_訂單彙總表_9月_第二次.xlsx")); pg.wait_for_selector("#od-commit", timeout=30000)
        check("傳有瑪氏的檔 → 預覽寫瑪氏幾列、另有哪些線別", "瑪氏" in pg.inner_text("#od-preview") and "另有" in pg.inner_text("#od-preview"), pg.inner_text("#od-preview")[:160])
        pg.click("#od-commit"); pg.wait_for_selector("#od-msg .bi-check-circle", timeout=30000); pg.wait_for_timeout(600)
        check("確認匯入：訂單狀態寫出最近一次匯入", "最近一次匯入" in pg.inner_text("#od-status") and "假_訂單彙總表_9月_第二次" in pg.inner_text("#od-status"), pg.inner_text("#od-status")[:160])
        check("匯完自動跳到這批的到貨日範圍（9/1 開頭那個月）", pg.input_value("#d-from") < pg.input_value("#d-to") and pg.input_value("#d-from").startswith("2026-09"), f"{pg.input_value('#d-from')}～{pg.input_value('#d-to')}")
        check("同一張 PO 的幾份用 PO 標題列包起來（標題列數＝PO 張數）", pg.eval_on_selector_all("#sp-table tr.pog", "els => els.length") == len(set(pg.eval_on_selector_all("#sp-table tr.sp td:nth-child(5)", "els => els.map(e => e.innerText)"))) > 3)
        check("月曆：9 月有訂單的天亮起來、格子裡有箱數和 PO 數", pg.eval_on_selector_all("#cal .day.has", "els => els.length") >= 2 and "箱" in pg.inner_text("#cal .day.has >> nth=0") and "PO" in pg.inner_text("#cal .day.has >> nth=0"), pg.inner_text("#cal .day.has >> nth=0")[:60])
        pg.click("#cal .day[data-date='2026-09-18']"); pg.wait_for_timeout(800)
        check("點月曆 9/18 → 起訖都變 9/18、那格變深色", pg.input_value("#d-from") == "2026-09-18" and pg.input_value("#d-to") == "2026-09-18" and pg.eval_on_selector("#cal .day[data-date='2026-09-18']", "e => e.classList.contains('on')"))
        check("五個上傳處都有看得見的虛線拖曳框（多了 ⑥ 驗收單）", pg.eval_on_selector_all(".dzbox:not(#fe-dz)", "els => els.length") == 5)
        check("9/18：5 份拆單表、都是「還沒產出」，但單號格子和每列按鈕已經能用、一條 PO 標題列寫拆成 5 份", pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length") == 5
              and pg.eval_on_selector_all("#sp-table .st-new", "els => els.length") == 5 and not pg.is_disabled("#sp-table input.eip")
              and pg.eval_on_selector_all("#sp-table button.eipf", "els => els.length") == 5
              and pg.eval_on_selector_all("#sp-table tr.pog", "els => els.length") == 1 and "拆成 5 份" in pg.inner_text("#sp-table tr.pog"))
        pg.click("#sp-table tr.sp >> nth=0"); pg.wait_for_timeout(300)
        check("點一列展開看品項（下採料號、箱數）", pg.eval_on_selector_all("#sp-table tr.sub", "els => els.length") == 1 and "下採料號" in pg.inner_text("#sp-table tr.sub"))
        # Jerry 2026-10-13：畫面上標清楚哪個是原廠的料號、哪個是 EIP 用的（只改畫面，匯出不動）
        sub_head = pg.eval_on_selector_all("#sp-table tr.sub thead th", "els => els.map(e => e.innerText.trim())")
        check("展開的品項表頭標示「永豐料號（原廠）」「下採料號（EIP）」", sub_head[:2] == ["永豐料號（原廠）", "下採料號（EIP）"], str(sub_head[:3]))
        # 搜尋（Jerry 2026-10-08：「你居然沒有做查詢功能」）：打料號 → 只剩有那個料號的幾份、而且自動展開；找別天的 PO → 沒有，按一下改成全部日期
        sku = pg.inner_text("#sp-table tr.sub td.fn >> nth=0").strip()
        pg.fill("#q", sku); pg.wait_for_timeout(400)
        n_sp, n_sub = pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length"), pg.eval_on_selector_all("#sp-table tr.sub", "els => els.length")
        check(f"搜尋料號 {sku} → 只剩有這個料號的幾份、每份都自動展開、上方寫符合幾份", 1 <= n_sp < 5 and n_sub == n_sp and sku in pg.inner_text("#sp-table tr.sub") and "符合" in pg.inner_text("#sum"), f"{n_sp} 份、展開 {n_sub}")
        pg.fill("#q", "13000000600016"); pg.wait_for_timeout(400)
        check("搜尋別天的 PO → 這期間沒有、出現「改成全部日期再找」", pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length") == 0 and "沒有符合" in pg.inner_text("#sp-table") and pg.is_visible("#q-all"))
        pg.click("#q-all"); pg.wait_for_selector("#sp-table tr.pog", timeout=10000); pg.wait_for_timeout(300)
        check("按下去 → 日期改成全部、找到那張 PO", pg.input_value("#d-from") < "2026-09-18" and "13000000600016" in pg.inner_text("#sp-table tr.pog") and pg.eval_on_selector_all("#sp-table tr.pog", "els => els.length") == 1, f"{pg.input_value('#d-from')}～{pg.input_value('#d-to')}")
        pg.fill("#q", ""); pg.wait_for_timeout(400)
        wh_col = lambda: pg.eval_on_selector_all("#sp-table tr.sp td:nth-child(6)", "els => els.map(e => e.innerText.trim())")   # noqa: E731
        date_order = wh_col()
        pg.select_option("#f-sort", "wh"); pg.wait_for_timeout(400)
        whs_sorted, pos_seen = wh_col(), pg.eval_on_selector_all("#sp-table tr.pog", "els => els.map(e => (e.innerText.match(/PO\\s*(\\d{14})/) || [])[1] + '@' + (e.innerText.match(/\\d+\\/\\d+ 到貨/) || [''])[0])")
        check("排序選「依倉別」→ 倉那欄由小到大、同一張 PO 還是只有一條標題列", len(set(whs_sorted)) > 1 and whs_sorted == sorted(whs_sorted, key=lambda w: (w.rstrip("0123456789"), int("0" + w[len(w.rstrip("0123456789")):])))
              and whs_sorted != date_order and len(pos_seen) == len(set(pos_seen)), f"{whs_sorted} / 原本 {date_order} / 標題 {pos_seen}")
        pg.reload(); pg.wait_for_selector("#sp-table tr.sp", timeout=15000); pg.wait_for_timeout(500)
        check("重新整理後還記得依倉別排序", pg.input_value("#f-sort") == "wh")
        pg.select_option("#f-sort", "date"); pg.wait_for_timeout(300)
        pg.click("#cal .day[data-date='2026-09-18']"); pg.wait_for_timeout(800)
        check("清掉搜尋、點回 9/18 → 5 份都回來", pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length") == 5 and "符合" not in pg.inner_text("#sum"))
        with pg.expect_download() as dl:
            pg.click("#sp-table button.eipf >> nth=1")
        pg.wait_for_timeout(900)
        check("還沒產出的列直接按 EIP → 先把那一份存起來再下載（只有那一份變已產出）", dl.value.suggested_filename.endswith("_EIP上傳.xls")
              and pg.eval_on_selector_all("#sp-table .st-generated", "els => els.length") == 1 and pg.eval_on_selector_all("#sp-table .st-new", "els => els.length") == 4, dl.value.suggested_filename)
        with pg.expect_download() as dl:
            pg.click("#btn-gen")
        check("按「下載 EIP 採購單」下載 zip（瑪氏EIP採購單_到貨日）", dl.value.suggested_filename == "瑪氏EIP採購單_20260918.zip" and "下載 EIP 採購單" in pg.inner_text("#btn-gen"), dl.value.suggested_filename)
        pg.wait_for_timeout(900)
        check("產出後狀態變「已產出」、單號格子可以填", pg.eval_on_selector_all("#sp-table .st-generated", "els => els.length") == 5 and not pg.is_disabled("#sp-table input.eip"))
        pg.fill("#sp-table input.eip >> nth=0", "PO123"); pg.keyboard.press("Enter"); pg.wait_for_timeout(700)
        check("單號格式不對 → 格子變紅、不存", pg.eval_on_selector("#sp-table input.eip", "e => e.classList.contains('bad')"))
        pg.fill("#sp-table input.eip >> nth=0", "PO202609901"); pg.keyboard.press("Enter"); pg.wait_for_timeout(900)
        check("填對 → 狀態「已回填」", pg.eval_on_selector_all("#sp-table .st-filled", "els => els.length") == 1)
        check("只有一份用這個 EIP 單號 → 沒有「合併」鈕", pg.eval_on_selector_all("#sp-table button.po-merge", "els => els.length") == 0)
        # 照訂單彙總表的表頭做一份只有兩列瑪氏的檔（跟 test_mars 的 special_sheet 同做法；不 import test_mars，它一載入就會換資料庫）
        wbm = openpyxl.load_workbook(os.path.join(FAKE, "假_訂單彙總表_9月_第一次.xlsx")); wsm = wbm.worksheets[0]
        hdr = [x.value for x in wsm[1]]; wsm.delete_rows(2, wsm.max_row)
        for sku, bc, yf, nm, q, u, bx in [("910000000000011", "4710000000011", "M69072565", "喵愛餡 盒", 24, "盒", 12),
                                          ("910000000000012", "4710000000012", "M80676511", "喵喵鮮 包", 48, "包", 24)]:
            row = {"訂單類型": "一般", "線別": "瑪氏", "PO單號": "13000000583199", "到貨倉別": "TAO3", "地址": "桃園市大園區中山南路472號",
                   "交付日期": datetime.datetime(2026, 10, 12), "NO": 1, "品牌": "測試", "酷澎下單單價(含稅)": 100, "SKU ID": sku, "條碼(國條)": bc,
                   "永豐料號": yf, "品名": nm, "下單數量(酷澎單位)": q, "出貨數量": q, "單位": u, "箱入數": bx}
            wsm.append([row.get(h) for h in hdr])
        bm = io.BytesIO(); wbm.save(bm)
        pv = tc.post("/api/master/import/preview", data={"file": (io.BytesIO(bm.getvalue()), "盒包合併.xlsx")}, content_type="multipart/form-data").get_json()
        check("另匯一張 10/12 的單：盒、包同品類同中標（測合併瑪氏採購單）", tc.post("/api/master/import/commit", json={"batch_id": pv["batch_id"], "add_missing_products": True}).status_code == 200)
        pg.fill("#d-from", "2026-10-12"); pg.fill("#d-to", "2026-10-12"); pg.dispatch_event("#d-to", "change"); pg.wait_for_selector("#sp-table tr.sp", timeout=10000); pg.wait_for_timeout(500)
        pg.fill("#sp-table input.eip >> nth=0", "PO202610903"); pg.keyboard.press("Enter"); pg.wait_for_timeout(1200)
        n_merge = pg.eval_on_selector_all("#sp-table button.po-merge", "els => els.length")
        check("盒那列填 PO202610903 → 包一起填上、兩列都多一顆「合併」", n_merge == 2 and pg.eval_on_selector_all("#sp-table input.eip", "els => els.map(e => e.value)") == ["PO202610903", "PO202610903"], str(n_merge))
        with pg.expect_download() as dl:
            pg.click("#sp-table button.po-merge >> nth=1")
        check("按「合併」→ 下載一張瑪氏採購單，檔名 PO202610903 開頭、單位寫(盒+包)", dl.value.suggested_filename.startswith("PO202610903永豐Mars採購單") and dl.value.suggested_filename.endswith("(盒+包).xlsx"), dl.value.suggested_filename)
        pg.click("#cal .day[data-date='2026-09-18']") if pg.query_selector("#cal .day[data-date='2026-09-18']") else None
        pg.fill("#d-from", "2026-09-18"); pg.fill("#d-to", "2026-09-18"); pg.dispatch_event("#d-to", "change"); pg.wait_for_timeout(900)
        pg.click("#f-noeip"); pg.wait_for_timeout(300)
        check("「只看還沒填 EIP 單號的」→ 剩 4 份、填了的那份不見", pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length") == 4 and pg.eval_on_selector_all("#sp-table .st-filled", "els => els.length") == 0 and "符合 4／5" in pg.inner_text("#sum"), pg.inner_text("#sum"))
        pg.click("#f-noeip"); pg.wait_for_timeout(300)
        check("再按一次 → 5 份都回來", pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length") == 5)
        pg.fill("#sp-table input.slot >> nth=0", "12:30~15:30（1台車）"); pg.keyboard.press("Enter"); pg.wait_for_timeout(900)
        check("約倉時間存進去、重新整理還在", pg.eval_on_selector("#sp-table input.slot", "e => e.value") == "12:30~15:30（1台車）")
        # ③ 瑪氏採購單：有 EIP 單號就能按；倉庫資料沒填只提示會留空；到採購單設定填 TAO1 → 提示消失 → 下載
        check("採購單按鈕：單號填了就能按，TAO1 倉庫資料還缺只在提示寫會留空、上面有提醒", not pg.is_disabled("#sp-table button.po >> nth=0")
              and "TAO1" in pg.get_attribute("#sp-table button.po >> nth=0", "title") and "留空" in pg.get_attribute("#sp-table button.po >> nth=0", "title") and "TAO1" in pg.inner_text("#alerts"), pg.get_attribute("#sp-table button.po >> nth=0", "title"))
        check("其他沒填單號的那幾份：按鈕鎖住、提示寫還沒填 EIP 採購單號", "EIP" in pg.get_attribute("#sp-table button.po >> nth=1", "title"))
        pg.click("#ps-details summary"); pg.wait_for_selector("#wh-table input.wh"); pg.wait_for_timeout(200)
        check("採購單設定：倉庫清單有 TAO1、地址已從訂單帶入、標缺電話與 ship-to", "缺 電話、ship-to" in pg.inner_text("#wh-table tr[data-code='TAO1']")
              and pg.input_value("#wh-table tr[data-code='TAO1'] input[data-f='address']") != "", pg.inner_text("#wh-table tr[data-code='TAO1']")[:120])
        pg.fill("#wh-table tr[data-code='TAO1'] input[data-f='phone']", "02-5592-7598"); pg.fill("#wh-table tr[data-code='TAO1'] input[data-f='ship_to']", "17617037"); pg.keyboard.press("Enter"); pg.wait_for_timeout(900)
        check("填電話與 ship-to 離開格子就存 → 那倉變「齊了」、提醒消失", "已填齊" in pg.inner_text("#wh-table tr[data-code='TAO1']") and "TAO1" not in pg.inner_text("#alerts"), pg.inner_text("#wh-table tr[data-code='TAO1']")[:120])
        check("按鈕提示不再寫會留空、統計寫 1／5 可產瑪氏採購單", "留空" not in pg.get_attribute("#sp-table button.po >> nth=0", "title") and "1／5 可產出瑪氏採購單" in pg.inner_text("#sum").replace("\n", ""), pg.inner_text("#sum"))
        with pg.expect_download() as dl:
            pg.click("#sp-table button.po >> nth=0")
        check("按採購單 → 下載 PO202609901永豐Mars採購單(箱單位)-…_TAO1_PO(盒).xlsx", dl.value.suggested_filename.startswith("PO202609901永豐Mars採購單(箱單位)-") and dl.value.suggested_filename.endswith("_TAO1_13000000600028(盒).xlsx"), dl.value.suggested_filename)
        with pg.expect_download() as dl:
            pg.click("#btn-po")
        check("「瑪氏採購單（全部）」→ zip", dl.value.suggested_filename == "瑪氏採購單_20260918.zip", dl.value.suggested_filename)
        # ④ EMMA 匯入檔：單份、勾選合併、全部
        check("每列有 EMMA 鈕；沒填單號的帶驚嘆號、填了的沒有", pg.eval_on_selector_all("#sp-table button.emma", "els => els.length") == 5
              and pg.eval_on_selector_all("#sp-table button.emma .bi-exclamation-circle", "els => els.length") == 4, str(pg.eval_on_selector_all("#sp-table button.emma", "els => els.map(e => e.innerText)")))
        with pg.expect_download() as dl:
            pg.click("#sp-table button.emma >> nth=0")
        check("按 EMMA（已填單號的）→ 直接下載 酷澎訂單匯入_0918交貨-TAO1_PO_…", dl.value.suggested_filename.startswith("酷澎訂單匯入_0918交貨-TAO1_13000000600028_"), dl.value.suggested_filename)
        check("沒勾選時「勾選的合併」鎖住", pg.is_disabled("#btn-emma-sel"))
        pg.check("#sp-table input.ck >> nth=0"); pg.check("#sp-table input.ck >> nth=1"); pg.wait_for_timeout(200)
        check("勾兩份 → 按鈕解鎖、寫（2）", not pg.is_disabled("#btn-emma-sel") and "（勾選 2 份）" in pg.inner_text("#btn-emma-sel"), pg.inner_text("#btn-emma-sel"))
        pg.once("dialog", lambda d: d.accept())
        with pg.expect_download() as dl:
            pg.click("#btn-emma-sel")
        check("勾選合併（有一份沒單號會先問，答是）→ 下載 酷澎訂單匯入_0918交貨-TAO1.xlsx", dl.value.suggested_filename == "酷澎訂單匯入_0918交貨-TAO1.xlsx", dl.value.suggested_filename)
        pg.once("dialog", lambda d: d.accept())
        with pg.expect_download() as dl:
            pg.click("#btn-emma-all")
        check("EMMA 檔（全部）→ 同樣檔名（同一天同一倉）", dl.value.suggested_filename == "酷澎訂單匯入_0918交貨-TAO1.xlsx", dl.value.suggested_filename)
        wbw = openpyxl.Workbook(); wsw = wbw.active; wsw.append(["倉別", "中文地址", "電話"]); wsw.append(["TAO5", "桃園市觀音區寶倉街108號5樓", "+886-0911556291"]); wsw.append(["TAO9", "桃園市大園區建國路102號3樓", None])
        whp = os.path.join(tempfile.gettempdir(), "ui_倉庫.xlsx"); wbw.save(whp)
        pg.set_input_files("#wh-file", whp); pg.wait_for_selector("#wh-msg .bi-check-circle", timeout=15000); pg.wait_for_timeout(500)
        check("上傳倉庫資料表：寫出幾個倉新增幾個、TAO9 出現在清單、電話轉成 0911-556-291", "2 個倉" in pg.inner_text("#wh-msg") and pg.eval_on_selector_all("#wh-table tr[data-code='TAO9']", "els => els.length") == 1
              and pg.input_value("#wh-table tr[data-code='TAO5'] input[data-f='phone']") == "0911-556-291", pg.inner_text("#wh-msg"))
        pg.fill("#wh-new-code", "tao8"); pg.fill("#wh-new input[data-f='address']", "桃園市某處"); pg.click("#wh-add"); pg.wait_for_timeout(800)
        check("最下面直接加一個倉 TAO8 → 出現在清單、有刪除鈕；訂單有的倉沒有刪除鈕", pg.eval_on_selector_all("#wh-table tr[data-code='TAO8'] .wh-del", "els => els.length") == 1
              and pg.eval_on_selector_all("#wh-table tr[data-code='TAO1'] .wh-del", "els => els.length") == 0, pg.inner_text("#wh-table tr[data-code='TAO8']")[:80] if pg.query_selector("#wh-table tr[data-code='TAO8']") else "沒有 TAO8")
        pg.once("dialog", lambda d: d.accept()); pg.click("#wh-table tr[data-code='TAO8'] .wh-del"); pg.wait_for_timeout(800)
        check("按刪除、確認 → TAO8 不見", pg.eval_on_selector_all("#wh-table tr[data-code='TAO8']", "els => els.length") == 0)
        # ⑤ 勇信缺貨：假 PDF（第一份拆單表 PO202609901，一個品項少 1 箱）→ 比對表 → 確認 → 下修檔
        import pymupdf
        sp0 = next(x for x in tc.get("/api/mars/splits?from=2026-09-18&to=2026-09-18").get_json()["splits"] if x["eip_po"] == "PO202609901")
        doc = pymupdf.open(); page = doc.new_page(); y = 40
        lines = ["BMS100222-永豐商店酷澎-TAO1", "指送日期：2026/09/18", "收貨單號：PO202609901-酷澎-T"]
        for n, it in enumerate(sp0["items"], start=1):
            q = max(int(it["cases"]) - (1 if n == 1 else 0), 0)
            lines += [f"58880{it['mars_code']}", str(n), f"假品名 1:6:10", f"{q}C  ", "625D1TAP01", "2027/06/18"]
        lines += [f"合計箱數:{sum(max(int(it['cases']) - (1 if n == 1 else 0), 0) for n, it in enumerate(sp0['items'], start=1))}"]
        for ln in lines:
            page.insert_text((40, y), ln, fontname="china-t", fontsize=9); y += 14
        yxp = os.path.join(tempfile.gettempdir(), "ui_勇信.pdf"); doc.save(yxp)
        # 主管 2026-10-08：核不到就是缺。第二份填一個 PDF 裡沒有的號碼 → 那份算全缺；其他沒填號碼的列紅框（無法核對）
        pg.fill("#sp-table input.eip >> nth=1", "PO202609902"); pg.keyboard.press("Enter"); pg.wait_for_timeout(900)
        pg.set_input_files("#yx-file", yxp); pg.wait_for_selector("#yx-result:not(.hidden)", timeout=20000); pg.wait_for_timeout(400)
        row = pg.inner_text("#yx-result tr:has(input.yx-scope)") if pg.query_selector("#yx-result input.yx-scope") else ""
        check("「這次要出的訂單」：那天那個倉的 PO 一列、預設勾、寫表上有幾份、表上沒有幾份算全缺、幾份還沒填 EIP 單號", pg.eval_on_selector_all("#yx-result input.yx-scope", "els => els.length") == 1
              and pg.is_checked("#yx-result input.yx-scope") and "表上有 1 份" in row and "表上沒有 1 份，算全缺" in row and "還沒填 EIP 單號" in row, row)
        check("沒填 EIP 單號的份數列在紅框（無法核對、不改）、PO202609902 那份在比對表標表上沒有", "還沒填 EIP 採購單號，無法跟勇信配送明細表核對" in pg.inner_text("#yx-result .al-bad")
              and "表上沒有" in pg.inner_text("#yx-result table.t >> nth=1") and "改 1 筆" not in pg.inner_text("#yx-apply"), pg.inner_text("#yx-apply"))
        pg.eval_on_selector("#yx-result input.yx-scope", "e => e.scrollIntoView({block: 'center'})"); pg.wait_for_timeout(200); y0 = pg.evaluate("window.scrollY")
        pg.uncheck("#yx-result input.yx-scope"); pg.wait_for_timeout(900)
        check("點勾勾畫面不會跳回上面（位置差不到 5 像素）", abs(pg.evaluate("window.scrollY") - y0) < 5, f"{y0} → {pg.evaluate('window.scrollY')}")
        check("取消勾選（這次不出）→ 那列寫這次不核、比對表沒有、確認鈕鎖住", not pg.is_checked("#yx-result input.yx-scope") and "這次不核" in pg.inner_text("#yx-result") and pg.is_disabled("#yx-apply"))
        pg.check("#yx-result input.yx-scope"); pg.wait_for_timeout(900)
        check("再勾回來 → 又核了", pg.is_checked("#yx-result input.yx-scope") and not pg.is_disabled("#yx-apply"))
        pg.click("#yx-clear")
        pg.fill("#sp-table input.eip >> nth=1", ""); pg.keyboard.press("Enter"); pg.wait_for_timeout(900)          # 清掉 902，下面照原本只測部分缺那一筆
        pg.set_input_files("#yx-file", yxp); pg.wait_for_selector("#yx-result:not(.hidden)", timeout=20000); pg.wait_for_timeout(400)
        check("上傳勇信 PDF → 比對表：1 張 PO、有一筆部分缺、確認鈕寫改 1 筆", "對到 1 張 PO" in pg.inner_text("#yx-result").replace("\n", "") and pg.eval_on_selector_all("#yx-result .st-changed", "els => els.length") >= 1
              and "改 1 筆" in pg.inner_text("#yx-apply"), pg.inner_text("#yx-apply"))
        pg.once("dialog", lambda d: d.accept())
        with pg.expect_download() as dl:
            pg.click("#yx-apply")
        check("確認 → 下載 酷澎下修_0918交貨-TAO1.xlsx、比對區收起來、寫已確認", dl.value.suggested_filename == "酷澎下修_0918交貨-TAO1.xlsx", dl.value.suggested_filename)
        pg.wait_for_timeout(900)
        check("確認後訊息寫系統數量已更新", "已確認" in pg.inner_text("#yx-msg"), pg.inner_text("#yx-msg"))
        # ⑥ 驗收單缺貨劃線（Jerry 2026-10-08）：那張 PO 的驗收單照系統出貨數量劃線；做一張假的驗收單（表格照真檔欄寬）
        po_r = sp0["po_number"]
        cur_r = {o["sku_id"]: o["qty_ship"] for o in tc.get(f"/api/master/orders?month=2026-09&q={po_r}").get_json()["rows"]}
        rows_r = [(n, sku, f"假品名{n}", int(q) + (6 if n == 1 else 0)) for n, (sku, q) in enumerate([kv for kv in cur_r.items() if kv[1]], start=1)][:3]
        doc = pymupdf.open(); page = doc.new_page(width=595, height=842)
        page.insert_text((60, 120), f"採購單號 {po_r}", fontname="china-t", fontsize=9)
        xs = [37, 62, 114, 167, 271, 313, 355, 400, 480, 559]; ys = [380]
        for i, t_ in enumerate(["No", "SKU No.", "條碼", "品名", "出貨數量", "實際數量", "管理", "允收效期", "差異原因"]):
            page.insert_text((xs[i] + 2, 393), t_, fontname="china-t", fontsize=7)
        ys.append(400)
        for no, sku, nm, q in rows_r:
            top = ys[-1]
            page.insert_text((xs[0] + 8, top + 16), str(no), fontsize=8); page.insert_text((xs[1] + 2, top + 11), sku[:11], fontsize=7)
            page.insert_text((xs[1] + 2, top + 21), sku[11:], fontsize=7); page.insert_text((xs[3] + 2, top + 16), nm, fontname="china-t", fontsize=7)
            page.insert_text((xs[4] + 14, top + 16), str(q), fontsize=8); ys.append(top + 28)
        page.insert_text((xs[0] + 5, ys[-1] + 13), "合計", fontname="china-t", fontsize=8); ys.append(ys[-1] + 20)
        for yy in ys:
            page.draw_line((xs[0], yy), (xs[-1], yy), width=0.6)
        for xx in xs:
            page.draw_line((xx, ys[0]), (xx, ys[-1]), width=0.6)
        rcp = os.path.join(tempfile.gettempdir(), f"signed_PO_DELIVERY_INVOICE_{po_r}.pdf"); doc.save(rcp)
        pg.set_input_files("#rc-file", rcp); pg.wait_for_selector("#rc-result:not(.hidden)", timeout=15000); pg.wait_for_timeout(300)
        check("⑥ 丟簽好名的驗收單 → 先列出：改數量 1 品、原本幾個 → 系統幾個", "改數量 1 品" in pg.inner_text("#rc-result") and f"{rows_r[0][3]} → {cur_r[rows_r[0][1]]}" in pg.inner_text("#rc-result").replace(",", ""),
              pg.inner_text("#rc-result")[:200])
        with pg.expect_download() as dl:
            pg.click("#rc-go")
        check("按下載 → 一份就直接是 PDF、檔名照原檔", dl.value.suggested_filename == os.path.basename(rcp), dl.value.suggested_filename)
        pg.click("#hs-open"); pg.wait_for_selector("#hs-table tbody tr", timeout=10000); pg.wait_for_timeout(300)
        check("瑪氏頁「歷程」視窗：看得到剛才的 EIP 單號、勇信缺貨改數量", "EIP 採購單號" in pg.inner_text("#hs-table") and "出貨數量" in pg.inner_text("#hs-table") and "勇信出" in pg.inner_text("#hs-table"), pg.inner_text("#hs-count"))
        pg.fill("#hs-q", "PO202609901"); pg.click("#hs-go"); pg.wait_for_timeout(500)
        check("歷程搜尋：打單號只剩那幾筆", pg.eval_on_selector_all("#hs-table tbody tr", "els => els.length") >= 1 and "PO202609901" in pg.inner_text("#hs-table") and "出貨數量" not in pg.inner_text("#hs-table").replace("EIP 採購單號", ""), pg.inner_text("#hs-count"))
        pg.click("#hs-close"); pg.wait_for_timeout(200)
        pg.fill("#wh-table tr[data-code='TAO1'] input[data-f='special_note']", "司機需加入TAO1 line領取排隊號碼"); pg.keyboard.press("Enter"); pg.wait_for_timeout(800)
        pg.fill("#ps-holidays", "2026-09-17\n9/16"); pg.fill("#ps-shelf-GUM", "2/3效期以上"); pg.click("#ps-save"); pg.wait_for_timeout(900)
        pg.reload(); pg.wait_for_selector("#ps-holidays", state="attached"); pg.wait_for_timeout(800); pg.click("#ps-details summary")
        check("設定裡寫內建國定假日涵蓋哪幾年", "2026、2027 年" in pg.inner_text("#ps-builtin-years"), pg.inner_text("#ps-builtin-years"))
        check("假日、效期（GUM 改 2/3，CHO 維持預設 3/5）、倉庫加註存了、重新整理還在", pg.input_value("#ps-holidays") == "2026-09-16\n2026-09-17" and pg.input_value("#ps-shelf-GUM") == "2/3效期以上" and pg.input_value("#ps-shelf-CHO") == "3/5效期以上"
              and pg.input_value("#wh-table tr[data-code='TAO1'] input[data-f='special_note']") == "司機需加入TAO1 line領取排隊號碼", pg.input_value("#ps-holidays"))
        print("\n【6c】竹運出貨拋檔頁")
        pg.goto(f"{base}/zhuyun"); pg.wait_for_selector("#dz"); pg.wait_for_timeout(300)
        check("竹運頁打得開、有虛線拖曳框、預設收件資料顯示在上面", pg.eval_on_selector_all(".dzbox", "els => els.length") == 2 and "酷澎股份有限公司" in pg.inner_text("#s-recipient"))
        pg.click("#btn-ph"); pg.wait_for_selector("#ph-box:not(.hidden)")
        pg.fill("#ph-new-code", "txrc29"); pg.fill("#ph-new-phone", "+886-986368794"); pg.click("#ph-add"); pg.wait_for_timeout(500)
        check("倉別手機表：頁面上直接新增 txrc29 → 存成 TXRC29／0986368794", "TXRC29" in pg.inner_text("#ph-rows") and pg.eval_on_selector("#ph-rows tr[data-code='TXRC29'] .ph-in", "e => e.value") == "0986368794", pg.inner_text("#ph-rows"))
        pg.fill("#ph-rows tr[data-code='TXRC29'] .ph-in", "0911-222-333"); pg.press("#ph-rows tr[data-code='TXRC29'] .ph-in", "Enter"); pg.wait_for_timeout(500)
        check("在格子裡改手機、按 Enter 就存（0911-222-333 → 0911222333）", pg.eval_on_selector("#ph-rows tr[data-code='TXRC29'] .ph-in", "e => e.value") == "0911222333")
        pg.click("#ph-close")
        pg.set_input_files("#file", os.path.join(BASE, "samples", "整合表範例.xlsx")); pg.wait_for_selector("#result:not(.hidden)", timeout=30000); pg.wait_for_timeout(400)
        check("傳訂單彙總表範例 → 兩組（8/23 TAO5、TXRC29）、每組有兩個檔名和預覽表", pg.eval_on_selector_all("#groups .grp", "els => els.length") == 2 and "指定到貨日20260823,酷澎_PG(TAO5).xlsx" in pg.inner_text("#groups"), pg.inner_text("#sum"))
        check("TXRC29 那組手機用倉別手機表的 0911222333，TAO5 用預設手機", "0911222333（倉別手機表）" in pg.inner_text("#groups") and "02-55927598（預設手機）" in pg.inner_text("#groups"), pg.inner_text("#groups .grp .gh"))
        pg.fill("#q", "ARIEL"); pg.wait_for_timeout(300)
        check("竹運搜尋 ARIEL → 兩組都有、上方寫符合幾列、沒有「僅顯示前 60 列」", pg.eval_on_selector_all("#groups .grp", "els => els.length") == 2 and "符合" in pg.inner_text("#sum") and pg.eval_on_selector_all("#groups tbody tr", "els => els.length") >= 3 and "僅顯示" not in pg.inner_text("#groups"), pg.inner_text("#sum"))
        pg.fill("#q", "txrc29 4987176340900"); pg.wait_for_timeout(300)
        check("兩個關鍵字（倉＋料號）→ 只剩 TXRC29 那組的 1 列", pg.eval_on_selector_all("#groups .grp", "els => els.length") == 1 and pg.eval_on_selector_all("#groups tbody tr", "els => els.length") == 1 and "4987176340900" in pg.inner_text("#groups"))
        pg.fill("#q", "沒這個東西"); pg.wait_for_timeout(300)
        check("找不到 → 寫沒有符合的資料", "沒有符合" in pg.inner_text("#groups") and pg.eval_on_selector_all("#groups .grp", "els => els.length") == 0)
        pg.fill("#q", ""); pg.wait_for_timeout(300)
        check("清掉 → 兩組都回來", pg.eval_on_selector_all("#groups .grp", "els => els.length") == 2 and "符合" not in pg.inner_text("#sum"))
        pg.click("#btn-ph"); pg.wait_for_selector("#ph-box:not(.hidden)")
        pg.once("dialog", lambda dlg: dlg.accept()); pg.click("#ph-rows tr[data-code='TXRC29'] .ph-del"); pg.wait_for_timeout(600)
        check("刪掉 TXRC29 → 表空了、下面那組手機回到預設手機", "TXRC29" not in pg.inner_text("#ph-rows") and "0911222333" not in pg.inner_text("#groups"), pg.inner_text("#groups .grp .gh"))
        pg.click("#ph-close")
        with pg.expect_download() as dl:
            pg.click("#btn-all")
        check("全部下載 → 竹運拋檔_0823交貨.zip", dl.value.suggested_filename == "竹運拋檔_0823交貨.zip", dl.value.suggested_filename)
        with pg.expect_download() as dl:
            pg.click("#groups .grp .dl >> nth=1")
        check("單組下載也是 zip", dl.value.suggested_filename.endswith(".zip"), dl.value.suggested_filename)
        pg.click("#lm-btn"); pg.wait_for_timeout(200)
        check("線別工具寶僑欄有「竹運出貨拋檔」而且標目前在這", pg.eval_on_selector("#lm-pop a.lm-item.on", "e => e.innerText").startswith("竹運出貨拋檔"))
        pg.keyboard.press("Escape")
        pg.click("#hs-open"); pg.wait_for_selector("#hs-table tbody tr", timeout=10000); pg.wait_for_timeout(300)
        check("竹運頁「歷程」視窗：剛才的兩次匯出都在", "匯出竹運拋檔" in pg.inner_text("#hs-table") and "酷澎_PG" in pg.inner_text("#hs-table"), pg.inner_text("#hs-count"))
        pg.click("#hs-close")
        print("\n【6d】YFS 訂單系統入口（/portal）")
        pg.set_viewport_size({"width": 1440, "height": 900})
        pg.goto(f"{base}/"); pg.wait_for_selector("#chs .ch", timeout=15000); pg.wait_for_selector("#c-line svg"); pg.wait_for_selector("#c-bar svg"); pg.wait_for_timeout(300)
        check("入口頁：四個通路卡（含 MOMO）、兩張圖都畫出來", pg.eval_on_selector_all("#chs .ch", "els => els.length") == 4 and pg.eval_on_selector_all("#c-bar svg .hb", "els => els.length") == 7)
        nav = pg.inner_text("#nav")
        check("左側選單：酷澎底下有訂單管理、驗收單簽名、採購表轉換、商品主檔、竹運、瑪氏，還有 SOP 與歷程", all(k in nav for k in ("訂單管理", "驗收單簽名", "採購表轉換", "商品主檔自動化", "竹運出貨拋檔", "瑪氏出貨", "營運 SOP 檢索", "歷程紀錄")))
        check("沒有寫買斷、寄倉", "買斷" not in pg.inner_text("body") and "寄倉" not in pg.inner_text("body"))
        check("左上角是文字「YFS 訂單系統」、選單不畫捲軸", pg.inner_text(".brand").strip() == "YFS 訂單系統" and pg.eval_on_selector(".side", "e => getComputedStyle(e).scrollbarWidth") == "none")
        check("各通路系統卡片：logo 圖都載得出來（52px 看得到吉祥物）、標題是 XX 訂單管理（不寫系統）、酷澎線紅色、MOMO 桃紅、PChome 線藍色",
              pg.eval_on_selector_all("#chs .ch-h img", "els => els.length === 4 && els.every(i => i.complete && i.naturalWidth > 0)")
              and "酷澎訂單管理" in pg.inner_text("#ch-coupang") and "蝦皮特選訂單管理" in pg.inner_text("#ch-shopee") and "PChome 訂單管理" in pg.inner_text("#ch-pchome")
              and pg.eval_on_selector("#ch-coupang", "e => getComputedStyle(e, '::before').backgroundColor") == "rgb(224, 20, 34)"
              and pg.eval_on_selector("#ch-pchome", "e => getComputedStyle(e, '::before').backgroundColor") == "rgb(0, 70, 191)"
              and "MOMO 訂單管理" in pg.inner_text("#ch-momo") and pg.get_attribute("#ch-momo .ch-go", "href") == "/momo"
              and pg.eval_on_selector("#ch-momo", "e => getComputedStyle(e, '::before').backgroundColor") == "rgb(228, 0, 127)"
              and pg.eval_on_selector_all("#chs .ch-h b", "els => els.every(b => b.scrollHeight <= b.clientHeight + 2)")
              and pg.eval_on_selector("#ch-coupang .ch-go", "e => getComputedStyle(e).backgroundColor") == "rgb(224, 20, 34)")
        check("蝦皮特選卡片有「進入後台」連到 /shopee、本月出貨箱數先放 —；PChome 不放數字", pg.get_attribute("#ch-shopee .ch-go", "href") == "/shopee"
              and "—" in pg.inner_text("#ch-shopee") and "規劃中" in pg.inner_text("#ch-pchome") and "—" in pg.inner_text("#ch-pchome"))
        pg.click("#nav-hist"); pg.wait_for_selector("#hs-dlg[open]", timeout=5000)
        check("左側「歷程紀錄」打開歷程視窗", pg.is_visible("#hs-dlg")); pg.click("#hs-close")
        pg.click("#nav a.sub:has-text('驗收單簽名')"); pg.wait_for_selector("#dlg-sign[open]", timeout=15000)
        check("點「驗收單簽名」捷徑 → 到訂單管理並直接打開簽名視窗", pg.is_visible("#dlg-sign") and not pg.evaluate("location.hash"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)   # 先關掉簽名視窗
        pg.click("#btn-yfs"); pg.wait_for_selector("#chs .ch", timeout=15000)
        check("酷澎訂單管理系統按「YFS 訂單系統」回到入口首頁（網站首頁 /）", pg.url.rstrip("/") == base)
        pg.click("#nav a:has-text('改單統計')"); pg.wait_for_selector("#st-kpi .kpi", timeout=15000); pg.wait_for_timeout(500)
        check("左側「改單統計」：在 YFS 訂單系統裡看，四張數字卡（PO 張數、有修改紀錄的 PO、處理次數、處理完成後又有改動；金額卡 2026-10-30 拿掉）、清單、每月表、主因、各階段、處理最多次的 PO 都有",
              pg.url.endswith("/portal/stats") and pg.eval_on_selector_all("#st-kpi .kpi", "els => els.length") == 4 and pg.query_selector("#kpi-pend")
              and "原始下單" not in pg.inner_text("#st-kpi") and "原始下單" in pg.inner_text("#st-months thead") and "處理次數" in pg.inner_text("#st-kpi") and pg.query_selector("#pd-table thead")
              and all(pg.query_selector(sel) for sel in ("#st-months thead", "#st-reasons thead", "#st-stages thead", "#st-top thead")) and "/api/master/stats/export" in pg.get_attribute("#st-export", "href"), pg.inner_text("#st-kpi"))
        check("拿掉的：含測試、待補原因、酷澎修改／我方修改、改單類型；寫了修改紀錄從 2026-10-20 起",
              pg.query_selector("#st-test") is None and "待補" not in pg.inner_text("main") and "我方修改" not in pg.inner_text("#st-months thead") and pg.query_selector("#st-kinds") is None
              and "2026-10-20" in pg.inner_text(".head .sub"))
        pg.fill("#st-from", "2026-09"); pg.dispatch_event("#st-from", "change"); pg.wait_for_timeout(800)
        check("9 月的修改次數是 0、註明 10/20 前沒有修改紀錄；系統抓到的變動次數（灰字）還在、PO 張數和金額照舊有",
              "—" in pg.inner_text("#st-months tbody") and "才開始記" in pg.inner_text("section:has(#st-months) .st-note") and pg.query_selector("#st-months .st-n.st-g")
              and "原始下單" in pg.inner_text("#st-months thead") and "萬" in pg.inner_text("#st-months tbody"), pg.inner_text("#st-months"))
        check("YFS 訂單系統的每月統計：可點的數字不畫底線", pg.eval_on_selector("#st-months .st-n", "e => getComputedStyle(e).textDecorationLine") == "none")
        pg.click("#st-months .st-n.st-g >> nth=0"); pg.wait_for_selector("#st-dlg[open]", timeout=8000)
        check("點灰字 → 跳出系統抓到的變動明細、沒有原因欄", "系統抓到的變動" in pg.inner_text("#st-dlg-title") and "原因" not in pg.inner_text("#st-dlg-table thead")); pg.click("#st-dlg-x")
        # 記一輪 OK → 統計跟著變（酷澎 PO 視窗那段的操作在【6g】測，這裡用 API 記）
        if not tc.get("/api/pos?page_size=5").get_json()["rows"]:
            pv_s = tc.post("/api/import/preview", data={"file": (open(os.path.join(BASE, "samples", "整合表範例.xlsx"), "rb"), "整合表範例.xlsx")}, content_type="multipart/form-data").get_json()
            tc.post("/api/import/commit", json={"operator": "小真", "batch_id": pv_s["batch_id"]})
        rows_s = tc.get("/api/pos?page_size=5").get_json()["rows"]; po_s, po_s2 = rows_s[2]["po_number"], rows_s[3]["po_number"]   # 第一張留給【6g】從 0 次開始記
        d_s = tc.get(f"/api/pos/{po_s2}").get_json(); o_s = d_s["skus"][0]
        tc.put(f"/api/orders/{o_s['id']}", json={"version": o_s["version"], "qty_ship": max(0, (o_s["qty_ship"] or 1) - 1), "operator": "小真"})   # 改過、沒記 OK
        r_s = tc.post(f"/api/pos/{po_s}/rounds", json={"operator": "小真", "stage": "1", "kinds": ["qty", "date"], "reasons": ["庫存不足", "約倉問題"], "note": "UI 測試"})
        assert r_s.status_code == 200, r_s.get_json()
        pg.check("#st-all"); pg.wait_for_timeout(1200)
        check("勾「全部期間」→ 月份欄位不能選、匯出網址帶 all=1；記了一輪 OK 之後修改次數 ≥ 1、曾改單 ≥ 1 張",
              pg.is_disabled("#st-from") and "all=1" in pg.get_attribute("#st-export", "href") and int(pg.inner_text("#st-kpi .kpi >> nth=2").split("\n")[1].replace("次", "").replace(",", "").strip()) >= 1
              and int(pg.inner_text("#st-kpi .kpi >> nth=1").split("\n")[1].replace("張", "").replace(",", "").strip()) >= 1, pg.inner_text("#st-kpi"))
        check("主因表照類別列、只列用到的：數量下修／庫存不足、改交期／倉別／約倉問題", all(w in pg.inner_text("#st-reasons") for w in ("數量下修", "庫存不足", "改交期／倉別", "約倉問題")) and "停產" not in pg.inner_text("#st-reasons"), pg.inner_text("#st-reasons"))
        check("各階段做了什麼：第一階段有數字", "第一階段" in pg.inner_text("#st-stages") and pg.query_selector("#st-stages .st-n"))
        check("修改次數最多的 PO：那張在、PO 單號連到酷澎頁開 PO 視窗、各輪寫階段和做了什麼", po_s in pg.inner_text("#st-top") and f"/coupang?po={po_s}" in pg.get_attribute("#st-top a.link", "href") and "下修" in pg.inner_text("#st-top"))
        pg.click("#st-top .st-n >> nth=0"); pg.wait_for_selector("#st-dlg[open]", timeout=8000)
        check("點修改次數 → 跳出修改紀錄明細（完成日期／階段／做了什麼／主因／說明）", "修改紀錄" in pg.inner_text("#st-dlg-title") and "UI 測試" in pg.inner_text("#st-dlg-table") and "庫存不足" in pg.inner_text("#st-dlg-table")); pg.click("#st-dlg-x")
        check("處理完成後又有改動：卡片是警示、清單有那張 PO、寫出改了什麼、有「填寫修改紀錄」", "warn" in pg.get_attribute("#kpi-pend", "class") and po_s2 in pg.inner_text("#pd-table") and "下修" in pg.inner_text("#pd-table") and pg.query_selector("#pd-table a.go"), pg.inner_text("#pd-table"))
        href_go = pg.get_attribute("#pd-table a.go", "href")
        pg.uncheck("#st-all"); pg.wait_for_timeout(600)
        pg.goto(f"{base}{href_go}"); pg.wait_for_selector("#dlg-po[open]", timeout=15000); pg.wait_for_selector("#po-rounds #rd-ok", timeout=10000)
        check("「填寫修改紀錄」連到 /coupang?po=單號 → 直接打開那張 PO 的視窗、看得到修改紀錄", pg.inner_text("#po-title").strip() == po_s2 and "修改紀錄" in pg.inner_text("#po-rounds") and "po=" not in pg.url)
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        pg.goto(f"{base}/"); pg.wait_for_selector("#cost .hero", timeout=15000)
        check("首頁「本月改單」卡：處理次數、三個階段（收單／排單／出貨 N 次）、主因，沒有估計工時", "處理次數" in pg.inner_text("#cost") and "收單" in pg.inner_text("#cost")
              and "估計工時" not in pg.inner_text("#cost") and "我方修改" not in pg.inner_text("#cost"), pg.inner_text("#cost"))
        pg.goto(f"{base}/master#stats"); pg.wait_for_selector("#tab-stats:not(.hidden)", timeout=15000)
        check("商品主檔的 /master#stats 還是能直接打開改單統計分頁（原本的分頁先留著）", pg.is_visible("#tab-stats"))
        pg.set_viewport_size({"width": 390, "height": 844}); pg.goto(f"{base}/"); pg.wait_for_selector("#c-bar svg"); pg.wait_for_timeout(300)
        check("手機寬度不會左右捲動", not pg.evaluate("document.documentElement.scrollWidth > innerWidth"))
        pg.click("#menu"); pg.wait_for_timeout(350)
        check("手機上按左上角打開選單", pg.evaluate("document.querySelector('.side').getBoundingClientRect().left") >= 0)
        pg.set_viewport_size({"width": 1500, "height": 900})
        print("\n【6e】蝦皮特選（2026-10-13 改版：一個畫面）：一個上傳入口、作業流程格子、全部訂單、差異、改履約方式")
        def shopee_po(rows, path):
            hdr = ["PR ID", "PO ID", "Warehouse", "Expected Delivery Time", "Shopee SKU ID", "Supplier SKU ID", "Shopee SKU Name",
                   "Shopee Requested Qty (unit)", "Selling Type", "EAN / UPC"]
            wbs = openpyxl.Workbook(); wss = wbs.active; wss.append(hdr)
            for sku, sup, name, qty in rows:
                wss.append(["PRTWCOTWA202610019999", None, "TWA", "2026-10-19", sku, sup, name, str(qty), "Pcs", sup.split("_")[0]])
            wbs.save(path)
        def close_side():                                          # 訂單視窗（2026-10-30 改成中間大視窗）會蓋住清單，點清單前先關
            if pg.is_visible("#side"):
                pg.keyboard.press("Escape"); pg.wait_for_selector("#side", state="hidden", timeout=5000)
        def stage(k):
            close_side()
            pg.click(f"#stages .st[data-s='{k}']"); pg.wait_for_selector(f"#stages .st.on[data-s='{k}']", timeout=10000); pg.wait_for_timeout(700)
        def count(k):
            return int(pg.inner_text(f"#stages .st[data-s='{k}'] .n").replace(",", ""))
        sp1 = os.path.join(tempfile.gettempdir(), "ui_PurchaseOrder_1.xlsx"); sp2 = os.path.join(tempfile.gettempdir(), "ui_PurchaseOrder_2.xlsx")
        shopee_po([("11_1", "4987176232878_CNN", "幫寶適 清新幫 拉拉褲", 240), ("11_2", "4987176232878_RNN", "幫寶適 清新幫 箱購", 26)], sp1)
        shopee_po([("11_1", "4987176232878_CNN", "幫寶適 清新幫 拉拉褲", 200), ("11_2", "4987176232878_RNN", "幫寶適 清新幫 箱購", 26)], sp2)
        pg.goto(f"{base}/shopee"); pg.wait_for_selector("#stages .st", timeout=15000)
        check("一個畫面：上傳入口一個、作業流程六格（差異待確認 → … → 全部訂單）、沒有舊頁籤",
              pg.eval_on_selector_all(".dzbox", "els => els.length") == 1 and pg.eval_on_selector_all("#stages .st", "els => els.map(e => e.dataset.s).join()") == "diff,buy,eip,ship,notify,all"
              and pg.query_selector("#tabs") is None)
        pg.wait_for_selector("#stages .st.on", timeout=10000); pg.wait_for_timeout(500)
        check("還沒有任何蝦皮訂單：停在「全部訂單」，清單寫「尚未匯入蝦皮訂單」並說要上傳哪個檔、按哪個鈕",
              pg.get_attribute("#stages .st.on", "data-s") == "all" and "尚未匯入蝦皮訂單" in pg.inner_text("#tbl")
              and "PurchaseOrder" in pg.inner_text("#tbl") and "確認匯入" in pg.inner_text("#tbl"), pg.inner_text("#tbl")[:200])
        pg.wait_for_selector("#fresh span", timeout=8000)
        check("上傳區右邊第一個是「蝦皮採購單未匯入」（黃色），不是只列參考資料", pg.inner_text("#fresh span >> nth=0").strip() == "蝦皮採購單未匯入"
              and "w" in (pg.get_attribute("#fresh span >> nth=0", "class") or ""), pg.inner_text("#fresh"))
        check("上傳區也列瑪氏商品總表（跟瑪氏出貨頁共用，前面 6b 傳過）", "瑪氏商品總表 今天" in pg.inner_text("#fresh"), pg.inner_text("#fresh"))
        pg.fill("#f-from", "2026-01-01"); pg.dispatch_event("#f-from", "change"); pg.wait_for_timeout(500)
        pg.set_input_files("#file", sp1); pg.wait_for_selector(".pv .ok", timeout=15000)
        check("上傳採購單 → 預覽：蝦皮採購單、新增 2 項、沒選履約方式不能按確認", "蝦皮採購單" in pg.inner_text(".pv") and "新增 2 項" in pg.inner_text(".pv").replace("\n", " ")
              and pg.is_disabled(".pv .ok"), pg.inner_text(".pv")[:200])
        pg.reload(); pg.wait_for_selector(".pv .ok", timeout=15000)
        check("還沒按確認匯入就重新整理 → 預覽卡片放回來，寫「還沒按確認匯入」", "還沒按「確認匯入」" in pg.inner_text(".pv")
              and "新增 2 項" in pg.inner_text(".pv").replace("\n", " ") and pg.eval_on_selector_all(".pv", "els => els.length") == 1, pg.inner_text("#previews")[:200])
        pg.fill("#f-from", "2026-01-01"); pg.dispatch_event("#f-from", "change"); pg.wait_for_timeout(500)
        pg.click(".pv .ff[data-f='竹運出貨']"); pg.click(".pv .ok"); pg.wait_for_function("document.querySelectorAll('.pv').length === 0", timeout=10000); pg.wait_for_timeout(800)
        check("匯入後：「待產拋單檔」那格 1 張", count("ship") == 1, pg.inner_text("#stages"))
        check("匯入後上傳區寫「蝦皮採購單 今天 HH:MM」（一天傳好幾次要看得出是哪一次，Jerry 2026-10-28）",
              __import__("re").search(r"蝦皮採購單 今天 \d{2}:\d{2}", pg.inner_text("#fresh")) is not None, pg.inner_text("#fresh"))
        sp3 = os.path.join(tempfile.gettempdir(), "ui_PurchaseOrder_mixed.xlsx")
        wbm = openpyxl.Workbook(); wsm = wbm.active
        wsm.append(["PR ID", "PO ID", "Warehouse", "Expected Delivery Time", "Shopee SKU ID", "Shopee SKU Name", "Supplier SKU ID",
                    "Shopee Requested Qty (unit)", "Selling Type", "EAN / UPC"])
        wsm.append(["PRTWCOTWA202612019999", None, "TWA", "2026-12-01", "MXA", "幫寶適 混線", "4987176232878_CNN", "24", "Pcs", "4987176232878"])
        wsm.append(["PRTWCOTWA202612019999", None, "TWA", "2026-12-01", "MXB", "M&M'S 混線", "M10273409_QMN", "6", "Pcs", "M10273409"])
        wbm.save(sp3)
        pg.set_input_files("#file", sp3); pg.wait_for_selector(".pv .ffl", timeout=15000)
        rows_ = pg.eval_on_selector_all(".pv .ffl", "els => els.map(e => e.dataset.l)")
        check("寶僑、瑪氏混在一個檔：卡片上每個線別一列各選履約方式；瑪氏帶好供應商直送、竹運採購進貨不能選",
              rows_ == ["寶僑", "瑪氏"] and pg.get_attribute(".pv .ffl[data-l='瑪氏'] .ff.on", "data-f") == "供應商直送"
              and "off" in pg.get_attribute(".pv .ffl[data-l='瑪氏'] .ff[data-f='竹運採購進貨']", "class"), str(rows_))
        check("寶僑還沒選：確認匯入按不了、提示寫「請先選擇寶僑的履約方式」", pg.is_disabled(".pv .ok") and "請先選擇寶僑的履約方式" in pg.inner_text(".pv .need"))
        pg.click(".pv .ffl[data-l='寶僑'] .ff[data-f='竹運採購進貨']")
        check("選了寶僑 → 可以按確認匯入", not pg.is_disabled(".pv .ok") and pg.inner_text(".pv .need") == "")
        pg.click(".pv .cancel"); pg.wait_for_function("document.querySelectorAll('.pv').length === 0", timeout=10000)
        stage("all")
        row = pg.inner_text("#tbl tr.og >> nth=0")
        check("全部訂單：這張 PR、寶僑、觀音、竹運出貨，下一步「產出拋單檔」，列上沒有一排按鈕", "PRTWCOTWA202610019999" in row and "寶僑" in row and "觀音" in row and "竹運出貨" in row
              and "產出拋單檔" in row and pg.eval_on_selector_all("#tbl tr.og button", "els => els.length") == 0, row[:200])
        pg.click("#tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_selector("#side:not(.hidden) .sd-it", timeout=8000)
        check("點那一列 → 畫面中間打開訂單視窗（2026-10-30 從右側面板改）：品項一品一列（包／箱）、履約方式、上一張／下一張", "包" in pg.inner_text("#side") and "箱" in pg.inner_text("#side")
              and "履約方式" in pg.inner_text("#sd-title") and pg.eval_on_selector_all("#sd-items tr.sd-it", "els => els.length") == 2 and pg.is_visible("#sd-next")
              and pg.eval_on_selector_all("#tbl tr.sub", "els => els.length") == 0)
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        check("Esc 關掉訂單視窗", pg.is_hidden("#side"))
        pg.set_input_files("#file", sp2); pg.wait_for_selector(".pv .ok", timeout=15000)
        check("重傳改過數量 → 預覽紅字列出差異 240 → 200", "跟上一版有 1 處不同" in pg.inner_text(".pv") and "200" in pg.inner_text(".pv"), pg.inner_text(".pv")[:300])
        pg.click(".pv .ok"); pg.wait_for_function("document.querySelectorAll('.pv').length === 0", timeout=10000); pg.wait_for_timeout(800)
        check("匯入後「差異待確認」那格變 1", count("diff") == 1, pg.inner_text("#stages"))
        pg.click("#tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_function("document.querySelector('#side').innerText.includes('處不同')", timeout=10000)
        check("打開那張：紅字列出差異、有「確認差異」、那品標差異待確認", "240 → 200" in pg.inner_text("#side") and pg.is_visible("#side .sd-dconf")
              and "差異待確認" in pg.inner_text("#sd-items"), pg.inner_text("#side")[:300])
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        stage("diff")
        check("點「差異待確認」：只列這張、列上直接有「確認差異」", pg.eval_on_selector_all("#tbl tr.og", "els => els.length") == 1 and pg.is_visible("#tbl tr.og button.dd"))
        pg.click("#tbl tr.og button.dd"); pg.wait_for_selector("#dd-body table", timeout=8000)
        check("差異視窗：數量 240 → 200、-40、待確認", "240" in pg.inner_text("#dd-body") and "-40" in pg.inner_text("#dd-body") and "待確認" in pg.inner_text("#dd-body"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        # 批次確認差異（Alice 2026-10-21「用批次點選的方式確認，不要一張一張點」；Jerry：差異待確認也一起做）
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        pg.click("#tbl tr.og input.a-ck >> nth=0"); pg.wait_for_selector("#a-bar .bb-confirm", timeout=5000)
        check("勾選那張 → 下面出現批次列：履約方式改成…、「勾選的一起確認差異」；勾選不會跳出訂單視窗（才能繼續勾）", "履約方式改成" in pg.inner_text("#a-bar") and "一起確認差異" in pg.inner_text("#a-bar")
              and pg.is_hidden("#side"), pg.inner_text("#a-bar"))
        pg.click("#a-bar .bb-confirm"); pg.wait_for_timeout(1000)
        check("按「勾選的一起確認差異」→ 那格回到 0、清單沒有了", count("diff") == 0 and "這一步目前沒有要處理的訂單" in pg.inner_text("#tbl"), pg.inner_text("#stages"))
        stage("all"); pg.click("#tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_selector("#side:not(.hidden)", timeout=8000)
        pg.wait_for_function("document.querySelector('#sd-title').innerText.includes('履約方式')", timeout=10000)   # CI 比較慢：等卡片資料讀完再看
        check("訂單視窗的履約方式那格寫「點一下可改」", pg.query_selector("#side .sd-ff") is None and "點一下可改" in pg.inner_text("#sd-title"), pg.inner_text("#sd-title"))
        pg.click("#side .ff-cell"); pg.wait_for_selector("#ff-pop", timeout=5000)
        check("點視窗裡的履約方式 → 跳三個選項、不用原因（選單蓋在視窗上面）", pg.eval_on_selector_all("#ff-pop .ff-opt", "els => els.length") == 3 and "不用填原因" in pg.inner_text("#ff-pop"))
        pg.click("#ff-pop .ff-opt[data-f='供應商直送']")
        pg.wait_for_function("document.querySelector('#sd-title').innerText.includes('供應商直送')", timeout=10000); pg.wait_for_timeout(300)
        row = pg.inner_text("#tbl tr.og >> nth=0")
        check("改履約方式 → 面板寫供應商直送、下一步「產出 EIP 採購單」、待產拋單檔 0、待產採購單 1", "產出 EIP 採購單" in row
              and count("ship") == 0 and count("buy") == 1, row[:200])
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        check("按 Esc 關掉訂單視窗、清單那張也是供應商直送", pg.is_hidden("#side") and "供應商直送" in pg.inner_text("#tbl tr.og >> nth=0"))
        pg.click("#tbl tr.og .ff-cell >> nth=0"); pg.wait_for_selector("#ff-pop", timeout=5000)
        check("清單上的履約方式那格一樣可以點（不會開訂單視窗）", pg.eval_on_selector_all("#ff-pop .ff-opt", "els => els.length") == 3 and pg.is_hidden("#side"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(200)
        # 批次改履約方式、不用原因（Alice 2026-10-21）：改成竹運出貨再改回供應商直送，後面的段落照舊
        pg.click("#tbl tr.og input.a-ck >> nth=0"); pg.wait_for_selector("#a-bar .bb-apply", timeout=5000)
        check("全部訂單勾一張 → 批次列、套用鈕要先選履約方式；訂單視窗不開", pg.is_disabled("#a-bar .bb-apply") and "已勾選 1 張" in pg.inner_text("#a-bar").replace("\n", " ") and pg.is_hidden("#side"), pg.inner_text("#a-bar"))
        pg.click("#a-bar .bb-ff[data-f='竹運出貨']"); pg.click("#a-bar .bb-apply")
        pg.wait_for_function("document.querySelector('#tbl tr.og').innerText.includes('竹運出貨')", timeout=10000); pg.wait_for_timeout(300)
        check("選竹運出貨、套用（沒填原因）→ 那張改成竹運出貨、勾選清掉、待產拋單檔 1", "竹運出貨" in pg.inner_text("#tbl tr.og >> nth=0") and pg.is_hidden("#a-bar") and count("ship") == 1, pg.inner_text("#stages"))
        pg.click("#tbl tr.og input.a-ck >> nth=0"); pg.wait_for_selector("#a-bar .bb-apply", timeout=5000); pg.click("#a-bar .bb-ff[data-f='供應商直送']"); pg.click("#a-bar .bb-apply")
        pg.wait_for_function("document.querySelector('#tbl tr.og').innerText.includes('供應商直送')", timeout=10000); pg.wait_for_timeout(300)
        check("再批次改回供應商直送", "供應商直送" in pg.inner_text("#tbl tr.og >> nth=0") and count("buy") == 1)
        # 匯出訂單明細、刪除／還原訂單（Alice 2026-10-30，Jerry 看過示意圖）
        k0 = pg.inner_text("#tbl tr.og >> nth=0 >> .fn >> nth=0").strip(); n0 = pg.locator("#tbl tr.og").count()
        pg.click("#tbl tr.og input.a-ck >> nth=0"); pg.wait_for_selector("#a-bar .bb-export", timeout=5000)
        with pg.expect_download(timeout=15000) as dle:
            pg.click("#a-bar .bb-export")
        wbe = openpyxl.load_workbook(io.BytesIO(open(dle.value.path(), "rb").read()))
        check("勾一張 → 批次列「匯出訂單明細（1 張）」→ 下載 Excel，一張工作表「訂單明細」、有這張單", "匯出訂單明細（1 張）" in pg.inner_text("#a-bar")
              and dle.value.suggested_filename.startswith("蝦皮特選訂單明細_") and dle.value.suggested_filename.endswith("_1張.xlsx")
              and wbe.sheetnames == ["訂單明細"] and wbe.active.cell(2, 1).value == k0, f"{dle.value.suggested_filename} {wbe.sheetnames}")
        pg.click("#a-bar .bb-clear"); pg.wait_for_timeout(300)
        pg.click("#tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_selector("#side .sd-del", timeout=8000)
        pg.click("#side .sd-del"); pg.wait_for_selector("#dlg-del[open]", timeout=5000)
        check("訂單視窗右上角「刪除這張訂單」→ 跳出刪除視窗，原因寫選填", pg.is_visible("#dlg-del") and k0 in pg.inner_text("#dl-title") and "選填" in pg.inner_text("#dlg-del"))
        pg.click("#dl-chips .dl-r >> nth=0"); pg.click("#dl-go"); pg.wait_for_selector("#side", state="hidden", timeout=8000)
        pg.wait_for_function(f"document.querySelectorAll('#tbl tr.og').length === {n0 - 1}", timeout=8000)
        check("點原因「贈品不用出」→ 刪除：視窗關掉、清單少一張、出現「顯示已刪除 1」", pg.is_visible("#a-del") and "顯示已刪除 1" in pg.inner_text("#a-del")
              and k0 not in pg.inner_text("#tbl"), pg.inner_text("#sum"))
        pg.click("#a-del"); pg.wait_for_selector("#tbl tr.og.del", timeout=8000)
        check("點「顯示已刪除」→ 只列刪掉的那張（灰色），寫誰刪的、原因、有「還原」", "贈品不用出" in pg.inner_text("#tbl tr.og.del") and "Jerry" in pg.inner_text("#tbl tr.og.del")
              and pg.is_visible("#tbl tr.og.del .rs-del") and pg.locator("#tbl tr.og:not(.del)").count() == 0 and "1 張已刪除的訂單" in pg.inner_text("#sum"), pg.inner_text("#tbl")[:300])
        pg.click("#tbl tr.og.del .rs-del"); pg.wait_for_function("!document.querySelector('#tbl tr.og.del')", timeout=8000); pg.wait_for_timeout(300)
        pg.wait_for_function(f"document.querySelectorAll('#tbl tr.og').length === {n0}", timeout=8000)
        check("按還原 → 最後一張還原了就回到一般清單：那張回來、回到待產採購單、沒有刪掉的單就不顯示「顯示已刪除」", k0 in pg.inner_text("#tbl") and count("buy") == 1
              and pg.query_selector("#a-del") is None, pg.inner_text("#stages"))
        # 訂單視窗：上一張／下一張照清單順序；點視窗外面也關
        stage("buy"); pg.wait_for_selector("#b-tbl tr.og", timeout=10000); pg.click("#b-tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_selector("#side:not(.hidden) .sd-sec", timeout=8000)
        check("待產採購單打開一張：寫第 1／1 張、上一張下一張都按不了", "第 1／1 張" in pg.inner_text("#sd-pos") and pg.is_disabled("#sd-prev") and pg.is_disabled("#sd-next"), pg.inner_text("#sd-pos"))
        pg.mouse.click(5, 300); pg.wait_for_timeout(300)
        check("點視窗外面的灰色區域 → 關掉", pg.is_hidden("#side"))
        stage("all"); pg.wait_for_selector("#tbl tr.og", timeout=10000)
        pg.fill("#f-from", "2027-01-01"); pg.dispatch_event("#f-from", "change"); pg.wait_for_selector("#tbl .clr-f", timeout=8000)
        check("到貨日篩掉全部：清單寫「系統共有 1 張，到貨日 10/19～10/19」，有「清除篩選」", "系統共有 1 張" in pg.inner_text("#tbl")
              and "10/19～10/19" in pg.inner_text("#tbl"), pg.inner_text("#tbl")[:200])
        pg.click("#tbl .clr-f"); pg.wait_for_selector("#tbl tr.og", timeout=8000)
        check("按「清除篩選」→ 到貨日清空、那張單回來", pg.input_value("#f-from") == "" and "PRTWCOTWA202610019999" in pg.inner_text("#tbl"))
        pg.fill("#f-from", "2026-01-01"); pg.dispatch_event("#f-from", "change"); pg.wait_for_timeout(500)
        # 上傳日篩選、永豐料號欄（OP 2026-10-21，Jerry 看過示意圖）
        check("篩選列多「上傳日」；清單多「上傳日」「永豐料號」欄，料號照蝦皮採購單的 Supplier SKU ID、上傳日是今天",
              pg.query_selector("#f-up-from") and all(w in pg.inner_text("#tbl thead") for w in ("上傳日", "永豐料號"))
              and "4987176232878_CNN" in pg.inner_text("#tbl tr.og >> nth=0") and datetime.date.today().isoformat() in pg.inner_text("#tbl tr.og >> nth=0"), pg.inner_text("#tbl tr.og >> nth=0")[:200])
        pg.fill("#f-up-from", "2030-01-01"); pg.dispatch_event("#f-up-from", "change"); pg.wait_for_selector("#tbl .clr-f", timeout=8000)
        check("上傳日篩掉全部 → 清單寫上傳日條件下沒有訂單、有清除篩選", "上傳日" in pg.inner_text("#tbl"), pg.inner_text("#tbl")[:200])
        pg.click("#tbl .clr-f"); pg.wait_for_selector("#tbl tr.og", timeout=8000)
        check("清除篩選 → 上傳日也清空", pg.input_value("#f-up-from") == "" and "PRTWCOTWA202610019999" in pg.inner_text("#tbl"))
        pg.fill("#f-from", "2026-01-01"); pg.dispatch_event("#f-from", "change"); pg.wait_for_timeout(500)
        pg.click("#btn-set"); pg.wait_for_selector("#view-set:not(.hidden)", timeout=5000); pg.click("#up-box summary"); pg.wait_for_timeout(500)
        check("資料與設定：上傳紀錄看得到兩次上傳、可下載原始檔", pg.eval_on_selector_all("#up-tbl a.lnk", "els => els.length") >= 2)
        check("資料與設定：庫存資料三種都還沒上傳、效期標準；沒有「出貨方式」可選（一律混合效期，2026-10-20）", pg.inner_text("#stock-st").count("尚未上傳") == 3
              and pg.input_value("#s-rule") == "shopee" and pg.query_selector("#s-mode") is None, pg.inner_text("#stock-st"))
        pg.select_option("#s-rule", "custom"); pg.wait_for_timeout(600)
        check("選自訂比例 → 出現比例欄、帶 0.66", pg.is_visible("#s-ratio") and pg.input_value("#s-ratio") == "0.66")
        pg.select_option("#s-rule", "shopee"); pg.wait_for_timeout(600)
        check("改回蝦皮進倉規範 → 比例欄收起來", pg.is_hidden("#s-ratio"))
        pg.click("#set-back"); pg.wait_for_selector("#view-work:not(.hidden)", timeout=5000)

        print("\n【6e-2】待產採購單：上傳入口傳價格本、除不盡改數量、產出 EIP 採購單；待回填 EIP 採購單號")
        spp = os.path.join(tempfile.gettempdir(), "ui_PG價格本整合.xlsx")
        wbp = openpyxl.Workbook(); wsp = wbp.active; wsp.title = "價格本整合"; wsp.append(["料號", "單位", "品名", "料號3", "箱入數"])
        wsp.append(["4987176232878", "包", "幫寶適 清新幫 拉拉褲", "4987176232878C", 3]); wsp.append(["4987176232878", "箱", "幫寶適 清新幫 箱購", "4987176232878R", 1])
        wbp.save(spp)
        stage("buy"); pg.wait_for_selector("#b-tbl tr.og", timeout=10000)
        check("待產採購單：價格本還沒傳，上傳入口旁標「寶僑價格本未上傳」、這張擋下來", "寶僑價格本未上傳" in pg.inner_text("#fresh")
              and "需處理" in pg.inner_text("#b-tbl tr.og >> nth=0") and pg.is_disabled("#b-go"), pg.inner_text("#b-tbl")[:200])
        check("擋下來的那張寫出原因「價格本找不到 2 項」、勾選框旁邊紅色驚嘆號（勾了只能批次改履約方式，不能產出），下面提示點該列看原因",
              "價格本找不到 2 項" in pg.inner_text("#b-tbl tr.og >> nth=0") and pg.is_visible("#b-tbl tr.og .blk-i")
              and "有 1 張需先處理" in pg.inner_text("#b-go-hint"), pg.inner_text("#b-tbl tr.og >> nth=0") + " / " + pg.inner_text("#b-go-hint"))
        # 表頭篩選（Jerry 2026-10-19「可以做個篩選嗎」，看示意圖選跟 Excel 一樣的）
        check("待產採購單的表頭：勾選框以外每一欄都可以篩（Jerry 2026-10-20「全部都要」）", pg.eval_on_selector_all("#b-tbl thead .hf-b", "els => els.map(e => e.dataset.hf).join()") == "no,line,date,up,wh,ff,codes,items,qty,chk")
        pg.click("#b-tbl thead .hf-b[data-hf='chk']"); pg.wait_for_selector("#hf-pop", timeout=5000)
        hpop = pg.inner_text("#hf-pop")
        check("檢查結果點開：需先處理底下列原因和張數（價格本找不到 1 張）", "需先處理" in hpop and "價格本找不到" in hpop and "1 張" in hpop, hpop)
        pg.uncheck("#hf-pop input.hf-g")
        check("全部取消勾選 → 套用不能按", pg.is_disabled("#hf-pop .hf-ok"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(200)
        check("Esc 關掉篩選視窗", pg.query_selector("#hf-pop") is None)
        pg.click("#b-sum .hf-q"); pg.wait_for_timeout(300)
        check("點上面「1 張需先處理」→ 檢查結果表頭變紅、寫篩選後 1 張", pg.query_selector("#b-tbl thead th.hf-on") is not None and "篩選後 1 張" in pg.inner_text("#b-sum"), pg.inner_text("#b-sum"))
        pg.click("#b-sum .hf-q"); pg.wait_for_timeout(300)
        check("再點一次 → 取消篩選", pg.query_selector("#b-tbl thead th.hf-on") is None and "篩選後" not in pg.inner_text("#b-sum"), pg.inner_text("#b-sum"))
        pg.click("#b-tbl tr.og .blk-i"); pg.wait_for_selector("#side:not(.hidden) .prob", timeout=8000)
        check("點紅色驚嘆號 → 訂單視窗打開，需處理的品項標紅排最前面、原因寫在品名下面，標題寫「2 項需處理，排在最前面」", "價格本找不到" in pg.inner_text("#side")
              and "2 項需處理，排在最前面" in pg.inner_text("#side .sd-sec .h >> nth=0") and pg.eval_on_selector_all("#sd-items tr.sd-it.bad", "els => els.length") == 2, pg.inner_text("#side")[:300])
        pg.click("#side .sd-fl[data-f='bad']"); pg.wait_for_timeout(200)
        check("視窗裡按「需處理 2」→ 只剩那 2 品", pg.eval_on_selector_all("#sd-items tr.sd-it", "els => els.length") == 2 and "顯示 2 品" in pg.inner_text("#sd-found"))
        pg.click("#side .sd-fl[data-f='all']"); pg.wait_for_timeout(200)
        pg.set_input_files("#file", spp); pg.wait_for_function("document.querySelector('#fresh').innerText.includes('寶僑價格本 ')", timeout=10000); pg.wait_for_timeout(800)
        check("同一個上傳入口丟價格本（視窗開著也能丟）→ 自己認出來、提示已更新", "價格本已更新" in pg.inner_text("#toasts"), pg.inner_text("#toasts"))
        pg.wait_for_selector("#side .sd-adj", timeout=8000)
        check("開著的視窗跟著更新：200 包除以轉換率 3 除不盡的原因和「修改數量」", "除不盡" in pg.inner_text("#side") and pg.is_visible("#side .sd-adj"), pg.inner_text("#side")[:300])
        pg.click("#side .sd-adj"); pg.wait_for_selector("#dlg-adj[open]", timeout=5000)
        check("修改數量視窗：原始數量、轉換率、計算結果、餘數，快速填入 198／201", all(w in pg.inner_text("#da-facts") for w in ("200", "轉換率", "66.67", "餘數"))
              and "198" in pg.inner_text("#da-down") and "201" in pg.inner_text("#da-up"), pg.inner_text("#da-facts"))
        pg.click("#da-down"); pg.click("#da-save"); pg.wait_for_timeout(500)
        check("沒寫原因不能存（視窗還開著）", pg.is_visible("#dlg-adj"))
        pg.fill("#da-reason", "與業務確認改為整箱"); pg.click("#da-save"); pg.wait_for_selector("#dlg-adj[open]", state="detached", timeout=5000)
        pg.wait_for_function("document.querySelector('#side').innerText.includes('已修改')", timeout=10000); pg.wait_for_timeout(300)
        check("存好後：清單變可產出、視窗裡下採 198 已修改、EIP 數量 66 箱、出現「產出這張單的 EIP 採購單」", "可產出" in pg.inner_text("#b-tbl tr.og >> nth=0") and "198" in pg.inner_text("#side")
              and "66 箱" in pg.inner_text("#sd-items") and pg.is_visible("#side .sd-regen"), pg.inner_text("#side")[:400])
        close_side()
        pg.click("#b-all")
        with pg.expect_download(timeout=15000) as dl:
            pg.click("#b-go")
        check("勾選後按產出 → 下載 zip（直送採購單＋到貨日）", dl.value.suggested_filename == "蝦皮特選直送採購單_1019.zip", dl.value.suggested_filename)
        pg.wait_for_timeout(1000)
        check("產出後：待產採購單 0、待回填 EIP 採購單號 2（箱單位、小單位兩個檔）", count("buy") == 0 and count("eip") == 2, pg.inner_text("#stages"))
        stage("eip"); pg.wait_for_selector("#e-tbl input.e-po", timeout=10000)
        names = pg.inner_text("#e-tbl")
        check("待回填 EIP 採購單號：兩個檔都在、可重新下載", pg.eval_on_selector_all("#e-tbl input.e-po", "els => els.length") == 2
              and "寶僑直送_1019到觀音-小單位.xls" in names and "寶僑直送_1019到觀音-箱單位.xls" in names, names[:300])
        pg.fill("#e-tbl input.e-po >> nth=0", "12345"); pg.click("#e-tbl button.e-save >> nth=0"); pg.wait_for_timeout(500)
        check("單號格式不對 → 紅色提示", "PO 加 9 碼數字" in pg.inner_text("#toasts"), pg.inner_text("#toasts"))
        pg.fill("#e-tbl input.e-po >> nth=0", "PO202610003"); pg.press("#e-tbl input.e-po >> nth=0", "Enter"); pg.wait_for_timeout(1000)
        check("回填 PO202610003 → 這份從待回填拿掉、格子變 1", "已回填 EIP 採購單號 PO202610003" in pg.inner_text("#toasts") and count("eip") == 1
              and pg.eval_on_selector_all("#e-tbl input.e-po", "els => els.length") == 1, pg.inner_text("#stages"))
        pg.click("#e-all"); pg.wait_for_timeout(300)
        check("按「顯示已回填」→ 看得到剛回填的那份、記下回填人", "PO202610003" in [pg.input_value(x) for x in ["#e-tbl input.e-po >> nth=1"]][0] and "Jerry" in pg.inner_text("#e-tbl tr >> nth=2"))
        check("沒有瑪氏 EIP 時不顯示「下載瑪氏採購單」", pg.is_hidden("#e-mars"))
        with pg.expect_download(timeout=15000) as dl:
            pg.click("#e-tbl a[href$='/direct_po']")
        check("寶僑直送回填後：那列多「直送採購單」、上方出現「下載直送採購單（已回填）」；檔名帶單號、到貨日、倉、單位（Alice 2026-10-21）",
              pg.is_visible("#e-direct") and dl.value.suggested_filename.startswith("PO202610003蝦皮特選直送採購單_1019_觀音_"), dl.value.suggested_filename)
        check("每份 EIP 那列有「取消產出」；已回填的那份灰掉", pg.eval_on_selector_all("#e-tbl button.e-cancel", "els => els.length") == 2
              and pg.eval_on_selector_all("#e-tbl button.e-cancel:disabled", "els => els.length") == 1, pg.inner_text("#e-tbl")[:200])
        # 勾選 → 一次下載勾的那幾份（Jerry 2026-10-30「就跟其他部份一樣」）
        pg.check("#e-ck-all"); pg.wait_for_selector("#e-bar .eb-eip", timeout=5000)
        check("表頭全選 → 下面一條：已勾選 2 份、下載勾選的 EIP 檔（2 份）、直送採購單只算已回填的（1 份）", "已勾選 2 份" in pg.inner_text("#e-bar")
              and "EIP 檔（2 份）" in pg.inner_text("#e-bar") and "直送採購單（1 份）" in pg.inner_text("#e-bar") and "還沒回填" in pg.inner_text("#e-bar"), pg.inner_text("#e-bar"))
        with pg.expect_download(timeout=15000) as dlz:
            pg.click("#e-bar .eb-eip")
        zf_ = zipfile.ZipFile(io.BytesIO(open(dlz.value.path(), "rb").read()))
        check("下載勾選的 EIP 檔 → 一個壓縮檔、裡面 2 份（回填過的檔名帶單號）", len(zf_.namelist()) == 2 and any("PO202610003" in n for n in zf_.namelist()), str(zf_.namelist()))
        with pg.expect_download(timeout=15000) as dlz:
            pg.click("#e-bar .eb-direct")
        check("下載勾選的直送採購單 → 只有回填過的那 1 份", len(zipfile.ZipFile(io.BytesIO(open(dlz.value.path(), "rb").read())).namelist()) == 1)
        pg.click("#e-bar .eb-clear"); pg.wait_for_timeout(300)
        check("取消勾選 → 那條收起來", pg.is_hidden("#e-bar") and not pg.is_checked("#e-tbl input.e-ck >> nth=0"))
        pg.once("dialog", lambda d: d.accept()); pg.click("#e-tbl button.e-cancel:not(:disabled)"); pg.wait_for_timeout(1200)
        check("沒回填的那份按「取消產出」→ 刪掉、提示回到待產採購單、待產採購單變 1", "回到「待產採購單」" in pg.inner_text("#toasts") and count("buy") == 1
              and pg.eval_on_selector_all("#e-tbl button.e-cancel", "els => els.length") == 1, pg.inner_text("#toasts") + pg.inner_text("#stages"))
        # 改履約方式前的提醒（Jerry 2026-10-22）：已回填 EIP 的單改成竹運出貨 → 先跳視窗講會怎樣、要做什麼；按「先不要」不改
        no_f = pg.evaluate("async () => (await (await fetch('/api/shopee/eips?from=2026-01-01')).json()).eips.find(e => e.eip_po).orders[0]")
        stage("all"); pg.wait_for_selector(f"#tbl .ff-cell[data-k='{no_f}']", timeout=10000)
        pg.click(f"#tbl .ff-cell[data-k='{no_f}']"); pg.wait_for_selector("#ff-pop", timeout=5000)
        pg.click("#ff-pop .ff-opt[data-f='竹運出貨']"); pg.wait_for_selector("#dlg-ff[open]", timeout=8000)
        txt = pg.inner_text("#dlg-ff")
        check("已回填 EIP 的單改成竹運出貨 → 先跳提醒：原廠已經下單、系統會自動做什麼、你要做的事（通知採購）",
              "改成「竹運出貨」" in txt and "原廠那邊已經下單了" in txt and "系統會自動" in txt and "你要做的事" in txt and "通知採購" in txt and "待產拋單檔" in txt, txt[:400])
        pg.click("#df-no"); pg.wait_for_timeout(600)
        check("按「先不要」→ 視窗關掉、履約方式沒改", not pg.evaluate("document.getElementById('dlg-ff').open")
              and "竹運出貨" not in pg.inner_text(f"#tbl .ff-cell[data-k='{no_f}']"), pg.inner_text(f"#tbl .ff-cell[data-k='{no_f}']"))
        pg.click("#btn-set"); pg.wait_for_selector("#view-set:not(.hidden)", timeout=5000)
        check("資料與設定：價格本列出寶僑、瑪氏（商品總表在瑪氏出貨上傳）", "寶僑" in pg.inner_text("#conv-st") and "瑪氏出貨" in pg.inner_text("#conv-st"), pg.inner_text("#conv-st"))
        pg.click("#wh-box summary"); pg.wait_for_selector("#wh-tbl tbody tr", timeout=10000)
        check("蝦皮倉資料：9 個倉、觀音帶出需求書的入倉倉別和 ship-to、進倉說明（直送採購單用）", pg.eval_on_selector_all("#wh-tbl tbody tr", "els => els.length") == 9
              and pg.input_value("#wh-tbl tr[data-code='TWA'] input[data-k='name']") == "永豐商店蝦皮特選-觀音倉(TWA)"
              and pg.input_value("#wh-tbl tr[data-code='TWA'] input[data-k='ship_to']") == "17617040"
              and "寶倉街" in pg.input_value("#wh-tbl tr[data-code='TWA'] textarea[data-k='notes']"), pg.inner_text("#wh-tbl")[:200])
        pg.fill("#wh-tbl tr[data-code='TWG'] input[data-k='ship_to']", "17600001"); pg.click("#wh-tbl tr[data-code='TWG'] button.wh-save"); pg.wait_for_timeout(500)
        check("改高鐵南 ship-to 按儲存 → 提示已更新", "已更新" in pg.inner_text("#toasts"), pg.inner_text("#toasts"))
        pg.click("#set-back"); pg.wait_for_selector("#view-work:not(.hidden)", timeout=5000); pg.wait_for_timeout(800)
        stage("ship")
        check("待產拋單檔：沒有竹運出貨的單時寫提示", "這一步目前沒有要處理的訂單" in pg.inner_text("#z-tbl"), pg.inner_text("#z-tbl")[:100])

        print("\n【6e-4】寶僑缺貨（全部訂單點開那張，出貨量直接改）→ 待通知蝦皮 → 產 LINE 通知")
        stage("all"); pg.click("#tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_selector("#side .sd-qe", timeout=8000)
        check("訂單視窗：寶僑的出貨量可以點、寫「點一下可改」；「填缺貨」按鈕改成「缺貨紀錄」（Jerry 2026-10-30「像酷澎那樣」）",
              "點一下可改" in pg.inner_text("#sd-items") and "缺貨紀錄" in pg.inner_text("#sd-items") and pg.query_selector("#sd-items button.sd-short") is None, pg.inner_text("#sd-items")[:300])
        pg.click("#side .sd-short >> nth=0"); pg.wait_for_selector("#dlg-short[open]", timeout=5000)
        check("「缺貨紀錄」一樣打開原本的視窗：第 1 次下修、預設勾庫存不足", "第 1 次下修" in pg.inner_text("#dk-round") and pg.is_checked(".dk-r[value='庫存不足']"))
        pg.click("#dlg-short .x"); pg.wait_for_timeout(300)
        q0 = int(pg.inner_text("#side .sd-qe >> nth=0").replace(",", ""))
        pg.click("#side .sd-qe >> nth=0"); pg.wait_for_selector("#side .qe-in", timeout=5000)
        check("點出貨量 → 變輸入框，下面缺貨原因預設勾庫存不足、有儲存／取消", pg.is_visible("#side .qe-save") and "庫存不足" in pg.inner_text("#side .qe-r.on"))
        pg.fill("#side .qe-in", str(q0 - 6)); pg.dispatch_event("#side .qe-in", "input")
        check("打新的數字 → 寫這次缺幾個", "缺 6" in pg.inner_text("#side .qe-cut"), pg.inner_text("#side .qe-cut"))
        pg.press("#side .qe-in", "Enter"); pg.wait_for_function("[...document.querySelectorAll('.toast')].some(t => t.innerText.includes('已列入「待通知蝦皮」'))", timeout=8000)
        pg.wait_for_timeout(1000)
        check("按 Enter 存好：視窗還開著、出貨量變新的、寫（已修改）", pg.is_visible("#side") and str(q0 - 6) in pg.inner_text("#side .sd-qe >> nth=0").replace(",", "")
              and "已修改" in pg.inner_text("#sd-items"), pg.inner_text("#sd-items")[:300])
        check("存好後：待通知蝦皮那格變 1、這張下一步多「通知蝦皮」", count("notify") == 1 and "通知蝦皮" in pg.inner_text("#tbl tr.og >> nth=0"), pg.inner_text("#stages"))
        stage("notify"); pg.wait_for_selector("#k-tbl tr.og", timeout=8000)
        check("待通知蝦皮：只列這張、待通知", pg.eval_on_selector_all("#k-tbl tr.og", "els => els.length") == 1 and "待通知" in pg.inner_text("#k-tbl tr.og >> nth=0"))
        pg.once("dialog", lambda d: d.accept())
        pg.click("#k-tbl input.k-ck >> nth=0"); pg.click("#k-go"); pg.wait_for_selector("#dlg-notify[open] #dn-msgs textarea", timeout=10000)
        txt = pg.input_value("#dn-msgs textarea.dn-text >> nth=0")
        check("產生通知 → 跳出 LINE 訊息（照 Alice 格式：下修通知、訂購／可出貨／缺貨、處理方式）、可下載明細", txt.startswith("【蝦皮缺貨下修通知｜首次通知】")
              and "／缺貨6" in txt and "處理方式：下修至" in txt and "/xlsx" in (pg.get_attribute("#dn-xlsx", "href") or ""), txt[:300])
        pg.click("#dn-msgs button.dn-copy >> nth=0"); pg.wait_for_timeout(300)
        check("按複製 → 提示已複製", "已複製" in pg.inner_text("#toasts"))
        pg.click("#dn-tab-mail"); pg.wait_for_timeout(200)
        check("通知視窗分 LINE、Email 兩個分頁：Email 那頁有主旨、內容、打開 Outlook 寄出、只複製內容（Jerry 2026-10-22 雛形）", pg.is_visible("#dn-mail") and pg.is_visible("#dn-mail-copy")
              and pg.is_hidden("#dn-line") and "蝦皮特選缺貨通知" in pg.inner_text("#dn-subj") and txt in pg.input_value("#dn-body"), pg.inner_text("#dn-subj"))
        pg.click("#dlg-notify .x"); pg.wait_for_timeout(800)
        check("關掉通知視窗 → 提示這次通知在下面的通知紀錄", "下面的通知紀錄" in pg.inner_text("#toasts"), pg.inner_text("#toasts"))
        check("通知後：待通知蝦皮那格回到 0", count("notify") == 0, pg.inner_text("#stages"))
        stage("notify"); pg.wait_for_selector("#k-hist button.k-n", timeout=8000)
        check("通知紀錄就在待通知蝦皮這一頁下面：剛剛那次一列，有 LINE 複製、Email 寄、明細 Excel，內容寫下修幾項幾則", pg.is_visible("#k-hist-box")
              and pg.eval_on_selector_all("#k-hist button.k-m", "els => els.length") >= 1 and pg.eval_on_selector_all("#k-hist a[href$='/xlsx']", "els => els.length") >= 1
              and "下修" in pg.inner_text("#k-hist tbody tr >> nth=0") and "則" in pg.inner_text("#k-hist tbody tr >> nth=0") and pg.is_visible("#nm-to"), pg.inner_text("#k-hist")[:200])
        pg.click("#k-hist button.k-n >> nth=0"); pg.wait_for_selector("#dlg-notify[open] #dn-msgs textarea", timeout=8000)
        check("按「LINE 複製」→ 打開那次的通知、停在 LINE 分頁", pg.input_value("#dn-msgs textarea.dn-text >> nth=0") == txt and pg.is_visible("#dn-line"))
        pg.click("#dlg-notify .x"); pg.wait_for_timeout(500)

        print("\n【6e-5】瑪氏缺貨：勇信配送明細表從同一個上傳入口進來 → 比對視窗 → 寫入缺貨 → 待通知蝦皮")
        # 瑪氏直送的單、EIP、回填單號用 API 先備好（畫面流程前面測過了），這裡只測勇信核對這一段
        cn = db.get_conn()
        cn.execute("INSERT INTO mst_mars_products (yf_sku, unit, mars_code, category, name, price, inner_per_case, pcs_per_case, unit_code, box_qty, "
                   "inner_label, note, po_case_note, shelf_days) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                   ("M69990001", "盒", "69990001", "Petcare", "UI 測試 瑪氏貓糧", 100.0, 12, 12, "Q", 12, "", "", "", 365))
        cn.commit(); cn.close()
        hdr = ["PR ID", "PO ID", "Warehouse", "Expected Delivery Time", "Shopee SKU ID", "Supplier SKU ID", "Shopee SKU Name",
               "Shopee Requested Qty (unit)", "Selling Type", "EAN / UPC"]
        wbm = openpyxl.Workbook(); wsm = wbm.active; wsm.append(hdr)
        wsm.append(["PRTWCOTWA202610209999", "POTWCTWA26102099M", "TWA", "2026-10-21", "99_1", "M69990001_QMN", "UI 測試 瑪氏貓糧", "24", "Pcs", "M69990001"])
        bm = io.BytesIO(); wbm.save(bm)
        up = tc.post("/api/shopee/import/preview", data={"file": (io.BytesIO(bm.getvalue()), "PurchaseOrder_ui_mars.xlsx")}, content_type="multipart/form-data").get_json()
        tc.post("/api/shopee/import/commit", json={"upload_id": up["upload_id"], "fulfil": "供應商直送"})
        gen = tc.post("/api/shopee/purchase/generate", json={"fulfil": "供應商直送", "order_nos": ["POTWCTWA26102099M"]})
        me = next(e for e in tc.get("/api/shopee/eips").get_json()["eips"] if e["line"] == "瑪氏")
        tc.put(f"/api/shopee/eips/{me['id']}", json={"eip_po": "PO202610299"})
        doc = pymupdf.open(); page = doc.new_page(); y = 40
        for ln in ["BMS100230-永豐商店-蝦皮特選", "指送日期：2026/10/21", "配 送 明 細 表(預排鮮度)", "收貨單號：PO202610299-蝦皮-T",
                   "5888069990001", "1", "UI 測試 瑪氏貓糧 1:6:10", "1C  ", "625D1TAP01", "2027/06/01", "合計箱數:1"]:
            page.insert_text((40, y), ln, fontname="china-t", fontsize=9); y += 14
        yxs = os.path.join(tempfile.gettempdir(), "ui_蝦皮_勇信.pdf"); doc.save(yxs)
        pg.set_input_files("#file", yxs); pg.wait_for_selector("#dlg-yx[open] #yx-out table.t", timeout=20000); pg.wait_for_timeout(300)
        out = pg.inner_text("#yx-out")
        check("PDF 丟上傳入口 → 跳出比對視窗：對到 PO202610299、訂 2 箱出 1 箱、部分缺貨、原因庫存不足、24 → 12、確認鈕可按", gen.status_code == 200 and "PO202610299" in out
              and "部分缺貨" in out and "下修至 12" in out and "原因：庫存不足" in out and "24 → 12" in out.replace("\t", " ") and not pg.is_disabled("#yx-go"), out[:400])
        check("效期欄：每批效期、箱數，下面寫允收哪天以後（到貨日 10/21＋365×2/5）", "2027-06-01 × 1" in out and "允收 2027-03-16 以後" in out, out[:600])
        pg.uncheck("#yx-out input.yx-ck"); pg.wait_for_function("document.querySelector('#yx-out').innerText.includes('這次不核')", timeout=10000)
        check("取消勾選 → 這次不核、確認鈕鎖住", pg.is_disabled("#yx-go"))
        pg.check("#yx-out input.yx-ck"); pg.wait_for_function("!document.querySelector('#yx-go').disabled", timeout=10000)
        pg.click("#yx-go"); pg.wait_for_function("document.querySelector('#toasts').innerText.includes('已寫入 1 筆缺貨')", timeout=10000)
        pg.wait_for_function("document.querySelector('#yx-out').innerText.includes('前次寫入')", timeout=10000)
        check("比對結果重算：那筆顯示前次寫入", "前次寫入" in pg.inner_text("#yx-out"))
        pg.click("#yx-clear"); pg.wait_for_selector("#dlg-yx[open]", state="detached", timeout=5000); pg.wait_for_timeout(800)
        check("關掉視窗：待通知蝦皮那格變 1", count("notify") == 1, pg.inner_text("#stages"))
        stage("notify"); pg.wait_for_selector("#k-tbl tr.og[data-k='POTWCTWA26102099M']", timeout=10000)
        row = pg.inner_text("#k-tbl tr.og[data-k='POTWCTWA26102099M']")
        check("待通知蝦皮：這張瑪氏單、待通知、累計下修 12", "瑪氏" in row and "待通知" in row and "12" in row, row)
        check("待通知蝦皮的表頭：勾選框以外每一欄都可以篩", pg.eval_on_selector_all("#k-tbl thead .hf-b", "els => els.map(e => e.dataset.hf).join()") == "no,line,ff,date,up,wh,codes,items,orig,final,cut,st")
        stage("all"); pg.wait_for_selector("#tbl tr.og", timeout=10000)
        n_all = pg.eval_on_selector_all("#tbl tr.og", "els => els.length")
        check("全部訂單：勾選框以外看得到的每一欄都有 ▾", pg.eval_on_selector_all("#tbl thead th", "els => els.filter(e => e.offsetParent && !e.querySelector('input')).every(e => e.querySelector('.hf-b'))"))
        pg.click("#tbl thead .hf-b[data-hf='qty']"); pg.wait_for_selector("#hf-pop .hf-sort", timeout=5000)
        pg.click("#hf-pop .hf-sort[data-d='-1']"); pg.wait_for_timeout(300)
        hq = [int(x.replace(",", "")) for x in pg.eval_on_selector_all("#tbl tr.og td:nth-child(10) b", "els => els.map(e => e.innerText)")]   # 多了勾選框、上傳日、永豐料號三欄（2026-10-21）
        check("數字欄可以排序：蝦皮數量由大到小、表頭有排序記號、數字列寫怎麼排的", len(hq) == n_all and hq == sorted(hq, reverse=True)
              and pg.query_selector("#tbl thead .hf-sorted") is not None and "依蝦皮數量由大到小排" in pg.inner_text("#sum"), f"{hq} / {pg.inner_text('#sum')}")
        pg.click("#sum .hf-clr"); pg.wait_for_timeout(300)
        pg.click("#tbl thead .hf-b[data-hf='next']"); pg.wait_for_selector("#hf-pop", timeout=5000)
        hopts = pg.eval_on_selector_all("#hf-pop input.hf-v", "els => els.map(e => e.value)")
        check("全部訂單：下一步的表頭點開，列出各種下一步和張數", "通知蝦皮" in hopts and len(hopts) >= 2 and "張" in pg.inner_text("#hf-pop"), str(hopts))
        for v in hopts:
            if v != "通知蝦皮":
                pg.uncheck(f"#hf-pop input.hf-v[value='{v}']")
        pg.click("#hf-pop .hf-ok"); pg.wait_for_timeout(300)
        hkeys = pg.eval_on_selector_all("#tbl tr.og", "els => els.map(e => e.dataset.k)")
        check("只勾「通知蝦皮」→ 只剩要通知的單、表頭變紅、數字列寫篩了哪一欄、篩選後幾張", 0 < len(hkeys) < n_all and "POTWCTWA26102099M" in hkeys
              and pg.query_selector("#tbl thead th.hf-on") is not None and f"篩選後 {len(hkeys)} 張" in pg.inner_text("#sum") and "下一步（通知蝦皮）" in pg.inner_text("#sum"),
              f"{hkeys} / {n_all} / {pg.inner_text('#sum')}")
        pg.click("#sum .hf-clr"); pg.wait_for_timeout(300)
        check("清除表頭篩選 → 全部回來", pg.eval_on_selector_all("#tbl tr.og", "els => els.length") == n_all and pg.query_selector("#tbl thead th.hf-on") is None)
        pg.click("#tbl thead .hf-b[data-hf='next']"); pg.wait_for_selector("#hf-pop", timeout=5000)
        for v in hopts:
            if v != "通知蝦皮":
                pg.uncheck(f"#hf-pop input.hf-v[value='{v}']")
        pg.click("#hf-pop .hf-ok"); pg.wait_for_timeout(300)
        stage("notify"); pg.wait_for_selector("#k-tbl tr.og[data-k='POTWCTWA26102099M']", timeout=10000); stage("all"); pg.wait_for_selector("#tbl tr.og", timeout=10000)
        check("換一步再回來，表頭篩選清掉（只管目前這一步）", pg.eval_on_selector_all("#tbl tr.og", "els => els.length") == n_all and pg.query_selector("#tbl thead th.hf-on") is None)
        # 訂單視窗裡的品項搜尋（2026-10-30）：打料號只留那幾品並標起來；上面清單的搜尋框有字，打開時帶進來
        pg.fill("#f-q", "4987176232878_CNN"); pg.press("#f-q", "Enter"); pg.wait_for_timeout(1500)
        pg.click("#tbl tr.og >> nth=0 >> .fn >> nth=0"); pg.wait_for_selector("#side:not(.hidden) #sd-items tr.sd-it", timeout=10000)
        hit_all = pg.eval_on_selector_all("#sd-items tr.sd-it", "els => els.length > 0 && els.every(e => e.classList.contains('hit') && e.innerText.includes('4987176232878_CNN'))")
        check("清單搜料號 4987176232878_CNN 再打開一張 → 視窗的品項搜尋帶進同一個字，只剩符合的品項、標起來、寫找到幾品",
              pg.input_value("#sd-q") == "4987176232878_CNN" and hit_all and "找到" in pg.inner_text("#sd-found"), pg.inner_text("#sd-found"))
        pg.fill("#sd-q", ""); pg.wait_for_timeout(300)
        check("清掉視窗裡的搜尋 → 品項全部回來、不再標", pg.query_selector("#side .sd-it.hit") is None and pg.eval_on_selector_all("#sd-items tr.sd-it", "els => els.length") > 1)
        close_side(); pg.fill("#f-q", ""); pg.press("#f-q", "Enter"); pg.wait_for_timeout(1200)
        stage("notify"); pg.wait_for_selector("#k-tbl tr.og[data-k='POTWCTWA26102099M']", timeout=10000)
        pg.click("#k-tbl tr.og[data-k='POTWCTWA26102099M'] .fn")
        pg.wait_for_function("document.querySelector('#sd-title').innerText.includes('POTWCTWA26102099M') && document.querySelectorAll('#side .sd-it').length > 0", timeout=10000)
        side = pg.inner_text("#side")
        check("訂單視窗：下修紀錄來源勇信核對、沒有填缺貨按鈕（寫依勇信配送明細表核對）、可直接產這張單的通知", "勇信核對" in side and "瑪氏的缺貨依勇信配送明細表核對" in side and side.count("勇信配送明細表核對") == 1
              and pg.query_selector("#side .sd-short") is None and pg.is_visible("#side .sd-notify"), side[:400])
        check("有缺貨的單：視窗裡有「蝦皮採購單（畫缺貨）」，還沒上傳就寫怎麼做", "蝦皮採購單（畫缺貨）" in side and "尚未上傳" in side and "後台下載" in side)
        # 蝦皮採購單 PDF 畫缺貨（Alice 2026-10-30）：丟後台下載的採購單 PDF → 視窗裡下載畫好的
        pdoc = pymupdf.open(); ppg = pdoc.new_page(width=842, height=595); pxs = [20, 45, 120, 200, 400, 470, 530, 600]
        for (px, py), tx in [((300, 30), "PURCHASE ORDER"), ((40, 50), "Purchase Order ID"), ((160, 50), "POTWCTWA26102099M")]:
            ppg.insert_text((px, py), tx, fontname="helv", fontsize=7)
        for y in (80, 110, 150):
            ppg.draw_line((pxs[0], y), (pxs[-1], y), width=0.8)
        for x in pxs:
            ppg.draw_line((x, 80), (x, 150), width=0.8)
        for x, h, v in zip(pxs, ["No.", "Shopee SKU ID", "Supplier SKU ID", "Purchase QTY (unit)", "Unit Name", "Purchase QTY (pcs)", "Comment"],
                           ["1", "99_1", "M69990001_QMN", "24", "Pieces", "24", ""]):
            ppg.insert_text((x + 2, 98), h, fontname="helv", fontsize=7); ppg.insert_text((x + 2, 132), v, fontname="helv", fontsize=7)
        ppdf = os.path.join(os.environ["COUPANG_OMS_DATA"], "678050_POTWCTWA26102099M_PO.pdf"); pdoc.save(ppdf)
        close_side(); pg.set_input_files("#file", [ppdf])
        pg.wait_for_function("[...document.querySelectorAll('.toast')].some(t => t.innerText.includes('已收到 POTWCTWA26102099M 的蝦皮採購單'))", timeout=15000)
        pg.click("#k-tbl tr.og[data-k='POTWCTWA26102099M'] .fn"); pg.wait_for_selector("#side a.sd-podl", timeout=10000)
        check("丟蝦皮後台下載的採購單 PDF → 跳提示；視窗裡列出要畫的那品、有「下載畫好的採購單」", "678050_POTWCTWA26102099M_PO.pdf" in pg.inner_text("#side")
              and "M69990001_QMN" in pg.inner_text("#side .sd-sec:has(a.sd-podl)"), pg.inner_text("#side .sd-sec:has(a.sd-podl)")[:300])
        with pg.expect_download(timeout=15000) as dlp:
            pg.click("#side a.sd-podl")
        check("按下載 → 拿到「…_已畫.pdf」", dlp.value.suggested_filename == "678050_POTWCTWA26102099M_PO_已畫.pdf", dlp.value.suggested_filename)
        close_side()
        print("\n【6e-6】瑪氏新品不在商品總表：面板按「補商品資料」直接補，不用改 Excel、不用到瑪氏出貨頁")
        wbm = openpyxl.Workbook(); wsm = wbm.active; wsm.append(hdr)
        wsm.append(["PRTWCOTWA202610219998", "POTWCTWA26102198N", "TWA", "2026-10-21", "98_1", "M69990002_QMN", "【蝦皮直營】UI 測試 瑪氏新品", "24", "Pcs", "M69990002"])
        bm = io.BytesIO(); wbm.save(bm)
        up = tc.post("/api/shopee/import/preview", data={"file": (io.BytesIO(bm.getvalue()), "PurchaseOrder_ui_mars2.xlsx")}, content_type="multipart/form-data").get_json()
        tc.post("/api/shopee/import/commit", json={"upload_id": up["upload_id"], "fulfil": "供應商直送"})
        pg.reload(); pg.wait_for_selector("#stages .st.on", timeout=15000)
        pg.fill("#f-from", "2026-01-01"); pg.dispatch_event("#f-from", "change"); pg.wait_for_timeout(500)   # 重新整理後到貨日回到預設「今天往前 7 天」，假單 10/21 會被篩掉（2026-10-29 CI 紅燈）
        stage("buy"); pg.wait_for_selector("#b-tbl tr.og[data-k='POTWCTWA26102198N']", timeout=10000)
        check("待產採購單：這張寫「瑪氏商品總表找不到 1 項」", "瑪氏商品總表找不到 1 項" in pg.inner_text("#b-tbl tr.og[data-k='POTWCTWA26102198N']"))
        pg.click("#b-tbl tr.og[data-k='POTWCTWA26102198N'] .fn"); pg.wait_for_selector("#side .sd-mp", timeout=10000)
        check("右側那品：寫「請按「補商品資料」填寫」、有「補商品資料」按鈕", "請按「補商品資料」填寫" in pg.inner_text("#side") and pg.is_visible("#side .sd-mp"))
        pg.click("#side .sd-mp"); pg.wait_for_selector("#dlg-mp[open]", timeout=5000)
        check("補商品資料視窗：料號、單位盒、蝦皮品名、三個品類可選", "M69990002" in pg.inner_text("#dm-facts") and "盒" in pg.inner_text("#dm-facts")
              and pg.eval_on_selector_all("#dm-cat option", "els => els.length") == 4)
        pg.click("#dm-save"); pg.wait_for_timeout(300)
        check("沒選品類不能存（視窗還開著）", pg.is_visible("#dlg-mp"))
        pg.select_option("#dm-cat", "Petcare"); pg.fill("#dm-box", "12"); pg.click("#dm-ln"); pg.click("#dm-save")
        pg.wait_for_selector("#dlg-mp[open]", state="detached", timeout=5000)
        pg.wait_for_function("document.querySelector('#b-tbl tr.og[data-k=\"POTWCTWA26102198N\"]') && document.querySelector('#b-tbl tr.og[data-k=\"POTWCTWA26102198N\"]').innerText.includes('可產出')", timeout=10000)
        try:                                                   # 清單和右側面板是兩個請求，面板可能晚一點才換好
            pg.wait_for_function("document.querySelector('#side').innerText.includes('EIP 檔：盒_PET_不貼中標')", timeout=10000)
        except Exception:
            pass
        check("存好 → 那張單變「可產出」、右側寫 EIP 檔：盒_PET_不貼中標", "EIP 檔：盒_PET_不貼中標" in pg.inner_text("#side"), pg.inner_text("#side")[:300])
        pg.keyboard.press("Escape")
        print("\n【6e-7】歷程紀錄可以刪、資料與設定可以清除資料（只有管理員；Jerry 2026-10-14）")
        pg.click("#hs-open"); pg.wait_for_selector("#hs-dlg[open] #hs-table .hs-ck", timeout=10000); pg.wait_for_timeout(300)
        n0 = pg.eval_on_selector_all("#hs-table tbody tr", "els => els.length")
        check("管理員打開歷程：每列有勾選框、「全選目前列出的」、「刪除勾選的 0 筆」先鎖住", pg.is_visible("#hs-adm") and pg.is_visible("#hs-all")
              and pg.eval_on_selector_all("#hs-table .hs-ck", "els => els.length") == n0 > 0 and pg.is_disabled("#hs-del"))
        pg.click("#hs-table .hs-ck >> nth=0")
        check("勾一筆 → 按鈕寫「刪除勾選的 1 筆」", "刪除勾選的 1 筆" in pg.inner_text("#hs-del") and not pg.is_disabled("#hs-del"))
        pg.once("dialog", lambda dlg: dlg.accept()); pg.click("#hs-del")
        pg.wait_for_function("document.querySelector('#hs-count').innerText.includes('已刪除 1 筆')", timeout=10000)
        check("確認後刪掉那 1 筆、清單少一列", pg.eval_on_selector_all("#hs-table tbody tr", "els => els.length") == n0 - 1, pg.inner_text("#hs-count"))
        pg.click("#hs-close")
        pg.click("#btn-set"); pg.wait_for_selector("#view-set:not(.hidden)", timeout=5000); pg.click("#reset-box summary"); pg.wait_for_timeout(300)
        check("資料與設定最下面有清除資料：訂單資料、歷程預設勾，參考資料不勾", pg.is_visible("#rs-go") and pg.is_checked("#rs-orders")
              and not pg.is_checked("#rs-refs") and pg.is_checked("#rs-logs"))
        pg.click("#rs-go"); pg.wait_for_timeout(300)
        check("沒打「清空資料」→ 提示、不清", "請輸入「清空資料」以確認" in pg.inner_text("#toasts"))
        pg.fill("#rs-confirm", "清空資料"); pg.once("dialog", lambda dlg: dlg.accept()); pg.click("#rs-go")
        pg.wait_for_function("document.querySelector('#toasts').innerText.includes('已清除訂單資料')", timeout=10000)
        pg.click("#set-back"); pg.wait_for_selector("#view-work:not(.hidden)", timeout=5000); pg.wait_for_timeout(1000)
        stage("all")
        check("清完：全部訂單 0 張、清單寫「尚未匯入蝦皮訂單」", count("all") == 0 and "尚未匯入蝦皮訂單" in pg.inner_text("#tbl"), pg.inner_text("#tbl")[:120])
        print("\n【6g】酷澎 PO 視窗：修改紀錄（一輪一個 OK，小真 2026-10-20）、備註自己打 OK 會被擋")
        if not tc.get("/api/pos?page_size=5").get_json()["rows"]:          # 前面的段落可能把訂單清掉了：補匯範例
            pv_g = tc.post("/api/import/preview", data={"file": (open(os.path.join(BASE, "samples", "整合表範例.xlsx"), "rb"), "整合表範例.xlsx")}, content_type="multipart/form-data").get_json()
            tc.post("/api/import/commit", json={"operator": "小真", "batch_id": pv_g["batch_id"]})
        po_g = tc.get("/api/pos?page_size=5").get_json()["rows"][0]["po_number"]
        pg.goto(f"{base}/coupang"); pg.wait_for_selector("#lm-btn", timeout=15000)
        pg.evaluate(f"openPo('{po_g}')"); pg.wait_for_selector("#po-rounds #rd-ok", timeout=10000)
        check("PO 視窗有修改紀錄：OK 鈕一開始按不下去（主因必選）", pg.is_disabled("#rd-ok") and "處理 0 次" in pg.inner_text("#po-rounds"))
        pg.check("#po-rounds .rd-kind[value='qty']"); pg.wait_for_timeout(200)
        opts_g = pg.eval_on_selector_all("#po-rounds .rd-reason", "els => els.map(e => e.value)")
        pg.check("#po-rounds .rd-reason[value='庫存不足']"); pg.check("#po-rounds .rd-reason[value='停產']"); pg.wait_for_timeout(200)
        check("勾數量下修 → 主因只列第一階段的下修原因、可複選；勾了主因 OK 鈕才能按", "庫存不足" in opts_g and "約倉問題" not in opts_g and not pg.is_disabled("#rd-ok"), str(opts_g))
        # 先存著（小真 2026-10-21「還沒完成可以先暫存嗎」）
        pg.fill("#rd-note", "先存的說明"); pg.click("#rd-draft"); pg.wait_for_selector("#po-rounds .rd-draft-row", timeout=8000)
        check("按「先存著（尚未完成）」→ 清單多一列草稿、修改還是 0 次、主因還勾著、寫是誰存的", "處理 0 次" in pg.inner_text("#po-rounds") and pg.is_checked("#po-rounds .rd-reason[value='停產']")
              and "先存著的" in pg.inner_text("#po-rounds") and pg.query_selector("#rd-discard"), pg.inner_text("#po-rounds")[:300])
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        pg.evaluate(f"openPo('{po_g}')"); pg.wait_for_selector("#po-rounds #rd-ok", timeout=10000)
        check("關掉再開 → 草稿還在、主因和說明照草稿帶、OK 鈕能按", pg.query_selector("#po-rounds .rd-draft-row") and pg.is_checked("#po-rounds .rd-reason[value='庫存不足']")
              and pg.input_value("#rd-note") == "先存的說明" and not pg.is_disabled("#rd-ok"), pg.inner_text("#po-rounds")[:300])
        pg.fill("#rd-note", "")
        pg.click("#rd-ok"); pg.wait_for_function("document.querySelector('#po-rounds').innerText.includes('處理 1 次')", timeout=8000)
        check("按 OK → 處理 1 次、兩個主因用「、」接、備註不動（修改紀錄才是佐證）、草稿那列沒了", "庫存不足、停產" in pg.inner_text("#po-rounds") and "OK" not in pg.input_value("#p-remarks")
              and pg.query_selector("#po-rounds .rd-draft-row") is None, pg.inner_text("#po-rounds")[:200])
        pg.click("#po-rounds .rd-edit >> nth=0"); pg.wait_for_selector("#rd-cancel", timeout=5000)
        check("點鉛筆 → 表單帶入那一筆（主因勾著）、按鈕變「儲存這筆修改」", pg.is_checked("#po-rounds .rd-reason[value='停產']") and "儲存這筆修改" in pg.inner_text("#rd-ok"))
        pg.uncheck("#po-rounds .rd-reason[value='停產']"); pg.fill("#rd-note", "改過的說明"); pg.click("#rd-ok")
        pg.wait_for_function("document.querySelector('#po-rounds').innerText.includes('改過的說明')", timeout=8000)
        check("儲存 → 那一筆更新（主因剩庫存不足、說明換了）、回到記新的一輪", "庫存不足、停產" not in pg.inner_text("#po-rounds") and pg.query_selector("#rd-cancel") is None and "處理 1 次" in pg.inner_text("#po-rounds"), pg.inner_text("#po-rounds")[:200])
        check("備註分兩格：出貨前備註、出貨後備註（Jerry 2026-10-22）", pg.is_visible("#p-remarks") and pg.is_visible("#p-remarks-after")
              and "出貨前備註" in pg.inner_text("#dlg-po") and "出貨後備註" in pg.inner_text("#dlg-po"))
        pg.fill("#p-remarks", pg.input_value("#p-remarks") + " 再處理=>10/21-OK"); pg.click("#btn-po-save"); pg.wait_for_timeout(900)
        check("備註自己再打一個 OK 存檔 → 擋下來、視窗還開著", "備註裡寫了「OK」" in pg.inner_text("#toast") and pg.evaluate("document.getElementById('dlg-po').open"), pg.inner_text("#toast"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        # 處理完成之後又改（Jerry 2026-10-22「講人話」）：PO 視窗上方提示、清單 PO 單號下面的小標
        pg.wait_for_timeout(1100)
        pg.evaluate("""async po => { const d = await (await fetch(`/api/pos/${po}`)).json(); const s = d.skus.find(x => (x.qty_ship || 0) > 0);
            await fetch(`/api/orders/${s.id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ version: s.version, qty_ship: s.qty_ship - 1, operator: "小真" }) }); }""", po_g)
        pg.evaluate(f"openPo('{po_g}')"); pg.wait_for_selector("#po-rounds #rd-after", timeout=10000)
        check("處理完成後又改了數量 → PO 視窗修改紀錄上方黃色提示「MM-DD 處理完成後又改：出貨數量下修 1 項…記得再記一次處理完成」",
              "處理完成後又改：" in pg.inner_text("#rd-after") and "下修出貨數量 1 項" in pg.inner_text("#rd-after") and "再記一次處理完成" in pg.inner_text("#rd-after"), pg.inner_text("#rd-after"))
        pg.keyboard.press("Escape"); pg.wait_for_timeout(300)
        pg.check("#f-round"); pg.wait_for_timeout(1200)
        chip = f"#tbody tr:has(button[data-open='{po_g}']) .rt-chip"
        check("勾「只看處理完成後又有改動的（修改紀錄）」→ 那張單 PO 單號下面一行黃色標籤「M/D 完成後又改 出貨數量」，不寫人名數字（Jerry 2026-10-22 雛形）",
              "完成後又改" in pg.inner_text(chip) and "出貨數量" in pg.inner_text(chip) and "→" not in pg.inner_text(chip) and "rt-after" in pg.get_attribute(chip, "class")
              and pg.is_hidden(f"#tbody tr:has(button[data-open='{po_g}']) .rt-tip"), pg.inner_text("#tbody")[:300])
        pg.hover(chip); pg.wait_for_timeout(300)
        check("滑鼠移到標籤上 → 跳出完整內容（誰改了什麼、哪天）", pg.is_visible(f"#tbody tr:has(button[data-open='{po_g}']) .rt-tip")
              and "下修出貨數量 1 項" in pg.inner_text(f"#tbody tr:has(button[data-open='{po_g}']) .rt-tip"), pg.inner_text(f"#tbody tr:has(button[data-open='{po_g}']) .rt-tip"))
        pg.uncheck("#f-round"); pg.wait_for_timeout(600)
        # 快到貨還沒完成（Jerry 2026-10-27「OK 照這樣做」）：紅卡、到貨日紅字＋剩幾天、點卡只看這些單
        import due
        cn = db.get_conn()
        cn.execute("UPDATE po_headers SET po_status = ? WHERE po_number = ?", ("已建立", po_g))
        cn.execute("UPDATE orders SET delivery_date = ? WHERE po_number = ?", (due.cutoff(datetime.date.fromisoformat(db.today())).isoformat(), po_g))
        cn.commit(); cn.close()
        pg.goto(f"{base}/coupang"); pg.wait_for_selector("#tbody tr", timeout=15000); pg.wait_for_timeout(500)
        n_due = int(pg.inner_text("#stat-due")); PG_SIZE = pg.evaluate("PAGE_SIZE")
        pg.click("#stat-due-card"); pg.wait_for_timeout(1200)
        row_g = f"#tbody tr:has(button[data-open='{po_g}'])"
        check("上面有紅卡「快到貨還沒完成（3 個工作天內）」；點一下 → 勾起「只看快到貨還沒完成」、清單只剩這些單",
              n_due >= 1 and "快到貨還沒完成" in pg.inner_text("#stat-due-card") and pg.is_checked("#f-due")
              and pg.locator("#tbody tr:has(button[data-open])").count() == min(n_due, PG_SIZE) and pg.query_selector(row_g), f"{n_due} {pg.inner_text('#tbody')[:200]}")
        check("那張單的到貨日紅字、旁邊寫「剩 3 天」", pg.inner_text(f"{row_g} .due-chip") == "剩 3 天"
              and pg.evaluate(f"getComputedStyle(document.querySelector(\"{row_g} .due-d\")).color") == "rgb(200, 30, 30)")
        pg.uncheck("#f-due"); pg.wait_for_timeout(600)
        pg.goto(f"{base}/coupang?due=1"); pg.wait_for_selector("#tbody tr", timeout=15000); pg.wait_for_timeout(800)
        check("首頁今日事項連過來（/coupang?due=1）→ 已經勾好、只看這些單", pg.is_checked("#f-due") and pg.locator("#tbody tr:has(button[data-open])").count() == min(n_due, PG_SIZE))
        pg.uncheck("#f-due"); pg.wait_for_timeout(600)

        print("\n【6h】MOMO 訂單管理系統（2026-10-28）：一次丟兩個下載檔 → 預覽 → 匯入 → 搜尋 → 右側面板 → 產 EIP → 回填")
        mhdr = ["項次", "倉別", "到倉日", "採購日", "入庫單號", "入庫預定日期", "採購總量", "商品編號", "商品名稱", "單品編號", "單品詳細", "採購數量", "商品原廠編號"]
        mfiles = []
        for k, (whm, nom, items) in enumerate([("DC6", "26109999-0001", [("11000001", "【彩豆豆】牛奶巧克力", "48", "M89297484_QMN")]),
                                               ("SDC", "26109999-0002", [("11000002", "【口香糖王】薄荷", "200", "M81232885_CMN"),
                                                                         ("11000003", "【喵愛餡】貓點心 箱購", "3", "M69072565_RMN")])]):
            wbm = openpyxl.Workbook(); wsm = wbm.active; wsm.append(mhdr)
            for j, (prod, nm, q, code) in enumerate(items):
                wsm.append([str(k + 1) if j == 0 else "", whm if j == 0 else "", "", "2026/10/20" if j == 0 else "", nom if j == 0 else "",
                            "2026/10/30" if j == 0 else "", str(sum(int(x[2]) for x in items)) if j == 0 else "", prod, nm, "001", "無", q, code])
            fp = os.path.join(os.environ["COUPANG_OMS_DATA"], f"UnsendList_{k + 1}.xlsx"); wbm.save(fp); mfiles.append(fp)
        tc.post("/api/mars/products/import", data={"file": (open(os.path.join(BASE, "samples", "mars", "fake", "假_瑪氏商品總表.xlsx"), "rb"), "假_瑪氏商品總表.xlsx")},
                content_type="multipart/form-data")
        pg.goto(f"{base}/momo"); pg.wait_for_selector("#stages .st", timeout=15000)
        pg.set_input_files("#file", mfiles); pg.wait_for_selector(".pv [data-a=ok]", timeout=15000)
        check("一次選兩個 MOMO 下載檔 → 合在一起預覽：2 個檔、2 張入庫單、新增 3", "MOMO 下載檔 2 個" in pg.inner_text(".pv") and "入庫單 2 張" in pg.inner_text(".pv")
              and "新增 3" in pg.inner_text(".pv"), pg.inner_text(".pv")[:300])
        pg.click(".pv [data-a=ok]"); pg.wait_for_function("document.querySelectorAll('#tbl tr.og').length === 2", timeout=15000)
        check("確認匯入 → 待產 EIP 採購單 2 張、清單兩列、都可產出", pg.inner_text("#stages .st.on .n").strip() == "2" and pg.locator("#tbl tr.og").count() == 2
              and pg.inner_text("#tbl").count("可產出") == 2, pg.inner_text("#tbl")[:300])
        pg.fill("#f-q", "26109999-0002"); pg.wait_for_function("document.querySelectorAll('#tbl tr.og').length === 1", timeout=8000)
        check("搜尋入庫單號：打字就出結果、只剩那張", "26109999-0002" in pg.inner_text("#tbl"))
        pg.fill("#f-q", ""); pg.wait_for_function("document.querySelectorAll('#tbl tr.og').length === 2", timeout=8000)
        pg.click("#pk-line [data-pk]"); pg.check("#pk-line .pick-pop input[value='瑪氏']"); pg.wait_for_timeout(800)
        check("線別複選：勾瑪氏馬上篩、選單還開著可以再勾", pg.locator("#tbl tr.og").count() == 2 and "瑪氏" in pg.inner_text("#pk-line [data-pk]")
              and pg.is_visible("#pk-line .pick-pop"))
        pg.check("#pk-line .pick-pop input[value='紙品']"); pg.wait_for_timeout(800)
        check("再勾紙品：兩個都勾著（瑪氏、紙品）", "瑪氏、紙品" in pg.inner_text("#pk-line [data-pk]") and pg.locator("#tbl tr.og").count() == 2)
        pg.uncheck("#pk-line .pick-pop input[value='紙品']")
        pg.uncheck("#pk-line .pick-pop input[value='瑪氏']"); pg.click("body", position={"x": 5, "y": 5}); pg.wait_for_timeout(600)
        pg.click("#tbl tr.og[data-no='26109999-0002']"); pg.wait_for_selector("#side:not(.hidden) #sd-items tr.sd-it", timeout=10000)
        check("點一張入庫單 → 畫面中間打開入庫單視窗（2026-10-30 從右側面板改）：出貨方式、品項一品一列（200 包 ÷ 200 ＝ 1 箱）、倉別資料、上一張／下一張",
              "出貨方式" in pg.inner_text("#sd-title") and "品項與下採" in pg.inner_text("#sd-body") and "1 箱" in pg.inner_text("#sd-items")
              and "台南市新市區" in pg.inner_text("#sd-body") and "／2 張" in pg.inner_text("#sd-pos"), pg.inner_text("#side")[:400])
        p0 = pg.inner_text("#sd-pos"); pg.click("#sd-next" if "第 1" in p0 else "#sd-prev"); pg.wait_for_timeout(800)
        check("按下一張（或上一張）→ 換成另一張入庫單", "26109999-0002" not in pg.inner_text("#sd-title") and pg.inner_text("#sd-pos") != p0, pg.inner_text("#sd-title")[:100])
        pg.keyboard.press("Escape"); pg.wait_for_selector("#side", state="hidden", timeout=5000)
        pg.check("#ck-all")
        with pg.expect_download(timeout=15000) as dlm:
            pg.click("#b-go")
        check("勾全部 → 產出 EIP 採購單（壓縮檔）", dlm.value.suggested_filename.startswith("MOMO_EIP採購單_1030"), dlm.value.suggested_filename)
        pg.wait_for_timeout(800)
        pg.click("#stages .st[data-st='eip']"); pg.wait_for_selector("#e-tbl .eip-in", timeout=8000)
        n_eip = pg.locator("#e-tbl .eip-in").count()
        check("待回填 EIP 採購單號：瑪氏依倉、品類、箱／盒包分成 3 份", n_eip == 3, str(n_eip))
        pg.fill("#e-tbl .eip-in >> nth=0", "PO202610777"); pg.click("#e-tbl .eip-save >> nth=0"); pg.wait_for_timeout(1000)
        check("回填一份 → 待回填剩 2 份", pg.locator("#e-tbl .eip-in").count() == 2)
        pg.click("#e-all"); pg.wait_for_timeout(800)
        check("顯示已回填 → 那份出現「瑪氏採購單」下載", "瑪氏採購單" in pg.inner_text("#e-tbl"), pg.inner_text("#e-tbl")[:300])

        print("\n【6i】MOMO ③ 待核缺貨與效期：丟勇信 PDF → 點單 → 兩批一樣多點一批 → 填效期匯入檔下載")
        import json as _json
        cm = db.get_conn()
        em = dict(cm.execute("SELECT * FROM momo_eip WHERE eip_po = ?", ("PO202610777",)).fetchone())
        erows = _json.loads(em["rows_json"])
        oids = [x["order_id"] for r_ in erows for x in r_["parts"]]
        ords = [dict(cm.execute("SELECT * FROM momo_orders WHERE id = ?", (i,)).fetchone()) for i in oids]
        cm.close()
        ydoc = pymupdf.open(); ypg = ydoc.new_page(); yy = 40
        ylines = ["配 送 明 細 表(預排鮮度)", "收貨單號：PO202610777-momo"]
        ytot = 0
        for r_ in erows:
            q = int(r_["qty"]); a, b_ = q // 2 or 1, q - (q // 2 or 1)
            ylines += [f"58880{r_['mars_code']}", "1", "假品名", f"{a}C", "X", "2028/02/08"] + ([f"{b_}C", "X", "2028/02/09"] if b_ else [])
            ytot += q
        ylines.append(f"合計箱數:{ytot}")
        for ln in ylines:
            ypg.insert_text((40, yy), ln, fontname="china-t", fontsize=9); yy += 14
        ypdf = os.path.join(os.environ["COUPANG_OMS_DATA"], "勇信配送明細表.pdf"); ydoc.save(ypdf)
        pg.goto(f"{base}/momo"); pg.wait_for_selector("#stages .st", timeout=15000)
        pg.set_input_files("#file", [ypdf]); pg.wait_for_selector(".pv h3:has-text('勇信配送明細表')", timeout=15000)
        pg.wait_for_selector("#stages .st.on[data-st='yx']", timeout=8000); pg.wait_for_selector("#tbl tr.og", timeout=8000)
        no0 = ords[0]["inbound_no"]
        check("丟勇信 PDF → 上傳區列出核對結果、跳到「待核缺貨與效期」、那張單已核對", no0 in pg.inner_text(".pv") and "已核對" in pg.inner_text(f"#tbl tr.og[data-no='{no0}']"),
              pg.inner_text(".pv")[:300])
        pg.click(f"#tbl tr.og[data-no='{no0}']"); pg.wait_for_selector("#sd-items th:has-text('③ 結果')", timeout=10000)
        ntie = pg.locator("#sd-items [data-pick]").count()
        check("入庫單視窗：下採和 ③ 在同一列，兩批數量一樣的標「要決定」、排在最前面、效期可以點", ntie >= 2 and "要決定" in pg.inner_text("#sd-items")
              and "ask" in (pg.get_attribute("#sd-items tr.sd-it >> nth=0", "class") or ""), pg.inner_text("#sd-items")[:300])
        pg.fill("#sd-q", ords[0]["factory_code"]); pg.wait_for_timeout(300)
        check("視窗裡搜尋原廠編號 → 只剩那品、寫「找到 1 品」", pg.locator("#sd-items tr.sd-it").count() == 1 and "找到 1 品" in pg.inner_text("#sd-found"), pg.inner_text("#sd-found"))
        pg.fill("#sd-q", ""); pg.wait_for_timeout(300)
        for _ in range(10):                                # 選過的列效期還能再點（可以改），所以只點還在「要決定」的列；最多 10 次
            if not pg.locator("#sd-items tr.ask [data-pick]").count():
                break
            pg.click("#sd-items tr.ask [data-pick] >> nth=0"); pg.wait_for_timeout(900)
        check("點一批的效期 → 寫入效期、不用再決定", "要決定" not in pg.inner_text("#sd-items") and "寫入效期" in pg.inner_text("#sd-items"), pg.inner_text("#sd-items")[:300])
        xh = ["倉別", "入庫單號", "入庫預定日期", "商品編號", "商品名稱", "單品編號", "商品有效日期", "入庫最低有效期限", "採購數量", "商品原廠編號"]
        wbx = openpyxl.Workbook(); wsx = wbx.active; wsx.append(xh)
        for j, o_ in enumerate(ords):
            wsx.append([o_["warehouse"] if j == 0 else "", o_["inbound_no"] if j == 0 else "", "2026/10/30" if j == 0 else "", o_["product_no"], o_["product_name"],
                        o_["item_no"], "", "2027/06/20", str(o_["qty"]), o_["factory_code"]])
        xfp = os.path.join(os.environ["COUPANG_OMS_DATA"], "UnsendList_效期.xlsx"); wbx.save(xfp)
        with pg.expect_download(timeout=15000) as dlx:
            pg.set_input_files("#ck-file", xfp)
        wsd = openpyxl.load_workbook(io.BytesIO(open(dlx.value.path(), "rb").read())).active
        check("視窗裡按「下載填好效期的匯入檔」選檔 → 下載 _已填效期、商品有效日期填好", dlx.value.suggested_filename == "UnsendList_效期_已填效期.xlsx"
              and wsd.cell(2, 7).value == "2028/02/08", f"{dlx.value.suggested_filename} {wsd.cell(2, 7).value}")
        pg.wait_for_timeout(800)
        rowx = pg.locator(f"#tbl tr.og[data-no='{no0}']")
        check("填好後這張單離開「待核缺貨與效期」（全部出貨、效期填過就完成）", rowx.count() == 0 or "已填" in rowx.inner_text(), str(rowx.count()))
        pg.click("#sd-close")

        print("\n【6f】帳號選單：每個系統點名字都有修改密碼、帳號管理、登出")
        for path, ready in [("/", "#chs .ch"), ("/coupang", "#lm-btn"), ("/shopee", "#stages .st"), ("/momo", "#stages .st"), ("/master", "#dd-line-btn"),
                            ("/mars", "#sp-table thead"), ("/zhuyun", "#dz"), ("/purchase", "#ac-btn")]:
            pg.goto(f"{base}{path}"); pg.wait_for_selector(ready, state="attached", timeout=15000)
            pg.click("#ac-btn"); pg.wait_for_selector("#ac-menu:not([hidden])", timeout=5000); pg.wait_for_selector("#ac-users-item:not([hidden])", timeout=5000)
            m = pg.inner_text("#ac-menu")
            check(f"{path}：名字選單有 Jerry、修改密碼、帳號管理（管理員）、登出", all(w in m for w in ("Jerry", "修改密碼", "帳號管理", "登出")), m)
            pg.keyboard.press("Escape")
        pg.goto(f"{base}/coupang"); pg.wait_for_selector("#btn-settings:not(.hidden)", timeout=10000)
        pg.click("#btn-settings"); pg.wait_for_selector("#dlg-settings[open]", timeout=5000)
        tabs = pg.inner_text("#dlg-settings .st-tab >> nth=0") + "／" + pg.inner_text("#dlg-settings .st-tab >> nth=1")
        check("酷澎齒輪只剩酷澎自己的：清空訂單資料、驗收同步", pg.eval_on_selector_all("#dlg-settings .st-tab", "els => els.length") == 2
              and tabs == "清空訂單資料／驗收同步", tabs)
        pg.keyboard.press("Escape")
        pg.goto(f"{base}/shopee"); pg.wait_for_selector("#stages .st")
        pg.click("#ac-btn"); pg.click("#ac-menu [data-ac='password']"); pg.wait_for_selector("#ac-dlg[open]", timeout=5000)
        pg.fill("#ac-old", "changeme123"); pg.fill("#ac-new", "changeme456"); pg.fill("#ac-new2", "changeme456"); pg.click("#ac-save-pw"); pg.wait_for_timeout(600)
        check("蝦皮頁改密碼成功", "已更新" in pg.inner_text("#ac-msg"), pg.inner_text("#ac-msg"))
        pg.fill("#ac-old", "changeme456"); pg.fill("#ac-new", "changeme123"); pg.fill("#ac-new2", "changeme123"); pg.click("#ac-save-pw"); pg.wait_for_timeout(600)
        check("再改回原本的密碼", "已更新" in pg.inner_text("#ac-msg"), pg.inner_text("#ac-msg"))
        pg.click("#ac-dlg .ac-tabs [data-tab='users']"); pg.wait_for_selector("#ac-userlist > div", timeout=5000)
        pg.fill("#ac-newuser", "介面測試員"); pg.fill("#ac-newpass", "abc12345"); pg.click("#ac-add-user"); pg.wait_for_timeout(600)
        check("帳號管理：新增帳號後列表看得到", "介面測試員" in pg.inner_text("#ac-userlist"), pg.inner_text("#ac-userlist"))
        pg.once("dialog", lambda d: d.accept()); pg.click("#ac-userlist [data-del='介面測試員']"); pg.wait_for_timeout(600)
        check("帳號管理：刪除後列表沒有了", "介面測試員" not in pg.inner_text("#ac-userlist"), pg.inner_text("#ac-userlist"))
        pg.click("#ac-close")
        pg.goto(f"{base}/purchase"); pg.wait_for_selector("#ac-btn"); pg.click("#ac-btn"); pg.click("#ac-logout"); pg.wait_for_load_state()
        check("採購表頁按登出 → 回到登入頁", "/login" in pg.url, pg.url)
        bad = [x for x in bad if ("/api/mars/splits/" not in x[1] or x[0] != 400) and not (x[1].endswith("/api/mars/emma") and x[0] == 409)
               and not (x[0] == 400 and ("/api/shopee/purchase/adjust" in x[1] or "/api/shopee/eips/" in x[1] or "/api/pos/" in x[1]))]   # 故意填錯的 400、EMMA 先問的 409 不算
        b.close()

    print("\n【7】整體")
    check("整個過程沒有 JavaScript 錯誤", not errs, str(errs[:3]))
    check("整個過程沒有 4xx／5xx", not bad, str(bad[:3]))
    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]; sys.exit(1)


if __name__ == "__main__":
    main()
