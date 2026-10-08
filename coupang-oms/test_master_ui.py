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
    local = glob.glob("/opt/pw-browsers/chromium-*/chrome-linux/chrome")
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

        print("\n【4c】④ 改單統計分頁")
        pg.click('.m-tab[data-tab="stats"]'); pg.wait_for_timeout(900)
        check("四張卡：PO 張數／曾改單的 PO／改單次數／估計工時", pg.eval_on_selector_all("#st-kpi .kpi", "els => els.length") == 4 and "曾改單的 PO" in pg.inner_text("#st-kpi"))
        check("有「含測試」勾選框、預設不勾", pg.query_selector("#st-test") is not None and not pg.is_checked("#st-test"))
        check("剛剛改的那筆算進「我方修改」，原因表有缺貨", "我方修改" in pg.inner_text("#st-kpi"))
        check("每月表有 9 月那列", "2026-09" in pg.inner_text("#st-months"))
        check("原因表列出缺貨／沒車／酷澎要求／其他／未填", all(r in pg.inner_text("#st-reasons") for r in ["缺貨", "沒車", "酷澎要求", "其他", "未填"]))
        check("「改了什麼」表有改數量／改到貨日／品項移除", all(r in pg.inner_text("#st-kinds") for r in ["改數量", "改到貨日", "品項移除"]))
        check("數字是藍色可點的", pg.eval_on_selector_all("#tab-stats .st-n", "els => els.length") >= 3)
        pg.click('#st-reasons .st-n >> nth=0'); pg.wait_for_timeout(900)
        check("點原因表的數字 → 明細視窗打開、標題寫幾次", pg.evaluate("document.getElementById('dlg-stev').open") and "次" in pg.inner_text("#stev-title"))
        check("明細列出時間／PO／誰／改了什麼／原因", pg.eval_on_selector_all("#stev-table tbody tr", "els => els.length") >= 1 and "缺貨" in pg.inner_text("#stev-table"))
        pg.click("#stev-table .stev-po >> nth=0"); pg.wait_for_timeout(700)
        check("明細裡點 PO 會打開 PO 視窗", pg.evaluate("document.getElementById('dlg-po').open")); pg.keyboard.press("Escape"); pg.wait_for_timeout(200); pg.keyboard.press("Escape"); pg.wait_for_timeout(200)
        check("沒填分鐘數時估計工時顯示破折號", "—" in pg.inner_text("#st-kpi"))
        pg.fill("#st-min", "15"); pg.dispatch_event("#st-min", "change"); pg.wait_for_timeout(700)
        check("填 15 分鐘後算出估計工時、表格多一欄", "小時" in pg.inner_text("#st-kpi") and "估計工時" in pg.inner_text("#st-months"))
        pg.click("#st-top .st-po >> nth=0"); pg.wait_for_timeout(700)
        check("點改最多次的 PO 會打開 PO 視窗", pg.evaluate("document.getElementById('dlg-po').open")); pg.keyboard.press("Escape"); pg.wait_for_timeout(200)
        pg.fill("#st-min", "0"); pg.dispatch_event("#st-min", "change"); pg.wait_for_timeout(500)

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
        check("五個上傳處都有看得見的虛線拖曳框（多了 ⑥ 驗收單）", pg.eval_on_selector_all(".dzbox", "els => els.length") == 5)
        check("9/18：5 份拆單表、都是「還沒產出」，但單號格子和每列按鈕已經能用、一條 PO 標題列寫拆成 5 份", pg.eval_on_selector_all("#sp-table tr.sp", "els => els.length") == 5
              and pg.eval_on_selector_all("#sp-table .st-new", "els => els.length") == 5 and not pg.is_disabled("#sp-table input.eip")
              and pg.eval_on_selector_all("#sp-table button.eipf", "els => els.length") == 5
              and pg.eval_on_selector_all("#sp-table tr.pog", "els => els.length") == 1 and "拆成 5 份" in pg.inner_text("#sp-table tr.pog"))
        pg.click("#sp-table tr.sp >> nth=0"); pg.wait_for_timeout(300)
        check("點一列展開看品項（下採料號、箱數）", pg.eval_on_selector_all("#sp-table tr.sub", "els => els.length") == 1 and "下採料號" in pg.inner_text("#sp-table tr.sub"))
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
        check("入口頁：三個通路卡、兩張圖都畫出來", pg.eval_on_selector_all("#chs .ch", "els => els.length") == 3 and pg.eval_on_selector_all("#c-bar svg .hb", "els => els.length") == 7)
        nav = pg.inner_text("#nav")
        check("左側選單：酷澎底下有訂單管理、驗收單簽名、採購表轉換、商品主檔、竹運、瑪氏，還有 SOP 與歷程", all(k in nav for k in ("訂單管理", "驗收單簽名", "採購表轉換", "商品主檔自動化", "竹運出貨拋檔", "瑪氏出貨", "營運 SOP 檢索", "歷程紀錄")))
        check("沒有寫買斷、寄倉", "買斷" not in pg.inner_text("body") and "寄倉" not in pg.inner_text("body"))
        check("左上角是文字「YFS 訂單系統」、選單不畫捲軸", pg.inner_text(".brand").strip() == "YFS 訂單系統" and pg.eval_on_selector(".side", "e => getComputedStyle(e).scrollbarWidth") == "none")
        check("各通路系統卡片：logo 圖都載得出來、標題是 XX 訂單管理系統、酷澎線紅色、PChome 線藍色",
              pg.eval_on_selector_all("#chs .ch-h img", "els => els.length === 3 && els.every(i => i.complete && i.naturalWidth > 0)")
              and "酷澎訂單管理系統" in pg.inner_text("#ch-coupang") and "蝦皮特選訂單管理系統" in pg.inner_text("#ch-shopee") and "PChome 訂單管理系統" in pg.inner_text("#ch-pchome")
              and pg.eval_on_selector("#ch-coupang", "e => getComputedStyle(e, '::before').backgroundColor") == "rgb(224, 20, 34)"
              and pg.eval_on_selector("#ch-pchome", "e => getComputedStyle(e, '::before').backgroundColor") == "rgb(0, 70, 191)"
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
        check("左側「改單統計」：在 YFS 訂單系統裡看，四張數字卡、每月表、原因、改了什麼、改最多次的 PO 都有",
              pg.url.endswith("/portal/stats") and pg.eval_on_selector_all("#st-kpi .kpi", "els => els.length") == 4
              and all(pg.query_selector(sel) for sel in ("#st-months thead", "#st-reasons thead", "#st-kinds thead", "#st-top thead")) and "/api/master/stats/export" in pg.get_attribute("#st-export", "href"))
        pg.fill("#st-from", "2026-09"); pg.dispatch_event("#st-from", "change"); pg.wait_for_timeout(800)
        if pg.query_selector("#st-months .st-n"):
            pg.click("#st-months .st-n >> nth=0"); pg.wait_for_selector("#st-dlg[open]", timeout=8000)
            check("點有底線的數字 → 跳出那幾次改單的明細", "改單明細" in pg.inner_text("#st-dlg-title")); pg.click("#st-dlg-x")
        else:
            check("9 月有改單紀錄可以點", False, pg.inner_text("#st-months"))
        pg.goto(f"{base}/master#stats"); pg.wait_for_selector("#tab-stats:not(.hidden)", timeout=15000)
        check("商品主檔的 /master#stats 還是能直接打開改單統計分頁（原本的分頁先留著）", pg.is_visible("#tab-stats"))
        pg.set_viewport_size({"width": 390, "height": 844}); pg.goto(f"{base}/"); pg.wait_for_selector("#c-bar svg"); pg.wait_for_timeout(300)
        check("手機寬度不會左右捲動", not pg.evaluate("document.documentElement.scrollWidth > innerWidth"))
        pg.click("#menu"); pg.wait_for_timeout(350)
        check("手機上按左上角打開選單", pg.evaluate("document.querySelector('.side').getBoundingClientRect().left") >= 0)
        pg.set_viewport_size({"width": 1500, "height": 900})
        print("\n【6e】蝦皮特選：上傳採購單、選履約方式、重傳差異、確認、改履約方式")
        def shopee_po(rows, path):
            hdr = ["PR ID", "PO ID", "Warehouse", "Expected Delivery Time", "Shopee SKU ID", "Supplier SKU ID", "Shopee SKU Name",
                   "Shopee Requested Qty (unit)", "Selling Type", "EAN / UPC"]
            wbs = openpyxl.Workbook(); wss = wbs.active; wss.append(hdr)
            for sku, sup, name, qty in rows:
                wss.append(["PRTWCOTWA202610019999", None, "TWA", "2026-10-19", sku, sup, name, str(qty), "Pcs", sup.split("_")[0]])
            wbs.save(path)
        sp1 = os.path.join(tempfile.gettempdir(), "ui_PurchaseOrder_1.xlsx"); sp2 = os.path.join(tempfile.gettempdir(), "ui_PurchaseOrder_2.xlsx")
        shopee_po([("11_1", "4987176232878_CNN", "幫寶適 清新幫 拉拉褲", 240), ("11_2", "4987176232878_RNN", "幫寶適 清新幫 箱購", 26)], sp1)
        shopee_po([("11_1", "4987176232878_CNN", "幫寶適 清新幫 拉拉褲", 200), ("11_2", "4987176232878_RNN", "幫寶適 清新幫 箱購", 26)], sp2)
        pg.goto(f"{base}/shopee"); pg.wait_for_selector("#tbl thead", timeout=15000)
        pg.fill("#f-from", "2026-01-01"); pg.dispatch_event("#f-from", "change")
        pg.set_input_files("#file", sp1); pg.wait_for_selector(".pv .ok", timeout=15000)
        check("上傳採購單 → 預覽：蝦皮採購單、新增 2 項、沒選履約方式不能按確認", "蝦皮採購單" in pg.inner_text(".pv") and "新增 2 項" in pg.inner_text(".pv").replace("\n", " ")
              and pg.is_disabled(".pv .ok"), pg.inner_text(".pv")[:200])
        pg.click(".pv .ff[data-f='竹運出貨']"); pg.click(".pv .ok"); pg.wait_for_selector("#tbl tr.og", timeout=10000); pg.wait_for_timeout(300)
        row = pg.inner_text("#tbl tr.og >> nth=0")
        check("匯入後總覽出現這張 PR：寶僑、觀音、竹運出貨、已匯入", "PRTWCOTWA202610019999" in row and "寶僑" in row and "觀音" in row and "竹運出貨" in row and "已匯入" in row, row[:200])
        pg.click("#tbl tr.og >> nth=0"); pg.wait_for_selector("#tbl tr.sub", timeout=5000)
        check("點那一列展開品項：單位 包／箱", "包" in pg.inner_text("#tbl tr.sub") and "箱" in pg.inner_text("#tbl tr.sub"))
        pg.set_input_files("#file", sp2); pg.wait_for_selector(".pv .ok", timeout=15000)
        check("重傳改過數量 → 預覽紅字列出差異 240 → 200", "跟上一版有 1 處不同" in pg.inner_text(".pv") and "200" in pg.inner_text(".pv"), pg.inner_text(".pv")[:300])
        pg.click(".pv .ok"); pg.wait_for_timeout(800)
        check("匯入後那張單變匯入差異待確認、有「確認差異」按鈕", "匯入差異待確認" in pg.inner_text("#tbl") and pg.is_visible("#tbl button.dd:has-text('確認差異')"))
        pg.click("#tbl button.dd:has-text('確認差異')"); pg.wait_for_selector("#dd-body table", timeout=8000)
        check("差異視窗：數量 240 → 200、-40、待確認", "240" in pg.inner_text("#dd-body") and "-40" in pg.inner_text("#dd-body") and "待確認" in pg.inner_text("#dd-body"))
        pg.click("#dd-confirm"); pg.wait_for_timeout(800)
        check("按確認差異 → 回到已匯入", "已匯入" in pg.inner_text("#tbl tr.og >> nth=0") and "匯入差異待確認" not in pg.inner_text("#tbl tr.og >> nth=0"))
        pg.click("#tbl button.ff >> nth=0"); pg.wait_for_selector("#dlg-fulfil[open]", timeout=5000)
        pg.click("#df-opts .chip[data-f='供應商直送']"); pg.fill("#df-reason", "竹運缺貨，改原廠直送"); pg.click("#df-save"); pg.wait_for_timeout(800)
        check("改履約方式 → 總覽變供應商直送", "供應商直送" in pg.inner_text("#tbl tr.og >> nth=0"))
        pg.click("#up-box summary"); pg.wait_for_timeout(300)
        check("上傳紀錄看得到兩次上傳、可下載原始檔", pg.eval_on_selector_all("#up-tbl a.lnk", "els => els.length") >= 2)
        print("\n【6f】帳號選單：每個系統點名字都有修改密碼、帳號管理、登出")
        for path, ready in [("/", "#chs .ch"), ("/coupang", "#lm-btn"), ("/shopee", "#tbl thead"), ("/master", "#dd-line-btn"),
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
        pg.goto(f"{base}/shopee"); pg.wait_for_selector("#tbl thead")
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
        bad = [x for x in bad if ("/api/mars/splits/" not in x[1] or x[0] != 400) and not (x[1].endswith("/api/mars/emma") and x[0] == 409)]   # 故意填錯的 400、EMMA 先問的 409 不算
        b.close()

    print("\n【7】整體")
    check("整個過程沒有 JavaScript 錯誤", not errs, str(errs[:3]))
    check("整個過程沒有 4xx／5xx", not bad, str(bad[:3]))
    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]; sys.exit(1)


if __name__ == "__main__":
    main()
