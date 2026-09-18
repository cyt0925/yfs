"""商品主檔自動化 前端煙霧測試（Playwright 開真的瀏覽器）。

後端測試（test_master.py）看不到 JavaScript 壞掉——今天就漏過「選回第一條沒反應」和
「誤刪 loadLogs」兩個 bug。這支把人每次手動點的那幾條路寫成腳本：登入、三個分頁、
篩選列、各次匯入的下拉、PO 視窗、就地編輯、上傳預覽、三個匯出連結，任何 JS 錯誤或
4xx/5xx 都算失敗。

執行：python test_master_ui.py
需要：pip install playwright && playwright install chromium（CI 會自動裝）。
資料：samples/master/fake/ 的假資料（商品、條碼、PO 全是編的）。
"""
import glob
import io
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
        check("訂單明細有資料", "筆" in pg.inner_text("#orders-count"))

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
        check("「算不出箱數」「人工調整過」有事才出現（假資料有算不出箱數的品項 → 顯示筆數）", "筆算不出箱數" in pg.inner_text("#chip-missing") or pg.evaluate("document.getElementById('chip-missing').classList.contains('hidden')"))

        print("\n【4】行事曆、倉別、PO 視窗、就地編輯")
        pg.click('.day.has[data-date="2026-09-04"]'); pg.wait_for_timeout(700)
        check("點日期後倉別區塊顯示已勾 1 天", "已勾 1 天" in pg.inner_text("#wh-block"))
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
        pg.click("#btn-clear"); pg.wait_for_timeout(500)

        print("\n【5】其他分頁與視窗")
        for tab in ["products", "summary", "orders"]:
            pg.click(f'.m-tab[data-tab="{tab}"]'); pg.wait_for_timeout(700)
        pg.click('.m-tab[data-tab="summary"]'); pg.wait_for_timeout(500); pg.select_option("#sum-line", "寶僑"); pg.wait_for_timeout(900)
        check("③ 總表顯示照哪份底稿匯出", "假_寶僑總表" in pg.inner_text("#tpl-info"))
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
        for sel in ["#btn-export-daily", "#btn-export-2", "#btn-export-daily-2"]:
            if sel != "#btn-export-daily":
                pg.click('.m-tab[data-tab="summary"]'); pg.wait_for_timeout(400)
            href = pg.get_attribute(sel, "href")
            req = urllib.request.Request(base + href, headers={"Cookie": "; ".join(f"{k}={v}" for k, v in cookies.items())})
            with urllib.request.urlopen(req) as resp:
                check(f"{sel} 下載 200 且是 Excel", resp.status == 200 and "spreadsheet" in resp.headers.get("Content-Type", ""))
        b.close()

    print("\n【7】整體")
    check("整個過程沒有 JavaScript 錯誤", not errs, str(errs[:3]))
    check("整個過程沒有 4xx／5xx", not bad, str(bad[:3]))
    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]; sys.exit(1)


if __name__ == "__main__":
    main()
