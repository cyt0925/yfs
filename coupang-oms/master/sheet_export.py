"""③ 匯出總表：把系統算好的數字填回業務的總表底稿，再附一張「系統看板」分頁。

原則（Jerry：公式與判讀才是重點，畫面跟匯出的數字不能對不起來）：
  1. 底稿是業務的：列、商品順序、其他月份、業務自己的公式一律不碰，只填「系統有值」的格子。
  2. 填哪些（用表頭文字認欄，不靠欄位字母）：
       各交貨日箱數＋該月 TTL          訂單明細現算（原本就有，缺的日期欄自動插）
       GIV／NIV                      主檔（Coupang Master 更新過的）；系統沒值的格不動
       Supply_CS／demand_CS（該月）   mst_month_stats；supply 表還沒上傳就不動（維持底稿原本的內容）
       目前庫存數量／目前庫存天數／到N月底庫存天數
                                     系統照庫存銷售表的公式現算，有庫存資料的商品才填，「到9月底」標題改成這個月。
                                     底稿這三格的 VLOOKUP 指歪了（見 docs/寶僑總表_欄位與公式對照.md），填成值才會對
       品牌區 目標／REBATE目標         mst_brand_targets（一季一個數字）；系統有值才填
     剩餘可供貨、Supply GIV、下單 GIV、COGS 含稅、永豐成本、品牌 SUMIF 都是底稿自己的公式：
     原料填對了，Excel 打開就會算出跟看板一樣的數字，系統不另外貼值蓋掉公式。
  3. 「系統看板」分頁：畫面上 ③ 的商品層與品牌層原樣輸出成值，給不想在一百欄裡找數字的人看。
  4. 「系統填入說明」分頁：這次填了哪些欄、幾格；沒填的為什麼；對不到的商品；自動新增的日期欄。
沒有底稿的線別（瑪氏、紙潔）走系統格式：總表（每日箱數）＋系統看板＋每日合計。
"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來
from .products import _template_row
from .summary import _build_summary, _ensure_date_columns, _row_index, _rewrite_month_total, _write_cases
from .sources import _h, _RE_SHEET_MONTH, _MON_EN, brand_block_cols
from .board import build_board

_RE_EOM_DAYS = re.compile(r"^到\d{1,2}月底庫存天數$")


def _norm_headers(ws, hdr):
    """底稿標題列 → {欄號: 正規化表頭}（小寫、去空白換行）。插欄之後要重算。"""
    return {c: _h(ws.cell(row=hdr, column=c).value) for c in range(1, ws.max_column + 1)}


def _value_cols(headers, mm):
    """系統要填值的欄：欄位名 → 欄號。找不到的就不填。"""
    cols = {}
    for c, h in headers.items():
        if h == "giv":
            cols.setdefault("giv", c)
        elif h == "niv":
            cols.setdefault("niv", c)
        elif h == "目前庫存數量":
            cols.setdefault("stock_remaining", c)
        elif h == "目前庫存天數":
            cols.setdefault("stock_days", c)
        elif _RE_EOM_DAYS.match(h):
            cols.setdefault("stock_days_eom", c)
        else:
            m = _RE_SHEET_MONTH.match(h)
            if m and _MON_EN.get(m.group(2)) == mm:
                cols.setdefault("supply" if m.group(1) == "supply_cs" else "demand", c)
    return cols


def _fill_values(ws, hdr, by_bc, by_sku, board, mm):
    """GIV／NIV、該月 Supply／demand、庫存三欄：對得到的商品列、系統有值的才填。回傳每個欄填了幾格與說明。"""
    headers = _norm_headers(ws, hdr)
    cols = _value_cols(headers, mm)
    n = collections.Counter()
    has_stock = board["totals"]["has_stock"]
    for r0 in board["rows"]:
        r = by_bc.get(r0["barcode"]) or by_sku.get(norm_key(r0["sku_id"]))
        if r is None:
            continue
        for k in ("giv", "niv", "supply", "demand"):
            c = cols.get(k)
            if c and r0.get(k) is not None:
                ws.cell(row=r, column=c).value = r0[k]; n[k] += 1
        if has_stock and r0["stock_remaining"] is not None:      # 有庫存資料的商品：三格一起填，算不出的留空（跟畫面一樣顯示空）
            for k in ("stock_remaining", "stock_days", "stock_days_eom"):
                c = cols.get(k)
                if c:
                    cell = ws.cell(row=r, column=c); cell.value = r0[k]; n[k] += 1
                    if k != "stock_remaining":
                        cell.number_format = "0.0"
    renamed = []
    c = cols.get("stock_days_eom")
    if c and has_stock and n["stock_remaining"]:
        want = f"到{mm}月底庫存天數"
        if norm_text(ws.cell(row=hdr, column=c).value) != want:
            renamed.append(f"「{norm_text(ws.cell(row=hdr, column=c).value)}」→「{want}」")
            ws.cell(row=hdr, column=c).value = want
    return {"cols": cols, "counts": dict(n), "renamed": renamed}


def _fill_brand_targets(ws, hdr, board):
    """品牌區（總表 BY～CU）的 目標／REBATE目標：系統有值的品牌才填，底稿沒列的品牌回報不硬加。"""
    headers = _norm_headers(ws, hdr)
    hlist = [headers.get(c, "") for c in range(1, ws.max_column + 1)]
    blk = brand_block_cols(hlist)
    if not blk:
        return {"found": False, "filled": 0, "not_in_sheet": []}
    i_key, i_t, i_r = blk
    key_col = i_key + 1; t_col = i_t + 1 if i_t is not None else None; r_col = i_r + 1 if i_r is not None else None
    rows_by_brand = {}
    for r in range(hdr + 1, ws.max_row + 1):
        k = norm_text(ws.cell(row=r, column=key_col).value)
        if k and k.lower() not in rows_by_brand:
            rows_by_brand[k.lower()] = r
    filled = 0; not_in_sheet = []
    for b in board["brands"]:
        if b["target_giv"] is None and b["rebate_target"] is None:
            continue
        r = rows_by_brand.get(norm_text(b["brand"]).lower())
        if r is None:
            not_in_sheet.append(b["brand"]); continue
        if t_col and b["target_giv"] is not None:
            ws.cell(row=r, column=t_col).value = b["target_giv"]; filled += 1
        if r_col and b["rebate_target"] is not None:
            ws.cell(row=r, column=r_col).value = b["rebate_target"]; filled += 1
    return {"found": True, "filled": filled, "not_in_sheet": not_in_sheet}


# ── 系統看板分頁 ──────────────────────────────────────────────────────────────
_G = {"prod": "EEF2F7", "price": "E8EEFB", "supply": "EEF7EE", "stock": "FBF3E6", "money": "F3EEFB", "q": "FDF2F8", "sub": "F8FAFC", "tot": "F1F5F9"}
_FMT_MONEY = "#,##0"; _FMT_CASES = "#,##0.##;-#,##0.##;\"-\""; _FMT_DAYS = "0.0"; _FMT_PCT = '0"%"'


def _board_sheet(wb, board, operator):
    """把 ③ 看板的品牌層、商品層照畫面的分組輸出成值。"""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L
    ws = wb.create_sheet("系統看板")
    mm = int(board["month"][5:7]); t = board["totals"]; ql = board["quarter"]["label"]
    thin = Side(style="thin", color="E2E8F0"); box = Border(left=thin, right=thin, top=thin, bottom=thin)
    bold = Font(bold=True); red = Font(color="B91C1C", bold=True); muted = Font(color="64748B", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = lambda hexv: PatternFill("solid", fgColor=hexv)   # noqa: E731

    def group_row(r, groups):
        """groups = [(標題, 欄數, 色)]，合併格。回傳下一欄號。"""
        c = 1
        for title, span, color in groups:
            ws.cell(row=r, column=c, value=title).fill = fill(color)
            ws.cell(row=r, column=c).font = bold; ws.cell(row=r, column=c).alignment = center
            if span > 1:
                ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + span - 1)
            for cc in range(c, c + span):
                ws.cell(row=r, column=cc).border = box
            c += span

    def head_row(r, names):
        for i, name in enumerate(names, start=1):
            cell = ws.cell(row=r, column=i, value=name); cell.font = bold; cell.fill = fill(_G["sub"]); cell.alignment = center; cell.border = box

    def put(r, c, v, fmt=None, font=None, bg=None):
        cell = ws.cell(row=r, column=c, value=v); cell.border = box
        if fmt and isinstance(v, (int, float)):
            cell.number_format = fmt
        if font:
            cell.font = font
        if bg:
            cell.fill = fill(bg)
        return cell

    ws.cell(row=1, column=1, value=f"{board['line']} {board['month'][:4]} 年 {mm} 月 看板").font = Font(bold=True, size=14)
    ws.cell(row=2, column=1, value=f"匯出 {now()} · {operator} · 數字跟畫面「③ 總表」一樣；公式寫在「系統填入說明」分頁").font = muted
    kpi = (f"{mm} 月下單 {t['ttl']:,.2f} 箱" + (f"（Supply {t['supply']:,.0f} 箱，已下 {round(t['ttl'] / t['supply'] * 100) if t['supply'] else 0}%）" if t["has_supply"] else "（supply 表還沒上傳）")
           + f" · 下單金額 GIV {t['ttl_giv']:,.0f} · COGS 含稅 {t['cogs_tax']:,.0f} · 永豐成本未稅 {t['yf_cost']:,.0f}"
           + (f" · 剩餘庫存 {t['stock_remaining']:,.0f} 箱、庫金 {t['stock_giv']:,.0f}" if t["has_stock"] else " · 庫存銷售表還沒上傳")
           + f" · 要注意 {t['alerts']} 個商品")
    ws.cell(row=3, column=1, value=kpi).font = Font(size=10)

    # ── 品牌 ──
    r = 5
    ws.cell(row=r, column=1, value="品牌").font = Font(bold=True, size=12); r += 1
    group_row(r, [("品牌", 3, _G["prod"]), (f"{mm} 月", 4, _G["money"]), (f"{ql}累計（目標是一季的）", 8, _G["q"]), ("庫存", 2, _G["stock"])]); r += 1
    head_row(r, ["品牌", "品類", "商品數", "Supply GIV", "下單 GIV", "COGS 含稅", "永豐成本未稅",
                 "Supply GIV", "下單 GIV", "目標", "達成", "差距（目標−下單）", "COGS 含稅", "REBATE 目標", "DIFF（目標−COGS）", "剩餘箱", "庫金"]); r += 1
    b_first = r
    for b in board["brands"]:
        put(r, 1, b["brand"], font=bold); put(r, 2, b["category"] or ""); put(r, 3, b["products"])
        put(r, 4, b["supply_giv"], _FMT_MONEY); put(r, 5, b["ttl_giv"], _FMT_MONEY); put(r, 6, b["cogs_tax"], _FMT_MONEY); put(r, 7, b["yf_cost"], _FMT_MONEY)
        put(r, 8, b["q_supply_giv"], _FMT_MONEY); put(r, 9, b["q_ttl_giv"], _FMT_MONEY)
        put(r, 10, b["target_giv"], _FMT_MONEY); put(r, 11, b["target_pct"], _FMT_PCT, font=red if (b["target_pct"] or 0) > 100 else None)
        put(r, 12, b["target_diff"], _FMT_MONEY, font=red if (b["target_diff"] or 0) < 0 else None)
        put(r, 13, b["q_cogs_tax"], _FMT_MONEY); put(r, 14, b["rebate_target"], _FMT_MONEY)
        put(r, 15, b["rebate_diff"], _FMT_MONEY, font=red if (b["rebate_diff"] or 0) < 0 else None)
        put(r, 16, b["stock_remaining"] or None, _FMT_CASES); put(r, 17, b["stock_giv"] or None, _FMT_MONEY)
        r += 1
    vals = ["合計", "", t["products"], t["supply_giv"], t["ttl_giv"], t["cogs_tax"], t["yf_cost"], t["q_supply_giv"], t["q_ttl_giv"],
            t["target_giv"], (round(t["q_ttl_giv"] / t["target_giv"] * 100) if t["target_giv"] else None),
            (round(t["target_giv"] - t["q_ttl_giv"], 2) if t["target_giv"] else None), t["q_cogs_tax"], t["rebate_target"],
            (round(t["rebate_target"] - t["q_cogs_tax"], 2) if t["rebate_target"] else None),
            t["stock_remaining"] if t["has_stock"] else None, t["stock_giv"] if t["has_stock"] else None]
    for i, v in enumerate(vals, start=1):
        put(r, i, v, _FMT_PCT if i == 11 else (_FMT_CASES if i == 16 else _FMT_MONEY), font=bold, bg=_G["tot"])
    r += 3

    # ── 商品 ──
    ws.cell(row=r, column=1, value="商品").font = Font(bold=True, size=12); r += 1
    group_row(r, [("商品", 7, _G["prod"]), ("單價", 3, _G["price"]), (f"{mm} 月供需（箱）", 4, _G["supply"]), ("庫存（箱）", 3, _G["stock"]), (f"{mm} 月金額", 4, _G["money"]), ("", 1, _G["prod"])]); r += 1
    head_row(r, ["品名", "條碼", "SKU ID", "品類", "品牌", "Note", "箱入數", "COGS", "GIV", "NIV",
                 "Supply", "demand", "下單", "剩餘可供貨", "剩餘", "天數", "到月底天數", "Supply GIV", "下單 GIV", "COGS 含稅", "永豐成本未稅", "注意"]); r += 1
    FLAG = {"over": "超打", "no_over": "不可超打卻超打", "low_stock": "庫存低", "no_box": "算不出箱數", "no_price": "缺 GIV"}
    for x in board["rows"]:
        put(r, 1, x["product_name"]); put(r, 2, x["barcode"]).number_format = "@"; put(r, 3, x["sku_id"]).number_format = "@"
        put(r, 4, x["category"]); put(r, 5, x["brand"]); put(r, 6, x["note"]); put(r, 7, x["box_size"], font=red if not x["box_size"] else None)
        put(r, 8, x["cost"], _FMT_CASES); put(r, 9, x["giv"], _FMT_CASES); put(r, 10, x["niv"], _FMT_CASES)
        put(r, 11, x["supply"], _FMT_CASES); put(r, 12, x["demand"], _FMT_CASES); put(r, 13, x["ttl"], _FMT_CASES, font=bold)
        put(r, 14, x["remaining_supply"], _FMT_CASES, font=red if (x["remaining_supply"] or 0) < 0 else None)
        put(r, 15, x["stock_remaining"], _FMT_CASES); put(r, 16, x["stock_days"], _FMT_DAYS, font=red if "low_stock" in x["flags"] else None); put(r, 17, x["stock_days_eom"], _FMT_DAYS)
        put(r, 18, x["supply_giv"], _FMT_MONEY); put(r, 19, x["ttl_giv"], _FMT_MONEY); put(r, 20, x["cogs_tax"], _FMT_MONEY); put(r, 21, x["yf_cost"], _FMT_MONEY)
        put(r, 22, "、".join(FLAG[f] for f in x["flags"]) or None, font=red if x["flags"] else None)
        r += 1
    tot = ["合計", f"{t['products']} 個商品", "", "", "", "", "", "", "", "", t["supply"] if t["has_supply"] else None, None, t["ttl"],
           (round(t["supply"] - t["ttl"], 2) if t["has_supply"] else None), t["stock_remaining"] if t["has_stock"] else None, None, None,
           t["supply_giv"], t["ttl_giv"], t["cogs_tax"], t["yf_cost"], ""]
    for i, v in enumerate(tot, start=1):
        put(r, i, v, _FMT_MONEY if i >= 18 else _FMT_CASES, font=bold, bg=_G["tot"])

    widths = {1: 42, 2: 17, 3: 17, 4: 8, 5: 14, 6: 12, 7: 7, 8: 8, 9: 10, 10: 10, 11: 9, 12: 9, 13: 9, 14: 11, 15: 9, 16: 8, 17: 10, 18: 13, 19: 13, 20: 13, 21: 13, 22: 16}
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w
    ws.freeze_panes = "B1"
    return ws


# ── 說明分頁、填底稿、匯出 ─────────────────────────────────────────────────────
def _write_report_sheet(wb, tpl, summary, month, operator, matched, filled, added, unmatched, info=None):
    """最後一個分頁「系統填入說明」：這次填了什麼、哪些欄沒填為什麼、自動新增了哪些日期欄、哪些商品總表沒有。"""
    from openpyxl.styles import Font
    sh = wb.create_sheet("系統填入說明")
    bold = Font(bold=True)
    sh.append(["商品主檔自動化 填入紀錄"]); sh["A1"].font = bold
    sh.append(["填入時間", now(), "操作者", operator])
    sh.append(["月份", month, "線別", summary["line"]])
    sh.append(["底稿", tpl["filename"], "工作表", tpl["sheet"]])
    sh.append(["對到的商品", matched, "填入箱數格數", filled])
    if added:
        sh.append(["自動新增的日期欄", "、".join(added)])
    if info:
        sh.append([])
        sh.append(["這次還填了哪些欄（用表頭認欄；系統有值才填，其他格一律不動）"]); sh.cell(row=sh.max_row, column=1).font = bold
        cnt = info["values"]["counts"]; cols = info["values"]["cols"]; mm = int(month[5:7])
        def line(label, keys, ok_note, none_note):
            n = sum(cnt.get(k, 0) for k in keys)
            if not any(k in cols for k in keys):
                sh.append([label, "底稿沒有這幾欄，跳過"])
            elif n:
                sh.append([label, f"填了 {n} 格", ok_note])
            else:
                sh.append([label, "沒填", none_note])
        line("GIV／NIV", ("giv", "niv"), "來源：商品主檔（Coupang Master 更新過的每箱進價）", "主檔裡沒有 GIV／NIV")
        line(f"Supply_CS／demand_CS ({mm}月)", ("supply", "demand"), "來源：寶僑 supply 表", "supply 表還沒上傳，維持底稿原本的內容")
        line("目前庫存數量／目前庫存天數／到月底庫存天數", ("stock_remaining", "stock_days", "stock_days_eom"),
             "系統照庫存銷售表公式現算：剩餘 = 累計下單 − 累計實銷；天數 = 剩餘 ÷ 日均實銷。底稿原本這三格的 VLOOKUP 指到別的欄，所以改填成值",
             "庫存銷售表還沒上傳，維持底稿原本的內容")
        if info["values"]["renamed"]:
            sh.append(["改了標題", "、".join(info["values"]["renamed"])])
        bt = info["targets"]
        if not bt["found"]:
            sh.append(["品牌目標／REBATE目標", "底稿沒有品牌區（目標／REBATE目標 欄），跳過"])
        elif bt["filled"]:
            sh.append(["品牌目標／REBATE目標", f"填了 {bt['filled']} 格", "一季一個數字（系統「③ 總表」品牌那頁填的）"])
        else:
            sh.append(["品牌目標／REBATE目標", "沒填", "系統裡這一季還沒有人填目標，維持底稿原本的數字"])
        if bt["not_in_sheet"]:
            sh.append(["系統有目標、但底稿品牌區沒列的品牌", "、".join(bt["not_in_sheet"])])
        sh.append([])
        sh.append(["剩餘可供貨、Supply GIV、下單 GIV、COGS 含稅、永豐成本、品牌 SUMIF 是底稿自己的公式，原料填對了 Excel 會自己算；系統不貼值蓋公式。"])
        sh.append(["「系統看板」分頁是畫面上 ③ 總表的原樣（值），數字應該跟 Sheet1 算出來的一樣；對不起來請回報。"])
    sh.append([])
    if unmatched:
        sh.append(["酷澎有下單、但總表裡沒有的商品（沒混進總表，列在這裡給你看）："]); sh.cell(row=sh.max_row, column=1).font = bold
        sh.append(["國條", "SKU ID", "品名", "品牌", "箱入數"] + [_md(d) for d in summary["dates"]] + ["月加總"])
        for c in sh[sh.max_row]:
            c.font = bold
        for r0 in unmatched:
            sh.append([r0["barcode"], r0["sku_id"], r0["product_name"], r0["brand"], r0["box_size"]]
                      + [r0["by_date"].get(d) for d in summary["dates"]] + [r0["month_total"]])
    sh.column_dimensions["A"].width = 30; sh.column_dimensions["B"].width = 26; sh.column_dimensions["C"].width = 70


def _fill_template(tpl, summary, month, operator, board=None):
    """把這個月的數字填進底稿：只動這個月的日期欄、系統有值的原料欄、對得到的商品列，其他列、其他月份、
    業務的公式、順序全部不碰。步驟：確保日期欄都在（缺的插欄）→ 對商品列 → 重寫 TTL → 填箱數
    → 填 GIV／NIV／供需／庫存 → 填品牌目標 → 系統看板 → 說明分頁。board 不給就只填箱數（測試用）。"""
    import base64
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(tpl["content_b64"])))   # 保留公式
    ws = wb[tpl["sheet"]]
    hdr = tpl["header_row"]; mm = int(month[5:7])
    lay, added = _ensure_date_columns(ws, hdr, mm, summary["dates"])
    by_bc, by_sku = _row_index(ws, hdr, lay["bc_col"], lay["sku_col"])
    block = sorted(lay["date_cols"].values()) + lay["spare"]
    _rewrite_month_total(ws, hdr, lay, block)
    matched, unmatched, filled = _write_cases(ws, lay, block, by_bc, by_sku, summary)
    info = None
    if board is not None:
        info = {"values": _fill_values(ws, hdr, by_bc, by_sku, board, mm), "targets": _fill_brand_targets(ws, hdr, board)}
        _board_sheet(wb, board, operator)
    _write_report_sheet(wb, tpl, summary, month, operator, matched, filled, added, unmatched, info)
    return wb, {"matched": matched, "filled": filled, "unmatched": len(unmatched),
                "missing_dates": [], "renamed": added, "info": info}


@master_bp.route("/api/master/export")
def api_export():
    """總表：有底稿照底稿填（寶僑），沒有就系統格式。兩種都附「系統看板」分頁。"""
    group = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not group or not _valid_month(month):
        return jsonify({"error": "請選線別與月份。"}), 400
    cfg = _line_groups()
    conn = get_conn()
    try:
        s = _build_summary(conn, group, month, cfg)
        board = build_board(conn, group, month, cfg)
        orders = [o for o in _month_orders(conn, month, cfg) if o["line_group"] == group]
        tpl = _template_row(conn, group)
    finally:
        conn.close()
    operator = _operator()
    if tpl is not None:
        # 有底稿就照底稿填（寶僑要跟他們的總表一模一樣，含商品順序），檔名沿用底稿的
        wb, _rep = _fill_template(tpl, s, month, operator, board)
        # 檔名就叫「總表.xlsx」（Jerry 2026-09-18 定案）：不能跟底稿同名（Excel 無法同時開兩個同名活頁簿），
        # 也不放月份——匯出的是整份總表，只是某個月的箱數填好了，叫「9月_總表」會誤導成只有 9 月。
        return _xlsx_response(wb, "總表.xlsx")
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "總表"
    mm = int(month[5:7])
    head = ["國條", "SKU ID", "永豐料號", "品類", "品牌", "品名", "箱入數"] + [_md(d) for d in s["dates"]] + \
           [f"{mm}月TTL下單總箱數", "備註"]
    ws.append(head)
    for r in s["rows"]:
        ws.append([r["barcode"], r["sku_id"], r["yf_sku"], r["category"], r["brand"], r["product_name"], r["box_size"]]
                  + [r["by_date"].get(d) for d in s["dates"]] + [r["month_total"], r["note"]])
    ws.append([])
    ws.append(["合計", "", "", "", "", "", ""] + [s["totals_by_date"].get(d) for d in s["dates"]] + [s["month_total"], ""])
    bold = Font(bold=True); fill = PatternFill("solid", fgColor="DBEAFE")
    for c in ws[1]:
        c.font = bold; c.fill = fill; c.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in ws[ws.max_row]:
        c.font = bold
    ws.freeze_panes = "H2"
    for i in range(1, len(head) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 40 if i == 6 else 14
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
        for c in row:
            c.number_format = "@"
    _board_sheet(wb, board, operator)
    ws2 = wb.create_sheet("每日合計"); ws2.append(["交貨日", "出貨總箱數", "PO 數", "品項數"])
    per_day = collections.OrderedDict()
    for o in orders:
        if not o["delivery_date"]:
            continue
        d = per_day.setdefault(o["delivery_date"], {"cases": 0.0, "pos": set(), "rows": 0})
        d["cases"] += o["cases"] or 0; d["pos"].add(o["po_number"]); d["rows"] += 1
    for dte, d in per_day.items():
        ws2.append([dte, round(d["cases"], 2), len(d["pos"]), d["rows"]])
    for c in ws2[1]:
        c.font = bold; c.fill = fill
    return _xlsx_response(wb, "總表.xlsx")


__all__ = ["api_export", "_fill_template"]
