/* 商品主檔自動化 前端（templates/master.html 載入）。
   結構：state（畫面狀態）→ loadMeta/loadOrders/loadProducts/loadSummary 撈 API → render* 畫畫面。
   所有請求走 api()，409 表示別人先改了會自動重載。改了這個檔要升 app.py 的 BUILD_VERSION，
   網址帶 ?v=版本 才不會被瀏覽器快取住舊的。 */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const fmt = n => n == null ? "" : (Number.isInteger(n) ? n : Number(n).toFixed(2).replace(/\.?0+$/, ""));
const state = { month: "", tab: "orders", lines: new Set(), dates: new Set(), pos: new Set(), brands: new Set(), warehouses: new Set(),
                edited: false, missing: false, q: "", facets: null, rows: [], selected: new Set(), groups: [], sumLine: "", sumMonth: "",
                batch: null, calView: localStorage.getItem("mst_calview") || "cal", dense: localStorage.getItem("mst_dense") === "1", prevMonth: null, prevTo: null,
                closedDates: new Set(), openPos: new Set(), hlDate: null };

function toast(msg, kind="ok") {
  const t = $("#toast"); t.textContent = msg; t.className = "t-" + kind;
  void t.offsetWidth; t.classList.add("show");
  clearTimeout(toast._h); toast._h = setTimeout(() => t.classList.remove("show"), kind === "err" ? 6500 : 3500);
}
async function api(url, opts={}) {
  const res = await fetch(url, opts);
  let data = {}; try { data = await res.json(); } catch (e) {}
  if (!res.ok) {
    const msg = data.message || data.error || `伺服器錯誤 (${res.status})`;
    if (res.status === 409) toast(msg + "　（重新載入中…）", "err");
    throw Object.assign(new Error(msg), { status: res.status, data });
  }
  return data;
}
const J = body => ({ method: "PUT", headers: {"Content-Type":"application/json"}, body: JSON.stringify(body) });
document.querySelectorAll("dialog [data-close]").forEach(b => b.addEventListener("click", e => e.target.closest("dialog").close()));
document.querySelectorAll("dialog").forEach(d => d.addEventListener("click", e => { if (e.target === d) d.close(); }));
function dropzone(el, input, onFiles) {
  // 一次拖好幾個檔也全部收（例如 7 月＋8 月兩份彙總表），交給 onFiles 一整串處理
  el.addEventListener("click", () => input.click());
  el.addEventListener("dragover", e => { e.preventDefault(); el.classList.add("dz-active"); });
  el.addEventListener("dragleave", () => el.classList.remove("dz-active"));
  el.addEventListener("drop", e => { e.preventDefault(); el.classList.remove("dz-active"); const fs = [...e.dataTransfer.files]; if (fs.length) onFiles(fs); });
  input.addEventListener("change", e => { const fs = [...e.target.files]; if (fs.length) onFiles(fs); input.value = ""; });
}
window.addEventListener("dragover", e => e.preventDefault()); window.addEventListener("drop", e => e.preventDefault());

/* ── 分頁 ── */
function setTab(name) {
  state.tab = name;
  document.querySelectorAll(".m-tab").forEach(x => x.classList.toggle("active", x.dataset.tab === name));
  ["orders","summary","products"].forEach(t => $("#tab-"+t).classList.toggle("hidden", t !== name));
}
document.querySelectorAll(".m-tab").forEach(b => b.addEventListener("click", () => { setTab(b.dataset.tab); refresh(); }));
async function loadMeta() {
  const d = await api("/api/master/lines");
  state.groups = d.groups;
  if (!state.month) state.month = localStorage.getItem("mst_month") || d.months[0] || d.this_month;
  $("#sel-month").value = state.month; if (!$("#exp-to").value || $("#exp-to").value < state.month) $("#exp-to").value = state.month;
  loadImportScopes();
  const sl = $("#sum-line"); const cur = state.sumLine || localStorage.getItem("mst_sum_line") || d.groups[0] || "";
  sl.innerHTML = d.groups.map(g => `<option ${g === cur ? "selected" : ""}>${esc(g)}</option>`).join("") || `<option value="">（還沒有訂單）</option>`;
  state.sumLine = sl.value; state.sumMonth = state.sumMonth || state.month; $("#sum-month").value = state.sumMonth; if (!$("#sum-exp-to").value || $("#sum-exp-to").value < state.sumMonth) $("#sum-exp-to").value = state.sumMonth;
  const pl = $("#prod-line"); const pv = pl.value;
  pl.innerHTML = `<option value="">全部線別</option>` + d.groups.map(g => `<option ${g === pv ? "selected" : ""}>${esc(g)}</option>`).join("");
}
async function checkMaster() {
  try { const d = await api("/api/master/products"); const empty = !d.products.length; $("#no-master-banner").classList.toggle("hidden", !empty); return empty; }
  catch (e) { return false; }
}
function refresh() {
  checkMaster();
  if (state.tab === "orders") loadOrders(); else if (state.tab === "summary") loadSummary(); else loadProducts();
}

/* ════════════ ② 訂單明細 ════════════ */
function shiftMonth(delta) {
  const [y, m] = state.month.split("-").map(Number); const d = new Date(y, m - 1 + delta, 1);
  setMonth(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`);
}
function setMonth(v) { if (!v) return; state.month = v; localStorage.setItem("mst_month", v); $("#sel-month").value = v; $("#exp-to").value = v; state.dates.clear(); state.pos.clear(); state.selected.clear(); loadOrders(); }
$("#m-prev").addEventListener("click", () => shiftMonth(-1)); $("#m-next").addEventListener("click", () => shiftMonth(1));
$("#sel-month").addEventListener("change", e => setMonth(e.target.value));
/* 月份藥丸：平常寫「2026 年 9 月」，拉了「到」就寫「2026 年 9 月 ～ 10 月」；點開才看到起訖兩格 */
function renderMonthPill() {
  const to = $("#exp-to").value; const [y, m] = state.month.split("-").map(Number);
  $("#month-pill-txt").textContent = (to && to > state.month) ? `${y} 年 ${m} 月 ～ ${Number(to.slice(5))} 月` : `${y} 年 ${m} 月`;
  $("#month-pill").classList.toggle("on", !!(to && to > state.month));
}
$("#month-pill").addEventListener("click", e => { e.stopPropagation(); document.querySelectorAll(".dd-menu").forEach(x => { if (x.id !== "month-menu") x.classList.add("hidden"); }); $("#month-menu").classList.toggle("hidden"); });
$("#month-one").addEventListener("click", () => { $("#exp-to").value = state.month; $("#month-menu").classList.add("hidden"); loadOrders(); });
$("#month-two").addEventListener("click", () => { const [y, m] = state.month.split("-").map(Number); const d = new Date(y, m, 1); $("#exp-to").value = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`; $("#month-menu").classList.add("hidden"); loadOrders(); });
function clearFilters() {
  state.lines.clear(); state.dates.clear(); state.pos.clear(); state.brands.clear(); state.warehouses.clear(); state.edited = state.missing = false; state.q = ""; state.batch = null;
  $("#q").value = ""; $("#chip-edited").classList.remove("on"); $("#chip-missing").classList.remove("on"); loadOrders();
}
$("#btn-clear").addEventListener("click", clearFilters);
$("#q").addEventListener("input", e => { state.q = e.target.value.trim(); clearTimeout(window._qt); window._qt = setTimeout(loadOrders, 250); });
$("#chip-edited").addEventListener("click", () => { state.edited = !state.edited; $("#chip-edited").classList.toggle("on", state.edited); loadOrders(); });
$("#chip-missing").addEventListener("click", () => { state.missing = !state.missing; $("#chip-missing").classList.toggle("on", state.missing); loadOrders(); });
function filterParams() {
  const to = $("#exp-to").value; const month_to = (to && to > state.month) ? to : state.month;
  return new URLSearchParams({ month: state.month, month_to, q: state.q, lines: [...state.lines].join(","), dates: [...state.dates].join(","),
    pos: [...state.pos].join(","), brands: [...state.brands].join(","), warehouses: [...state.warehouses].join(","),
    edited: state.edited ? "1" : "", missing: state.missing ? "1" : "", batch: state.batch || "", batch_scope: "all" });
}
$("#exp-to").addEventListener("change", () => { loadOrders(); });
function anyFilter() { return state.lines.size || state.dates.size || state.pos.size || state.brands.size || state.warehouses.size || state.q || state.edited || state.missing; }
function anyScope() { return anyFilter() || state.batch; }

async function loadOrders(opts = {}) {
  $("#btn-export-daily").href = "/api/master/export/daily?" + filterParams(); renderMonthPill();
  $("#filter-hint").classList.toggle("hidden", !anyFilter()); $("#btn-clear").classList.toggle("hidden", !anyFilter());
  let data;
  try { data = await api("/api/master/orders?" + filterParams()); } catch (e) { toast(e.message, "err"); return; }
  state.facets = data.facets; state.rows = data.rows;
  const curBatch = state.batch ? importBatches.find(b => b.id === state.batch) : null;
  $("#btn-export-daily").innerHTML = curBatch ? `<i class="bi bi-calendar-week"></i> 匯出這批的專案報價檔（${new Set(state.rows.map(r => r.po_number)).size} 張 PO）` : `<i class="bi bi-calendar-week"></i> 匯出專案報價檔`;
  $("#btn-export-daily").classList.toggle("btn-p", true);
  renderImportBar();
  // 拖選日期時行事曆／清單已經先變色了，不再重畫它，畫面才不會頓、閃
  if (!opts.keepCal) renderCalendar(); else $("#cal-clear-dates").classList.toggle("hidden", !state.dates.size);
  renderDropdowns(); renderOrders(); renderBatch(); renderWarehouses(); renderAlerts();
}
function renderAlerts() {
  const fx = state.facets || {};
  $("#chip-missing").classList.toggle("hidden", !(fx.missing_box_rows || state.missing)); $("#chip-missing").innerHTML = `<i class="bi bi-exclamation-triangle-fill"></i> ${fx.missing_box_rows || 0} 筆算不出箱數`;
  $("#chip-edited").classList.toggle("hidden", !(fx.edited_rows || state.edited)); $("#chip-edited").innerHTML = `<i class="bi bi-pencil"></i> ${fx.edited_rows || 0} 筆人工調整過`;
}

/* 倉別箱數：行事曆下面一塊。沒勾日期＝整個月各倉別合計；勾了幾天＝這幾天的合計，
   勾兩天以上再多一塊「每一天 × 倉別」。跟其他篩選（線別、PO、品牌…）連動。 */
function renderWarehouses() {
  const rows = state.rows || []; const box = $("#wh-block");
  if (!rows.length) { box.innerHTML = ""; return; }
  const agg = new Map(); const byDay = new Map(); const days = new Set();
  for (const r of rows) {
    const w = r.warehouse || "（未填倉別）"; const d = r.delivery_date || "未排日期"; days.add(d);
    const a = agg.get(w) || { cases: 0, pos: new Set(), rows: 0, missing: 0 }; a.rows++; a.pos.add(r.po_number);
    if (r.cases == null) a.missing++; else a.cases += r.cases; agg.set(w, a);
    const k = d + "|" + w; const b = byDay.get(k) || { cases: 0, pos: new Set() }; if (r.cases != null) b.cases += r.cases; b.pos.add(r.po_number); byDay.set(k, b);
  }
  const whs = [...agg.keys()].sort((a, b) => agg.get(b).cases - agg.get(a).cases);   // 箱數多的排前面
  const total = [...agg.values()].reduce((s, a) => s + a.cases, 0);
  const to = $("#exp-to").value;
  const scope = state.dates.size ? `已勾 ${state.dates.size} 天` : (to && to > state.month) ? `${Number(state.month.slice(5))}～${Number(to.slice(5))} 月` : `${Number(state.month.slice(5))} 月整月`;
  // 一個倉別一張卡、橫向排滿整個寬度；最前面一張是合計
  const card = (title, cases, sub, extra = "", style = "") => `<div class="stat" style="min-width:0;${style}"><span style="font-weight:700;color:var(--ink)">${title}</span><b style="white-space:nowrap">${fmt(cases)}${extra}</b><span>${sub}</span></div>`;
  let html = `<div class="flex items-center gap-2 mb-2"><b class="text-sm" style="color:var(--b900)"><i class="bi bi-building"></i> 倉別箱數</b><span class="kbd">${esc(scope)}${anyFilter() && !state.dates.size ? "（依目前篩選）" : ""} · 箱數多的排前面</span></div>
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:8px">
    ${card("全部倉別合計", total, `${new Set(rows.map(r => r.po_number)).size} 張 PO · ${rows.length} 品項`, "", "background:var(--b50);border-color:var(--b200)")}
    ${whs.map(w => { const a = agg.get(w); return card(esc(w), a.cases, `${a.pos.size} 張 PO · ${a.rows} 品項`, a.missing ? ` <span class="neg" style="font-size:13px" title="有 ${a.missing} 個品項算不出箱數，沒算進去">!${a.missing}</span>` : ""); }).join("")}
    </div>`;
  if (state.dates.size >= 2) {
    const dl = [...days].sort();
    html += `<div class="tbl-wrap mt-3"><table class="m"><thead><tr><th>日期</th>${whs.map(w => `<th class="num">${esc(w)}</th>`).join("")}<th class="num">當日合計</th></tr></thead><tbody>
      ${dl.map(d => { let t = 0; const cells = whs.map(w => { const b = byDay.get(d + "|" + w); if (!b) return `<td class="num muted">—</td>`; t += b.cases; return `<td class="num" title="${b.pos.size} 張 PO">${fmt(b.cases)}</td>`; }).join(""); return `<tr><td class="mono"><b>${esc(d)}</b></td>${cells}<td class="num"><b>${fmt(t)}</b></td></tr>`; }).join("")}
      <tr style="background:var(--b50)"><td><b>合計</b></td>${whs.map(w => `<td class="num"><b>${fmt(agg.get(w).cases)}</b></td>`).join("")}<td class="num"><b>${fmt(total)}</b></td></tr></tbody></table></div>`;
  }
  box.innerHTML = html;
}

/* 「最近一次匯入」那一行就是開關：全部訂單 ／ 只看這次匯入的變動。切到變動模式時畫面只剩
   那次新增＋有變的列（新增綠標、有變黃標寫出改了什麼），匯出按鈕也變成「匯出這次變動的專案報價檔」。
   要看更早的某一次，點「更早的匯入」再選。 */
let importBatches = [];
async function loadImportScopes() {
  let d; try { d = await api("/api/master/imports?limit=300"); } catch (e) { return; }
  importBatches = d.batches; renderImportBar();
}
/* 時間寫成人話：今天／昨天只留時刻，其他寫 月-日 時:分 */
function whenText(ts) {
  if (!ts) return "";
  const d = ts.slice(0, 10), hm = ts.slice(11, 16);
  const today = new Date(), y = new Date(Date.now() - 86400000), iso = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`;
  return d === iso(today) ? `今天 ${hm}` : d === iso(y) ? `昨天 ${hm}` : `${d.slice(5)} ${hm}`;
}
/* 數字以 PO 張數為主、品項數當小字（OP 的單位是 PO）；0 的灰掉 */
const countBadges = b => `<span class="badge ${b.new_pos ? "bd-ok" : "z"}">新增 ${b.new_pos} 張 PO${b.new_now ? ` <span style="font-weight:500;opacity:.8">· ${b.new_now} 品項</span>` : ""}</span><span class="badge ${b.changed_pos ? "bd-warn" : "z"}">有變 ${b.changed_pos}${b.changed_pos ? ` 張 PO <span style="font-weight:500;opacity:.8">· ${b.changed_now} 品項</span>` : ""}</span><span class="badge ${b.removed_pos ? "bd-bad" : "z"}">消失 ${b.removed_pos}${b.removed_pos ? ` 張 PO <span style="font-weight:500;opacity:.8">· ${b.removed_now} 品項</span>` : ""}</span>`;
const spanText = b => { const ms = b.months || []; return ms.length > 1 ? `交期跨 ${Number(ms[0].slice(5))}～${Number(ms[ms.length - 1].slice(5))} 月，已一次列出` : ""; };
/* 上面四張數字卡：今天要出貨／本週要出貨／最近一批匯入／需要處理。
   點卡就是動作：今天→勾今天、本週→勾這週有出貨的天、最近一批→拉開抽屜、需要處理→只看那些。
   依匯入批次篩選中時第三張卡變黃，寫「依匯入批次篩選：哪一批」，上面有「清除批次篩選」。 */
function renderImportBar() {
  const li = $("#last-import"); const last = importBatches[0];
  const f = state.facets || {}; const dates = f.dates || [];
  if (!last && !dates.length) { li.classList.add("hidden"); return; }
  li.classList.remove("hidden");
  const cur = importBatches.find(b => b.id === state.batch) || null;
  const iso = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`;
  const now = new Date(); const today = iso(now);
  const wk = []; const start = new Date(now); start.setDate(now.getDate() - now.getDay());
  for (let i = 0; i < 7; i++) { const d = new Date(start); d.setDate(start.getDate() + i); wk.push(iso(d)); }
  const inRange = d => d.slice(0, 7) >= state.month && d.slice(0, 7) <= ($("#exp-to").value || state.month);
  const sum = keys => dates.filter(d => keys.includes(d.date)).reduce((a, d) => ({ pos: a.pos + d.po_count, rows: a.rows + d.rows, cases: a.cases + d.cases }), { pos: 0, rows: 0, cases: 0 });
  const tdy = sum([today]); const wkS = sum(wk); const wkDays = wk.filter(k => dates.some(d => d.date === k));
  const need = (f.missing_box_rows || 0) + (f.edited_rows || 0);
  const todayOn = state.dates.size === 1 && state.dates.has(today); const weekOn = wkDays.length > 0 && wkDays.every(k => state.dates.has(k)) && state.dates.size === wkDays.length;
  const card1 = inRange(today)
    ? `<div class="kpi ${todayOn ? "on" : ""}" data-k="today" title="點一下只看今天"><div class="kt"><span>今天要出貨</span><i class="bi bi-truck"></i></div><div class="kn">${tdy.pos}<small>張 PO</small></div><div class="ks">${tdy.pos ? `${tdy.rows} 品項 · ${fmt(Math.round(tdy.cases * 100) / 100)} 箱` : "今天沒有出貨"}</div></div>`
    : `<div class="kpi" title="今天不在目前看的月份裡"><div class="kt"><span>今天要出貨</span><i class="bi bi-truck"></i></div><div class="kn muted">—</div><div class="ks">今天不在目前看的月份</div></div>`;
  const card2 = `<div class="kpi ${weekOn ? "on" : ""}" data-k="week" title="點一下只看這週"><div class="kt"><span>本週要出貨</span><i class="bi bi-calendar-week"></i></div><div class="kn">${wkS.pos}<small>張 PO</small></div><div class="ks">${wkS.pos ? `${wkS.rows} 品項 · ${fmt(Math.round(wkS.cases * 100) / 100)} 箱 · ${wkDays.length} 天` : "這週沒有出貨"}</div></div>`;
  const card3 = !last ? `<div class="kpi"><div class="kt"><span>最近匯入</span></div><div class="kn muted">—</div><div class="ks">還沒匯入過</div></div>`
    : cur ? `<div class="kpi batch" data-k="batch"><div class="kt"><span style="color:var(--warn);font-weight:700"><i class="bi bi-funnel-fill"></i> 依匯入批次篩選：${esc(whenText(cur.committed_at))} 這批</span><button class="btn btn-o btn-sm" id="btn-batch-clear"><i class="bi bi-x-lg"></i> 清除批次篩選</button></div><div class="kn">${cur.new_pos}<small>張 PO 新增</small></div><div class="ks">${countBadges(cur)}${spanText(cur) ? `<span>${spanText(cur)}</span>` : ""}<button class="btn btn-g btn-sm" id="btn-batch-filter" style="margin-left:auto" title="換一批"><i class="bi bi-funnel"></i> 依匯入批次篩選</button></div></div>`
    : `<div class="kpi" data-k="batch" title="點一下依匯入批次篩選"><div class="kt"><span>最近匯入</span><span>${esc(whenText(last.committed_at))}</span></div><div class="kn">${last.new_pos}<small>張 PO 新增</small></div><div class="ks">${countBadges(last)}<button class="btn btn-o btn-sm" id="btn-batch-filter" style="margin-left:auto"><i class="bi bi-funnel"></i> 依匯入批次篩選</button></div></div>`;
  const needOn = state.missing || state.edited;
  const card4 = `<div class="kpi ${need ? "warn" : ""} ${needOn ? "on" : ""}" data-k="need" title="${need ? "點一下只看要處理的品項" : "沒有要處理的"}"><div class="kt"><span>需要處理</span><i class="bi ${need ? "bi-exclamation-triangle-fill" : "bi-check-circle"}"></i></div><div class="kn">${need}<small>筆</small></div><div class="ks">${need ? `${f.missing_box_rows ? `算不出箱數 ${f.missing_box_rows}` : ""}${f.missing_box_rows && f.edited_rows ? " · " : ""}${f.edited_rows ? `人工調整過 ${f.edited_rows}` : ""}` : "箱數都算得出來、沒有人工調整"}</div></div>`;
  li.innerHTML = `<div class="kpis">${card1}${card2}${card3}${card4}</div>`;
  li.querySelectorAll(".kpi[data-k]").forEach(k => k.addEventListener("click", e => {
    if (e.target.closest("button")) return;
    const kind = k.dataset.k;
    if (kind === "today") { const on = todayOn; state.dates.clear(); if (!on) state.dates.add(today); loadOrders(); }
    else if (kind === "week") { const on = weekOn; state.dates.clear(); if (!on) wkDays.forEach(d => state.dates.add(d)); loadOrders(); }
    else if (kind === "batch") { if (cur) return; openDrawer(); }
    else if (kind === "need") { if (!need) return; if (needOn) { state.missing = state.edited = false; } else { state.missing = !!f.missing_box_rows; state.edited = !f.missing_box_rows && !!f.edited_rows; } $("#chip-missing").classList.toggle("on", state.missing); $("#chip-edited").classList.toggle("on", state.edited); loadOrders(); }
  }));
  const c = li.querySelector("#btn-batch-clear"); if (c) c.addEventListener("click", () => pickBatch(null));
  const b = li.querySelector("#btn-batch-filter"); if (b) b.addEventListener("click", () => openDrawer());
  if ($("#imp-drawer").classList.contains("open")) renderDrawer();
}
/* 右邊抽屜：每一次匯入一列（時間、新增／有變／消失、檔名），最新在上；點一列就依那批篩選，抽屜不關 */
/* 抽屜上面一條篩選：期間（本週／本月／上個月）＋搜檔名或 PO 單號；匯一百次也能三秒找到那批 */
let batchPos = {};   // batch id → 那批動到的 PO（搜單號用，第一次搜才撈）
function drawerFiltered() {
  const period = $("#imp-period").value; const q = $("#imp-q").value.trim().toLowerCase();
  const now = new Date(); const iso = x => `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}-${String(x.getDate()).padStart(2, "0")}`;
  const wk = new Date(now); wk.setDate(now.getDate() - now.getDay());
  const thisM = iso(now).slice(0, 7); const lastM = iso(new Date(now.getFullYear(), now.getMonth() - 1, 1)).slice(0, 7);
  return importBatches.filter(b => {
    const d = (b.committed_at || "").slice(0, 10);
    if (period === "week" && d < iso(wk)) return false;
    if (period === "month" && d.slice(0, 7) !== thisM) return false;
    if (period === "last" && d.slice(0, 7) !== lastM) return false;
    if (q && !(b.filename || "").toLowerCase().includes(q) && !(batchPos[b.id] || []).some(po => po.includes(q))) return false;
    return true;
  });
}
function renderDrawer() {
  const list = drawerFiltered();
  $("#imp-count").textContent = list.length === importBatches.length ? `共 ${importBatches.length} 批` : `${list.length}／${importBatches.length} 批`;
  $("#imp-drawer-body").innerHTML = list.map(b => `<div class="imp-row ${b.id === state.batch ? "on" : ""}" data-id="${b.id}" title="${esc(b.filename)}">
      <span class="t">${b.id === state.batch ? `<i class="bi bi-check-lg"></i> ` : ""}${esc(whenText(b.committed_at))}</span><span class="b">${countBadges(b)}</span>
      <span class="f">${b === importBatches[0] ? "最近一批 · " : ""}${esc(b.filename)}</span></div>`).join("") || `<div class="muted p-4 text-center">${importBatches.length ? "沒有符合的匯入紀錄" : "還沒有匯入過"}</div>`;
  $("#imp-drawer-body").querySelectorAll(".imp-row").forEach(r => r.addEventListener("click", () => pickBatch(r.dataset.id)));
}
$("#imp-period").addEventListener("change", renderDrawer);
$("#imp-q").addEventListener("input", async () => {
  const q = $("#imp-q").value.trim();
  // 打的像 PO 單號（一串數字）就去撈每批動到的 PO，讓「貼一個單號找它在哪幾批」做得到
  if (/^\d{4,}$/.test(q) && !Object.keys(batchPos).length) {
    try { const d = await api("/api/master/imports/pos"); batchPos = d.pos || {}; } catch (e) { /* 撈不到就只搜檔名 */ }
  }
  renderDrawer();
});
function openDrawer() { renderDrawer(); const d = $("#imp-drawer"); d.classList.add("open"); d.setAttribute("aria-hidden", "false"); document.body.classList.add("drawer-open"); }
function closeDrawer() { const d = $("#imp-drawer"); d.classList.remove("open"); d.setAttribute("aria-hidden", "true"); document.body.classList.remove("drawer-open"); }
// 點抽屜以外任何地方就關（開抽屜那顆按鈕除外，不然一開就關）
// 用 click 不用 pointerdown：按下去就關的話主畫面會先往右展開，滑鼠放開時底下已經不是原本那格，那一下點擊就丟了
document.addEventListener("click", e => { if ($("#imp-drawer").classList.contains("open") && !e.target.closest("#imp-drawer") && !e.target.closest("#btn-batch-filter")) closeDrawer(); }, true);
$("#imp-drawer-close").addEventListener("click", closeDrawer);
document.addEventListener("keydown", e => { if (e.key === "Escape" && $("#imp-drawer").classList.contains("open")) closeDrawer(); });
/* 依某批篩選：月份自動拉成那批交期涵蓋的範圍（9 月～10 月），表格、匯出一次包完，不用逐月切；
   清除時回到原本停的月份。 */
function pickBatch(id) {
  const b = importBatches.find(x => x.id === Number(id));
  if (b) {
    if (state.batch == null) { state.prevMonth = state.month; state.prevTo = $("#exp-to").value; }
    state.batch = b.id;
    if (b.months && b.months.length) { state.month = b.months[0]; $("#sel-month").value = state.month; $("#exp-to").value = b.months[b.months.length - 1]; }
  } else {
    state.batch = null;
    if (state.prevMonth) { state.month = state.prevMonth; $("#sel-month").value = state.month; $("#exp-to").value = state.prevTo || state.month; }
    state.prevMonth = state.prevTo = null;
  }
  state.dates.clear(); state.pos.clear();
  loadOrders();
}

/* 每一天被「最近一批」（或篩選中的那批）動到幾張 PO：新增／有變／消失 */
function changesByDate() {
  const mark = state.batch || (importBatches[0] && importBatches[0].id);
  const chg = {};
  for (const r of state.rows || []) {
    const k = !mark ? "" : r.first_batch_id === mark ? "n" : r.last_batch_id === mark ? (r.missing_in_file ? "g" : "u") : "";
    if (!k) continue;
    const d = r.delivery_date || ""; chg[d] = chg[d] || { n: new Set(), u: new Set(), g: new Set(), rows: 0 }; chg[d][k].add(r.po_number); chg[d].rows++;
  }
  return chg;
}
const cornerHtml = c => c ? `<div class="dc" title="最近一批動到這天：${c.n.size ? `新增 ${c.n.size} 張 PO` : ""}${c.u.size ? ` 有變 ${c.u.size} 張 PO` : ""}${c.g.size ? ` 消失 ${c.g.size} 張 PO` : ""}（${c.rows} 品項）">${c.n.size ? `<span class="dn">新${c.n.size}</span>` : ""}${c.u.size ? `<span class="du">變${c.u.size}</span>` : ""}${c.g.size ? `<span class="dg">消${c.g.size}</span>` : ""}</div>` : "";
const dateKeys = () => (state.facets?.dates || []).map(d => d.date).filter(Boolean).sort();
function applyDates(keys, on) { for (const k of keys) on ? state.dates.add(k) : state.dates.delete(k); afterPick(); }

function renderCalendar() {
  const [y, m] = state.month.split("-").map(Number);
  const to = $("#exp-to").value; const spans = to && to > state.month;
  $("#cal-title").textContent = spans ? `${y} 年 ${m} 月～${Number(to.slice(5))} 月 出貨` : `${y} 年 ${m} 月 出貨日曆`;
  const chg = changesByDate();
  const curB = state.batch ? importBatches.find(b => b.id === state.batch) : null;
  $("#cal-clear-dates").classList.toggle("hidden", !state.dates.size); $("#cal-clear-dates").onclick = () => { state.dates.clear(); loadOrders(); };
  // 行事曆／日期清單同一組切換，全部訂單跟依匯入批次篩選都一樣；批次模式只是沒動到的天壓淡、點一天跳到那天
  const mode = state.calView;
  $("#cal").classList.toggle("hidden", mode !== "cal"); $("#cal-list").classList.toggle("hidden", mode !== "list");
  $("#cal-hint").textContent = curB ? (mode === "list" ? "這批動到的日期，一列一天。點一列選一天、按著往下拖選好幾天；只選一天會順便捲到那天。" : "亮的是這批動到的天、角標是幾張 PO。點一天、按著滑過幾天都是選日期；只選一天會順便捲到那天並亮一下。")
    : mode === "list" ? "點一列勾一天；按著往下拖一次勾好幾天；Shift 點兩列中間整段一起勾。" : "點一天選一天；按著滑過好幾天一起選；Shift 點兩天中間整段一起選。再點一次取消。";
  if (mode === "list") { renderDateList(chg, curB); renderNoDateChip(); return; }

  const byDate = Object.fromEntries((state.facets?.dates || []).map(d => [d.date, d]));
  const months = []; let yy = y, mm = m;
  while (`${yy}-${String(mm).padStart(2, "0")}` <= (spans ? to : state.month) && months.length < 12) { months.push([yy, mm]); mm++; if (mm > 12) { mm = 1; yy++; } }
  const grid = (Y, M) => {
    const first = new Date(Y, M - 1, 1), days = new Date(Y, M, 0).getDate();
    let html = ["日","一","二","三","四","五","六"].map(w => `<div class="wd">${w}</div>`).join("");
    for (let i = 0; i < first.getDay(); i++) html += `<div class="day blank"></div>`;
    for (let d = 1; d <= days; d++) {
      const key = `${Y}-${String(M).padStart(2,"0")}-${String(d).padStart(2,"0")}`; const f = byDate[key];
      html += `<div class="day ${f ? "has" : ""} ${state.dates.has(key) ? "on" : ""} ${curB && f && !chg[key] ? "dim" : ""}" data-date="${key}"><div class="n">${d}</div>${f ? cornerHtml(chg[key]) : ""}
        ${f ? `<div class="c">${fmt(f.cases)} <span style="font-size:11px;font-weight:500">箱</span></div><div class="p">${f.po_count} 張 PO · ${f.rows} 品項</div>${f.missing_box ? `<span class="warn" title="${f.missing_box} 筆算不出箱數"><i class="bi bi-exclamation-triangle-fill"></i></span>` : ""}` : ""}</div>`;
    }
    return html;
  };
  $("#cal").innerHTML = months.length > 1
    ? months.map(([Y, M]) => `<div class="cal-month" style="grid-column:1/-1"><div class="cm-title">${Y} 年 ${M} 月</div><div class="cal">${grid(Y, M)}</div></div>`).join("")
    : grid(y, m);
  // 全部訂單跟依匯入批次篩選同一套操作：點、拖、Shift 都是選日期；批次模式壓淡的天（這批沒動到）不給選，
  // 只選一天時順便捲到那天並黃一下（afterPick）
  const sel = curB ? ".day.has:not(.dim)" : ".day.has";
  const keys = curB ? dateKeys().filter(k => chg[k]) : dateKeys();
  bindDragSelect($("#cal"), sel, el => el.dataset.date, (el, on) => el.classList.toggle("on", on), keys);
  renderNoDateChip();
}
/* 拖選：按下去那一格原本沒選就是「選」模式，反之「取消」模式；拖過去的格子畫面立刻變色，
   放開才重撈一次表格（行事曆不重畫）。用「滑鼠底下是哪一格」判斷，快滑也不漏格。
   Shift＋點：從上一次點的那格到這格，中間整段一起選。 */
let lastPick = null;
function bindDragSelect(root, sel, keyOf, paint, orderedKeys) {
  let drag = null;
  const apply = (k, on) => { on ? state.dates.add(k) : state.dates.delete(k); root.querySelectorAll(sel).forEach(el => { if (keyOf(el) === k) paint(el, on); }); };
  root.onpointerdown = e => {
    const el = e.target.closest(sel); if (!el || !root.contains(el) || e.button !== 0) return;
    if (e.target.closest("input[type=checkbox]")) e.preventDefault();
    e.preventDefault();
    const k = keyOf(el);
    if (e.shiftKey && lastPick && orderedKeys.includes(lastPick)) {
      const [a, b] = [orderedKeys.indexOf(lastPick), orderedKeys.indexOf(k)].sort((x, y) => x - y);
      orderedKeys.slice(a, b + 1).forEach(kk => apply(kk, true)); lastPick = k; afterPick(); return;
    }
    drag = { on: !state.dates.has(k), seen: new Set([k]) }; apply(k, drag.on); lastPick = k;
    root.setPointerCapture?.(e.pointerId);
  };
  root.onpointermove = e => {
    if (!drag) return;
    const under = document.elementFromPoint(e.clientX, e.clientY); const el = under && under.closest(sel);
    if (!el || !root.contains(el)) return;
    const k = keyOf(el); if (drag.seen.has(k)) return;
    drag.seen.add(k); apply(k, drag.on);
  };
  root.onpointerup = root.onpointercancel = () => { if (!drag) return; drag = null; afterPick(); };
}
// 選完日期：重撈表格（行事曆不重畫）；依匯入批次篩選中只選一天的話，順便捲到那天、黃一下
function afterPick() {
  loadOrders({ keepCal: true }).then(() => { if (state.batch && state.dates.size === 1) jumpToDate([...state.dates][0]); });
}
/* 從行事曆／日期清單跳到某一天：展開那天和它底下的 PO、捲過去、整塊亮 3 秒；其他日期維持原狀 */
function jumpToDate(d) {
  state.closedDates.delete(d);
  for (const r of state.rows) if ((r.delivery_date || "") === d) state.openPos.add(r.po_number);
  state.hlDate = d; renderOrders();
  const g = [...document.querySelectorAll("#orders-list .card[data-date]")].find(x => x.dataset.date === d);
  if (g) g.scrollIntoView({ behavior: "smooth", block: "start" });
}
function renderNoDateChip() {
  const nd = (state.facets?.dates || []).find(d => !d.date), chip = $("#chip-nodate");
  if (nd) { chip.classList.remove("hidden"); chip.classList.toggle("on", state.dates.has("")); chip.innerHTML = `<i class="bi bi-calendar-x"></i> 未排日期 ${nd.rows} 筆`; chip.onclick = () => { state.dates.has("") ? state.dates.delete("") : state.dates.add(""); loadOrders(); }; }
  else chip.classList.add("hidden");
}
/* 日期清單：只列有出貨的日子，一列一天。點一列勾一天、按著往下拖一次勾好幾天、Shift 點兩列中間整段一起勾。
   依匯入批次篩選時改成「這批動到的日期」：一列一天、點一列跳到那天；超過 6 天先收起來一行「還有 N 天」。 */
function renderDateList(chg, curB) {
  const wd = ["日","一","二","三","四","五","六"];
  const dayLabel = d => { const dt = new Date(d); return `<b>${Number(d.slice(5, 7))}/${Number(d.slice(8))}</b> <span class="muted">週${wd[dt.getDay()]}</span>`; };
  const badges = c => c ? `${c.n.size ? `<span class="badge bd-ok">新 ${c.n.size} PO</span> ` : ""}${c.u.size ? `<span class="badge bd-warn">變 ${c.u.size} PO</span> ` : ""}${c.g.size ? `<span class="badge bd-bad">消 ${c.g.size} PO</span>` : ""}` : "";
  if (curB) {
    // 批次模式：只列這批動到的天，操作跟全部訂單一樣（勾、拖、Shift），多一欄「這批」
    const byDate = new Map();
    for (const r of state.rows) { const d = r.delivery_date || ""; const o = byDate.get(d) || { pos: new Set(), rows: 0, cases: 0 }; o.pos.add(r.po_number); o.rows++; o.cases += r.cases || 0; byDate.set(d, o); }
    const touched = Object.keys(chg).filter(Boolean).sort();
    const facetBy = Object.fromEntries((state.facets?.dates || []).map(d => [d.date, d]));
    $("#cal-list").innerHTML = touched.length ? `<table class="dl"><thead><tr><th style="width:34px"></th><th>日期</th><th>PO</th><th>品項</th><th>箱數</th><th>這批</th></tr></thead><tbody>${touched.map(d => { const f = facetBy[d] || { po_count: 0, rows: 0, cases: 0 };
        return `<tr data-date="${d}" class="${state.dates.has(d) ? "on" : ""}"><td><input type="checkbox" ${state.dates.has(d) ? "checked" : ""} tabindex="-1"></td><td>${dayLabel(d)}</td><td>${f.po_count} 張 PO</td><td class="muted">${f.rows} 品項</td><td><b>${fmt(f.cases)}</b> 箱</td><td>${badges(chg[d])}</td></tr>`; }).join("")}
        <tr><td colspan="6" class="muted" style="text-align:center">這批共動到 ${touched.length} 天 · ${new Set(state.rows.map(r => r.po_number)).size} 張 PO · ${state.rows.length} 品項</td></tr></tbody></table>` : `<div class="muted p-3">這批沒有動到任何訂單。</div>`;
    bindDragSelect($("#cal-list"), "tbody tr[data-date]", el => el.dataset.date, (el, on) => { el.classList.toggle("on", on); const cb = el.querySelector("input"); if (cb) cb.checked = on; }, touched);
    return;
  }
  const ds = (state.facets?.dates || []).filter(d => d.date).sort((a, b) => a.date < b.date ? -1 : 1);
  $("#cal-list").innerHTML = ds.length ? `<table class="dl"><thead><tr><th style="width:34px"></th><th>日期</th><th>PO</th><th>品項</th><th>箱數</th><th>最近一批</th></tr></thead><tbody>${ds.map(d =>
      `<tr data-date="${d.date}" class="${state.dates.has(d.date) ? "on" : ""}"><td><input type="checkbox" ${state.dates.has(d.date) ? "checked" : ""} tabindex="-1"></td><td>${dayLabel(d.date)}</td><td>${d.po_count} 張 PO</td><td class="muted">${d.rows} 品項</td><td><b>${fmt(d.cases)}</b> 箱${d.missing_box ? ` <span class="badge bd-warn">${d.missing_box} 筆算不出</span>` : ""}</td><td>${badges(chg[d.date])}</td></tr>`).join("")}</tbody></table>` : `<div class="muted p-3">這個範圍沒有排日期的訂單。</div>`;
  bindDragSelect($("#cal-list"), "tbody tr[data-date]", el => el.dataset.date, (el, on) => { el.classList.toggle("on", on); const cb = el.querySelector("input"); if (cb) cb.checked = on; }, ds.map(d => d.date));
}
$("#cal-view").querySelectorAll("button").forEach(b => b.addEventListener("click", () => { state.calView = b.dataset.v; localStorage.setItem("mst_calview", state.calView); $("#cal-view").querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b)); renderCalendar(); }));

function ddRender(menuId, nId, items, set, key, labelFn, summaryFn) {
  const menu = $(menuId);
  const rows = items.map(it => `<label data-t="${esc(String(it[key] || "").toLowerCase())}"><input type="checkbox" data-v="${esc(it[key])}" ${set.has(it[key]) ? "checked" : ""}> ${labelFn(it)}</label>`).join("");
  const head = `<div class="dd-head">${items.length > 12 ? `<input type="text" placeholder="輸入關鍵字過濾…" class="dd-q">` : `<span>${items.length} 個選項</span>`}${set.size ? `<a class="dd-clear">取消勾選（${set.size}）</a>` : ""}</div>`;
  menu.innerHTML = items.length ? head + rows : `<div class="muted text-sm p-2">沒有選項</div>`;
  menu.querySelectorAll("input[type=checkbox]").forEach(cb => cb.addEventListener("change", () => { cb.checked ? set.add(cb.dataset.v) : set.delete(cb.dataset.v); loadOrders(); }));
  const q = menu.querySelector(".dd-q"); if (q) q.addEventListener("input", () => { const t = q.value.trim().toLowerCase(); menu.querySelectorAll("label").forEach(l => l.style.display = !t || l.dataset.t.includes(t) || l.innerText.toLowerCase().includes(t) ? "" : "none"); });
  const clr = menu.querySelector(".dd-clear"); if (clr) clr.addEventListener("click", () => { set.clear(); loadOrders(); });
  // 按鈕上直接寫目前選了什麼，不用點開才知道
  const n = $(nId); const vals = [...set];
  n.textContent = !vals.length ? "全部" : summaryFn ? summaryFn(vals) : (vals.length <= 2 ? vals.map(v => v || "（空白）").join("、") : `${vals.length} 個`);
  n.closest("button").classList.toggle("on", !!vals.length);
}
function renderDropdowns() {
  const f = state.facets || {};
  ddRender("#dd-line", "#dd-line-n", f.lines || [], state.lines, "line", it => `${it.line === "未分類" ? `<span class="badge bd-bad">${esc(it.line)}</span>` : esc(it.line)}<span class="muted" style="margin-left:auto">${it.rows}</span>`);
  ddRender("#dd-po", "#dd-po-n", f.pos || [], state.pos, "po_number", it => `<span class="mono">${esc(it.po_number)}</span> ${it.lines.map(l => `<span class="line-tag">${esc(l)}</span>`).join("")}<span class="muted" style="margin-left:auto">${esc(it.date || "未排")} · ${fmt(it.cases)} 箱</span>`, vals => `${vals.length} 張`);
  ddRender("#dd-brand", "#dd-brand-n", f.brands || [], state.brands, "brand", it => `${esc(it.brand)}<span class="muted" style="margin-left:auto">${it.rows}</span>`);
  ddRender("#dd-wh", "#dd-wh-n", f.warehouses || [], state.warehouses, "warehouse", it => `${esc(it.warehouse) || "（未填倉別）"}<span class="muted" style="margin-left:auto">${it.rows}</span>`, vals => vals.length <= 2 ? vals.map(v => v || "（未填）").join("、") : `${vals.length} 個`);
}
[["#dd-line-btn","#dd-line"],["#dd-po-btn","#dd-po"],["#dd-brand-btn","#dd-brand"],["#dd-wh-btn","#dd-wh"]].forEach(([b, m]) => {
  $(b).addEventListener("click", e => { e.stopPropagation(); document.querySelectorAll(".dd-menu").forEach(x => { if ("#" + x.id !== m) x.classList.add("hidden"); }); $(m).classList.toggle("hidden"); });
});
document.addEventListener("click", e => { if (!e.target.closest(".dd")) document.querySelectorAll(".dd-menu").forEach(m => m.classList.add("hidden")); });

function lineTag(o) { return `<span class="line-tag" title="原始線別：${esc(o.line || "空白")}">${esc(o.line_group)}${o.line && o.line !== o.line_group ? `·${esc(o.line.replace(/^CPG-/, ""))}` : ""}</span>`; }

function renderOrders() {
  const rows = state.rows;
  $("#orders-count").textContent = rows.length ? `${rows.length} 筆 · 合計 ${fmt(rows.reduce((s, r) => s + (r.cases || 0), 0))} 箱` : "";
  const curB = state.batch ? importBatches.find(b => b.id === state.batch) : null;
  if (!rows.length && curB) {
    const others = curB.month_counts.filter(x => x.m && x.m !== state.month);
    $("#orders-list").innerHTML = `<div class="card p-10 text-center"><div class="muted">${esc(whenText(curB.committed_at))} 這批在 <b>${Number(state.month.slice(5))} 月</b> 沒有動到任何訂單${anyFilter() ? "（或被篩選條件濾掉了）" : ""}。</div>
      ${others.length ? `<div class="mt-3 flex gap-2 justify-center flex-wrap">${others.map(x => `<button class="chip month-chip" data-m="${esc(x.m)}">變動在 ${Number(x.m.slice(5))} 月 → 切過去</button>`).join("")}</div>` : ""}</div>`;
    $("#orders-list").querySelectorAll(".month-chip").forEach(b => b.addEventListener("click", () => setMonth(b.dataset.m)));
    return;
  }
  $("#date-rail").classList.add("hidden"); document.body.classList.remove("rail-on");
  if (!rows.length) { $("#orders-list").innerHTML = `<div class="card p-10 text-center muted">${anyScope() ? "沒有符合篩選的資料" : `${state.month} 這個月還沒有任何訂單，按上面的「上傳訂單彙總表」。`}</div>`; return; }
  const byDate = new Map();
  for (const r of rows) { const d = r.delivery_date || ""; if (!byDate.has(d)) byDate.set(d, new Map()); const pos = byDate.get(d); if (!pos.has(r.po_number)) pos.set(r.po_number, []); pos.get(r.po_number).push(r); }
  const dates = [...byDate.keys()].sort((a, b) => (a || "9999") < (b || "9999") ? -1 : 1);
  let html = "";
  for (const d of dates) {
    const pos = byDate.get(d); const all = [...pos.values()].flat();
    const total = all.reduce((s, r) => s + (r.cases || 0), 0); const miss = all.filter(r => r.cases == null).length;
    const dClosed = state.closedDates.has(d);
    html += `<div class="card ${dClosed ? "collapsed" : ""} ${state.hlDate === d ? "hl" : ""}" data-date="${esc(d)}"><div class="grp-date" data-date="${esc(d)}" title="點一下收合／展開這一天"><i class="bi bi-chevron-down tg ${dClosed ? "closed" : ""}"></i> <i class="bi bi-calendar-event"></i> ${d ? esc(d) + " 交貨" : "未排日期"}
        <span class="badge bd-blue">${pos.size} 張 PO</span><span class="badge bd-gray">${all.length} 品項</span>
        <span style="margin-left:auto">出貨合計 <b style="font-size:16px">${fmt(total)}</b> 箱</span>${miss ? `<span class="badge bd-warn">${miss} 筆算不出箱數</span>` : ""}</div>`;
    for (const [po, list] of pos) {
      const sub = list.reduce((s, r) => s + (r.cases || 0), 0); const ov = list.some(r => r.delivery_date_overridden);
      const groups = [...new Set(list.map(r => r.line_group))];
      const pOpen = state.openPos.has(po);
      html += `<div class="grp-po ${pOpen ? "" : "collapsed"}" data-po="${esc(po)}" title="點一下展開／收合這張 PO 的品項">
        <i class="bi bi-chevron-down tg ${pOpen ? "" : "closed"}"></i>
        <input type="checkbox" class="po-sel" data-po="${esc(po)}" ${state.selected.has(po) ? "checked" : ""} title="勾選後可批次改交貨日">
        <span class="po" data-po="${esc(po)}" title="點開整張單一次修改">${esc(po)}</span><i class="bi bi-clipboard po-copy" data-po="${esc(po)}" title="複製單號"></i>
        ${groups.map(g => `<span class="line-tag">${esc(g)}</span>`).join("")}
        <span class="badge bd-gray">${esc(list[0].warehouse || "—")}</span>
        <span class="muted">${list.length} 品項 · ${fmt(sub)} 箱</span>
        <span style="margin-left:auto" class="flex items-center gap-2"><span class="kbd">交貨日</span>
          <input type="date" value="${esc(d)}" data-po="${esc(po)}" data-old="${esc(d)}" class="po-date">
          ${ov ? `<span class="badge bd-warn" title="整合表上的交貨日是 ${esc(list[0].delivery_date_file || "空白")}">人工改期</span><button class="btn btn-g btn-sm po-date-reset" data-po="${esc(po)}" data-old="${esc(d)}" title="恢復成整合表的交貨日"><i class="bi bi-arrow-counterclockwise"></i></button>` : ""}
        </span></div>
      <div class="po-body" style="overflow:auto"><table class="m"><thead><tr><th>線別</th><th>SKU ID</th><th>國條</th><th>品類</th><th>品名</th><th>品牌</th><th class="num">下單</th><th class="num">出貨數量</th><th>單位</th><th class="num">箱入數</th><th class="num">出貨(箱)</th><th>備註 <span class="kbd" style="color:var(--b100)">Note＋OP</span></th><th></th></tr></thead><tbody>`;
      for (const r of list) {
        const boxBadge = r.box_source === "master" ? "" : r.box_source === "file" ? `<span class="badge bd-warn" title="主檔裡沒有這個商品，箱入數先用訂單彙總表自帶的；把它補進主檔（或重匯酷澎主檔）就會以主檔為準">用整合表的</span>` : `<span class="badge bd-bad">缺</span>`;
        // 色條＋小標籤：預設對「最近一次匯入」標，選了某一次就對那一次標
        const mark = state.batch || (importBatches[0] && importBatches[0].id);
        const kind = !mark ? "" : r.first_batch_id === mark ? "new" : r.last_batch_id === mark ? (r.missing_in_file ? "gone" : "upd") : "";
        const chg = kind === "new" ? `<span class="badge bd-ok">新增</span>` : kind === "upd" ? `<span class="badge bd-warn" title="${esc(r.last_batch_changes || "")}">有變</span>` : kind === "gone" ? `<span class="badge bd-bad" title="${esc(r.last_batch_changes || "")}">消失</span>` : "";
        html += `<tr data-id="${r.id}" data-ver="${r.version}" class="${kind ? "chg-" + kind : ""}">
          <td>${lineTag(r)}${r.line ? "" : ` <span class="badge bd-bad">沒線別</span>`}${chg ? " " + chg : ""}${state.batch && (kind === "upd" || kind === "gone") && r.last_batch_changes ? `<div class="kbd" style="max-width:220px;white-space:normal">${esc(r.last_batch_changes)}</div>` : ""}</td>
          <td class="mono">${esc(r.sku_id)}</td>
          <td class="mono">${esc(r.barcode)}${r.in_master ? "" : ` <span class="badge bd-warn" title="總表／主檔沒有這個國條">未建檔</span>`}</td>
          <td>${esc(r.category || "")}</td>
          <td style="max-width:300px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(r.product_name)}">${esc(r.product_name)}</td>
          <td>${esc(r.brand_master || r.brand)}</td>
          <td class="num muted">${fmt(r.qty_coupang)}</td>
          <td class="num inline-cell ${r.qty_ship_overridden ? "edited" : ""}" data-field="qty_ship" title="${r.missing_in_file ? "最近一次上傳的整合表裡，這張 PO 已沒有這個品項，出貨數量歸 0（再出現會自動恢復）" : r.qty_ship_overridden ? "人工調整過，匯入不會覆蓋（整合表是 " + fmt(r.qty_file_ship ?? r.qty_coupang) + "）。點一下可改" : "點一下可改"}">
            ${r.missing_in_file ? `<span class="badge bd-bad">檔案已無此品項</span> ` : ""}<b>${fmt(r.qty_ship)}</b>${r.qty_ship_overridden ? ` <i class="bi bi-arrow-counterclockwise qty-reset" title="恢復成整合表數字" style="cursor:pointer;color:var(--warn)"></i>` : ""}</td>
          <td>${esc(r.unit)}</td>
          <td class="num">${fmt(r.box_size) || "—"} ${boxBadge}</td>
          <td class="num"><b>${r.cases == null ? `<span class="neg">—</span>` : fmt(r.cases)}</b></td>
          <td class="inline-cell ${r.remarks_overridden ? "edited" : ""}" data-field="remarks" title="${r.remarks_file ? "整合表自帶的備註：" + esc(r.remarks_file) + "（不會進報價檔）。" : ""}點一下可改">
            ${r.note_master ? `<div class="kbd" style="color:var(--b800)">Note：${esc(r.note_master)}</div>` : ""}<span class="rem-val">${esc(r.remarks_overridden ? r.remarks : "")}</span>${!r.remarks_overridden && !r.note_master ? `<span class="kbd">（OP 備註）</span>` : ""}</td>
          <td><button class="btn btn-danger btn-sm del" title="刪除這筆"><i class="bi bi-trash"></i></button></td></tr>`;
      }
      html += `</tbody></table></div>`;
    }
    html += `</div>`;
  }
  $("#orders-list").innerHTML = html;
  bindOrderEdits();
  // 兩層開合：點日期標題收合整天；點 PO 標題列展開品項。裡面的勾選框、單號、日期欄、複製不觸發開合
  $("#orders-list").querySelectorAll(".grp-date").forEach(h => h.addEventListener("click", () => { const d = h.dataset.date; state.closedDates.has(d) ? state.closedDates.delete(d) : state.closedDates.add(d); state.hlDate = null; renderOrders(); }));
  $("#orders-list").querySelectorAll(".grp-po").forEach(h => h.addEventListener("click", e => {
    if (e.target.closest("input, .po, .po-copy, button, .badge.bd-warn")) return;
    const po = h.dataset.po; state.openPos.has(po) ? state.openPos.delete(po) : state.openPos.add(po); state.hlDate = null; renderOrders();
  }));
  $("#orders-list").querySelectorAll(".po-copy").forEach(b => b.addEventListener("click", async e => { e.stopPropagation(); try { await navigator.clipboard.writeText(b.dataset.po); toast(`已複製 ${b.dataset.po}`); } catch { toast("瀏覽器不讓複製，請手動選取", "warn"); } }));
  if (state.hlDate) setTimeout(() => { state.hlDate = null; }, 3200);
  renderDateRail();
}
/* 右側日期導覽條：表格裡有哪幾天就列哪幾天，捲到哪天亮哪天，點哪個跳哪天。少於 3 天不顯示 */
let railObserver = null;
function renderDateRail() {
  const rail = $("#date-rail"); const cards = [...document.querySelectorAll("#orders-list .card[data-date]")];
  if (railObserver) { railObserver.disconnect(); railObserver = null; }
  if (cards.length < 3) { rail.classList.add("hidden"); document.body.classList.remove("rail-on"); return; }
  rail.classList.remove("hidden");
  rail.innerHTML = cards.map(c => { const d = c.dataset.date; const n = c.querySelectorAll(".grp-po").length; return `<a data-date="${esc(d)}" title="${d ? d + " 交貨" : "未排日期"} · ${n} 張 PO">${d ? `${Number(d.slice(5, 7))}/${Number(d.slice(8))}` : "未排"}<span class="n">${n}</span></a>`; }).join("");
  rail.querySelectorAll("a").forEach(a => a.addEventListener("click", () => { const c = cards.find(x => x.dataset.date === a.dataset.date); if (c) c.scrollIntoView({ behavior: "smooth", block: "start" }); }));
  railObserver = new IntersectionObserver(entries => {
    const vis = entries.filter(e => e.isIntersecting).sort((a, b) => a.boundingClientRect.top - b.boundingClientRect.top);
    if (!vis.length) return;
    const d = vis[0].target.dataset.date; rail.querySelectorAll("a").forEach(a => a.classList.toggle("on", a.dataset.date === d));
  }, { rootMargin: "-80px 0px -60% 0px", threshold: 0 });
  cards.forEach(c => railObserver.observe(c));
  updateRailVisibility();
}
// 導覽條只在表格捲進畫面時出現，免得蓋到上面的行事曆
function updateRailVisibility() {
  const rail = $("#date-rail"); const list = $("#orders-list");
  if (!rail.innerHTML || list.querySelectorAll(".card[data-date]").length < 3) return;
  const r = list.getBoundingClientRect(); const show = r.top < window.innerHeight * 0.6 && r.bottom > 120;
  rail.classList.toggle("hidden", !show);
  // 導覽條出現時主畫面右邊讓一點位，免得蓋到表格最右邊的刪除鍵
  document.body.classList.toggle("rail-on", show);
}
window.addEventListener("scroll", updateRailVisibility, { passive: true });
$("#density").querySelectorAll("button").forEach(b => b.addEventListener("click", () => { state.dense = b.dataset.d === "dense"; localStorage.setItem("mst_dense", state.dense ? "1" : "0"); $("#density").querySelectorAll("button").forEach(x => x.classList.toggle("on", x === b)); $("#orders-list").classList.toggle("dense", state.dense); }));
$("#orders-list").classList.toggle("dense", state.dense); $("#density").querySelectorAll("button").forEach(x => x.classList.toggle("on", (x.dataset.d === "dense") === state.dense));
$("#btn-expand-all").addEventListener("click", () => { state.closedDates.clear(); state.rows.forEach(r => state.openPos.add(r.po_number)); state.hlDate = null; renderOrders(); });
$("#btn-collapse-all").addEventListener("click", () => { state.openPos.clear(); state.hlDate = null; renderOrders(); });

function startInline(cell) {
  if (cell.dataset.editing) return;
  const tr = cell.closest("tr"); const field = cell.dataset.field;
  const row = state.rows.find(r => r.id == tr.dataset.id); if (!row) return;
  cell.dataset.editing = "1";
  const cur = field === "qty_ship" ? (row.qty_ship ?? "") : (row.remarks_overridden ? row.remarks : "");
  const keep = cell.innerHTML;
  cell.innerHTML = `<input class="${field === "qty_ship" ? "num" : ""}" value="${esc(cur)}" ${field === "qty_ship" ? 'type="number" min="0"' : 'placeholder="OP 備註，例如 打單缺貨"'}>`;
  const inp = cell.querySelector("input"); inp.focus(); inp.select();
  // done 只負責「同一次編輯不要處理兩遍」（Enter 之後 blur 又會觸發一次）；
  // 還原畫面的 restore 一定要能執行，不能被 done 擋掉——之前就是這裡卡住，
  // 沒改值按 Enter／點外面，格子永遠留在輸入框。
  let done = false;
  const restore = () => { delete cell.dataset.editing; cell.innerHTML = keep; };
  const cancel = () => { if (done) return; done = true; restore(); };
  const save = async () => {
    if (done) return; done = true;
    if (String(inp.value) === String(cur)) { restore(); return; }
    try {
      const body = { version: Number(tr.dataset.ver) }; body[field] = inp.value;
      const d = await api(`/api/master/orders/${tr.dataset.id}`, J(body));
      await loadOrders();
      const fresh = document.querySelector(`tr[data-id="${tr.dataset.id}"] td[data-field="${field}"]`);
      if (fresh) fresh.classList.add("save-flash");
      toast(field === "qty_ship" ? `已儲存，出貨 ${fmt(d.row.qty_ship)} → ${d.row.cases == null ? "算不出箱數" : fmt(d.row.cases) + " 箱"}` : "備註已儲存");
    } catch (e) { if (e.status === 409) loadOrders(); else { toast(e.message, "err"); restore(); } }
  };
  inp.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); save(); } if (e.key === "Escape") cancel(); });
  inp.addEventListener("blur", () => setTimeout(save, 0));
}
function bindCell(cell) { cell.addEventListener("click", e => { if (e.target.classList.contains("qty-reset")) return; startInline(cell); }); }
function bindOrderEdits() {
  const list = $("#orders-list");
  list.querySelectorAll("td.inline-cell").forEach(bindCell);
  list.querySelectorAll(".qty-reset").forEach(b => b.addEventListener("click", async e => {
    e.stopPropagation(); const tr = b.closest("tr");
    try { await api(`/api/master/orders/${tr.dataset.id}`, J({ version: Number(tr.dataset.ver), reset_qty_ship: true })); toast("已恢復成整合表數字"); loadOrders(); }
    catch (err) { if (err.status !== 409) toast(err.message, "err"); else loadOrders(); }
  }));
  list.querySelectorAll(".del").forEach(b => b.addEventListener("click", async () => {
    const tr = b.closest("tr"); if (!confirm("確定刪除這筆品項？會留下歷程。")) return;
    try { await api(`/api/master/orders/${tr.dataset.id}`, { method: "DELETE" }); toast("已刪除"); loadOrders(); } catch (e) { toast(e.message, "err"); }
  }));
  list.querySelectorAll(".po-date").forEach(inp => inp.addEventListener("change", async () => {
    if (!inp.value) { inp.value = inp.dataset.old; return; }
    try { const d = await api("/api/master/pos/date", J({ po_numbers: [inp.dataset.po], delivery_date: inp.value, expected: { [inp.dataset.po]: inp.dataset.old } }));
      toast(`PO ${inp.dataset.po} 已改到 ${d.delivery_date}，${d.moved} 個品項一起搬過去`); loadOrders(); }
    catch (e) { if (e.status === 409) loadOrders(); else { toast(e.message, "err"); inp.value = inp.dataset.old; } }
  }));
  list.querySelectorAll(".po-date-reset").forEach(b => b.addEventListener("click", async () => {
    try { await api("/api/master/pos/date", J({ po_numbers: [b.dataset.po], reset: true, expected: { [b.dataset.po]: b.dataset.old } })); toast(`PO ${b.dataset.po} 已恢復成整合表的交貨日`); loadOrders(); }
    catch (e) { if (e.status === 409) loadOrders(); else toast(e.message, "err"); }
  }));
  list.querySelectorAll(".po-sel").forEach(cb => cb.addEventListener("change", () => { cb.checked ? state.selected.add(cb.dataset.po) : state.selected.delete(cb.dataset.po); renderBatch(); }));
  list.querySelectorAll(".grp-po .po").forEach(el => el.addEventListener("click", () => openPo(el.dataset.po)));
}

/* ── 批次改期 ── */
function renderBatch() {
  const n = state.selected.size; $("#batch-n").textContent = n; $("#batchbar").classList.toggle("show", n > 0);
  const all = [...new Set(state.rows.map(r => r.po_number))]; $("#sel-all-po").checked = all.length > 0 && all.every(p => state.selected.has(p));
}
$("#sel-all-po").addEventListener("change", e => { const all = new Set(state.rows.map(r => r.po_number)); if (e.target.checked) all.forEach(p => state.selected.add(p)); else all.forEach(p => state.selected.delete(p)); renderOrders(); renderBatch(); });
$("#batch-clear").addEventListener("click", () => { state.selected.clear(); renderOrders(); renderBatch(); });
async function batchMove(reset) {
  const pos = [...state.selected]; if (!pos.length) return;
  const date = $("#batch-date").value; if (!reset && !date) { toast("先選要改到哪一天。", "warn"); return; }
  const expected = {}; for (const r of state.rows) if (state.selected.has(r.po_number)) expected[r.po_number] = expected[r.po_number] ?? (r.delivery_date || "");
  if (!confirm(reset ? `把 ${pos.length} 張 PO 的交貨日恢復成整合表的日期？` : `把 ${pos.length} 張 PO 全部改到 ${date}？每張單底下所有品項會一起搬。`)) return;
  try { const d = await api("/api/master/pos/date", J({ po_numbers: pos, delivery_date: date, reset, expected })); toast(`${d.po_count} 張 PO、${d.moved} 個品項${reset ? "已恢復整合表日期" : "已改到 " + d.delivery_date}`); state.selected.clear(); loadOrders(); }
  catch (e) { if (e.status === 409) loadOrders(); else toast(e.message, "err"); }
}
$("#batch-apply").addEventListener("click", () => batchMove(false)); $("#batch-reset").addEventListener("click", () => batchMove(true));

/* ── PO 視窗 ── */
let PO = null;
async function openPo(po) {
  try { PO = await api(`/api/master/pos/${encodeURIComponent(po)}`); } catch (e) { toast(e.message, "err"); return; }
  $("#po-title").textContent = PO.po_number;
  $("#po-sub").textContent = `${PO.lines.join("、")}${PO.lines_raw.length > 1 ? "（" + PO.lines_raw.join("／") + "）" : ""} · ${PO.warehouse || "—"} · ${PO.rows.length} 品項`;
  $("#po-date").value = PO.delivery_date; $("#po-date").dataset.old = PO.delivery_date;
  $("#po-date-note").textContent = PO.dates.length > 1 ? `這張單的品項目前分在 ${PO.dates.join("、")}，存檔會統一` : (PO.date_overridden ? `人工改期過，整合表是 ${PO.delivery_date_file || "空白"}` : "");
  $("#po-date-reset").classList.toggle("hidden", !PO.date_overridden);
  $("#po-total").textContent = fmt(PO.total_cases); $("#po-msg").textContent = "";
  $("#po-table").innerHTML = `<thead><tr><th>線別</th><th>國條</th><th>品名</th><th class="num">下單</th><th class="num">出貨數量</th><th class="num">箱入數</th><th class="num">出貨(箱)</th><th>OP 備註</th></tr></thead><tbody>
    ${PO.rows.map(r => `<tr data-id="${r.id}" data-ver="${r.version}">
      <td>${lineTag(r)}</td><td class="mono">${esc(r.barcode)}</td>
      <td style="max-width:280px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(r.product_name)}">${esc(r.product_name)}${r.note_master ? `<div class="kbd" style="color:var(--b800)">Note：${esc(r.note_master)}</div>` : ""}</td>
      <td class="num muted">${fmt(r.qty_coupang)}</td>
      <td class="num"><input type="number" min="0" class="txt po-qty num ${r.qty_ship_overridden ? "edited" : ""}" style="width:90px;text-align:right;padding:3px 6px" value="${r.qty_ship ?? ""}" data-orig="${r.qty_ship ?? ""}" data-box="${r.box_size ?? ""}"></td>
      <td class="num">${fmt(r.box_size) || "—"}</td>
      <td class="num po-cases"><b>${r.cases == null ? "—" : fmt(r.cases)}</b></td>
      <td><input class="txt po-rem" style="padding:3px 6px" value="${esc(r.remarks_overridden ? r.remarks : "")}" data-orig="${esc(r.remarks_overridden ? r.remarks : "")}" placeholder="例如 打單缺貨"></td></tr>`).join("")}</tbody>`;
  $("#po-table").querySelectorAll(".po-qty").forEach(inp => inp.addEventListener("input", () => {
    const box = Number(inp.dataset.box); const c = inp.closest("tr").querySelector(".po-cases b"); c.textContent = box && inp.value !== "" ? fmt(Number(inp.value) / box) : "—";
    let tot = 0; $("#po-table").querySelectorAll(".po-qty").forEach(i => { const b = Number(i.dataset.box); if (b && i.value !== "") tot += Number(i.value) / b; }); $("#po-total").textContent = fmt(tot);
  }));
  $("#po-logs").innerHTML = PO.logs.length ? PO.logs.map(l => `<div style="border-bottom:1px solid var(--line2);padding:3px 0"><span class="muted">${esc(l.changed_at)}</span> <b>${esc(l.operator)}</b> <span class="badge ${l.source === "import" ? "bd-blue" : "bd-gray"}">${l.source === "import" ? "匯入" : l.source === "system" ? "系統" : "手動"}</span><br>${esc(l.field_label)}${l.barcode ? ` <span class="mono kbd">${esc(l.barcode)}</span>` : ""}：${esc(fmt(l.old_value) || "空")} → <b>${esc(fmt(l.new_value) || "空")}</b>${l.note ? `<div class="kbd">${esc(l.note)}</div>` : ""}</div>`).join("") : `<div class="muted">還沒有修改紀錄</div>`;
  $("#dlg-po").showModal();
}
$("#po-date-reset").addEventListener("click", async () => {
  try { await api(`/api/master/pos/${encodeURIComponent(PO.po_number)}/save`, J({ reset_date: true, expected_date: $("#po-date").dataset.old, items: [] })); toast("已恢復成整合表的交貨日"); openPo(PO.po_number); loadOrders(); }
  catch (e) { if (e.status === 409) { $("#dlg-po").close(); loadOrders(); } else toast(e.message, "err"); }
});
$("#po-save").addEventListener("click", async () => {
  if (!PO) return;
  const items = [];
  $("#po-table").querySelectorAll("tbody tr").forEach(tr => {
    const q = tr.querySelector(".po-qty"), r = tr.querySelector(".po-rem"); const it = { id: Number(tr.dataset.id), version: Number(tr.dataset.ver) }; let touched = false;
    if (q.value !== q.dataset.orig) { it.qty_ship = q.value; touched = true; }
    if (r.value !== r.dataset.orig) { it.remarks = r.value; touched = true; }
    if (touched) items.push(it);
  });
  const body = { items }; const nd = $("#po-date").value;
  if (nd && nd !== $("#po-date").dataset.old) { body.delivery_date = nd; body.expected_date = $("#po-date").dataset.old; }
  if (!items.length && !body.delivery_date) { $("#po-msg").textContent = "沒有任何變動。"; return; }
  $("#po-save").disabled = true;
  try { const d = await api(`/api/master/pos/${encodeURIComponent(PO.po_number)}/save`, J(body)); toast(`已儲存，記錄了 ${d.changed} 項變更`); $("#dlg-po").close(); loadOrders(); }
  catch (e) { if (e.status === 409) { $("#dlg-po").close(); loadOrders(); } else $("#po-msg").innerHTML = `<span class="neg">${esc(e.message)}</span>`; }
  finally { $("#po-save").disabled = false; }
});

/* ── 匯入訂單 ── */
let pendingBatch = null;
$("#btn-upload").addEventListener("click", () => { stagedFiles = []; renderStaged(); $("#imp-result").classList.add("hidden"); $("#btn-commit").classList.add("hidden"); $("#imp-addmissing-wrap").classList.add("hidden"); pendingBatch = null; $("#dlg-import").showModal(); });
/* 檔案先放進暫存清單，可以分好幾次、從不同資料夾加；每加一次就把清單裡的全部重新解析一遍，
   預覽永遠是「目前清單」的結果，按確認匯入才真的寫進去。 */
let stagedFiles = [];
function renderStaged() {
  $("#imp-files").classList.toggle("hidden", !stagedFiles.length);
  $("#imp-files-list").innerHTML = stagedFiles.map((f, i) => `<span class="chip" style="cursor:default"><i class="bi bi-file-earmark-excel"></i> ${esc(f.name)} <span class="muted">${(f.size / 1024).toFixed(0)} KB</span> <button data-i="${i}" class="rm-staged" title="移除這個檔" style="margin-left:4px;color:var(--bad)"><i class="bi bi-x-lg"></i></button></span>`).join("");
  $("#imp-files-list").querySelectorAll(".rm-staged").forEach(b => b.addEventListener("click", () => { stagedFiles.splice(Number(b.dataset.i), 1); stagedChanged(); }));
}
function stagedChanged() {
  renderStaged();
  if (!stagedFiles.length) { $("#imp-result").classList.add("hidden"); $("#btn-commit").classList.add("hidden"); $("#imp-addmissing-wrap").classList.add("hidden"); pendingBatch = null; return; }
  previewImport(stagedFiles);
}
dropzone($("#dz"), $("#file-import"), files => {
  let dup = 0;
  files.forEach(f => { if (stagedFiles.some(x => x.name === f.name && x.size === f.size)) dup++; else stagedFiles.push(f); });
  if (dup) toast(`${dup} 個檔案已經在清單裡，略過`, "err");
  stagedChanged();
});
$("#imp-files-clear").addEventListener("click", () => { stagedFiles = []; stagedChanged(); });
async function previewImport(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  const box = $("#imp-result"); box.classList.remove("hidden"); box.innerHTML = `<p class="muted">解析中…</p>`;
  let d; try { d = await api("/api/master/import/preview", { method: "POST", body: fd }); }
  catch (e) { box.innerHTML = `<div class="badge bd-bad" style="font-size:13px;padding:8px 12px">${esc(e.message)}</div>`; return; }
  pendingBatch = d.batch_id;
  const lines = Object.entries(d.lines).map(([l, n]) => `<span class="badge ${l === "未分類" ? "bd-bad" : "bd-blue"}">${esc(l)} ${n} 筆</span>`).join(" ");
  const raws = Object.entries(d.lines_raw).map(([l, n]) => `${esc(l) || "（空白）"} ${n}`).join("、");
  const months = [...new Set(d.dates.map(x => x.slice(0, 7)))];
  const changes = d.updated.map(u => `<tr><td class="mono">${esc(u.po_number)}</td><td class="mono">${esc(u.barcode)}</td><td style="max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(u.product_name)}</td>
     <td>${u.changes.map(c => `<div><span class="muted">${esc(c.label)}</span> ${esc(fmt(c.old) || "空白")} → <b>${esc(fmt(c.new) || "空白")}</b>${(c.field === "qty_file_ship" && u.qty_ship_overridden) || (c.field === "delivery_date_file" && u.delivery_date_overridden) ? ` <span class="badge bd-warn">人工調整過，不覆蓋</span>` : ""}</div>`).join("")}</td></tr>`).join("");
  const missing = d.missing_products.map(m => `<tr><td>${esc(m.line) || "—"}</td><td class="mono">${esc(m.barcode)}</td><td>${esc(m.product_name)}</td><td class="num">${m.box_size ?? `<span class="neg">空白</span>`}</td></tr>`).join("");
  box.innerHTML = `
    <div class="flex gap-3 flex-wrap mb-3">
      <div class="stat"><b>${d.rows_total}</b><span>讀到的品項 · ${d.po_count} 張 PO${d.files && d.files.length > 1 ? ` · ${d.files.length} 份檔案` : ""}</span></div>
      <div class="stat"><b style="color:var(--ok)">${d.new_count}</b><span>新增</span></div>
      <div class="stat"><b style="color:var(--warn)">${d.updated_count}</b><span>內容有變</span></div>
      <div class="stat"><b class="muted">${d.identical_count}</b><span>沒變（略過）</span></div>
      ${d.removed_count ? `<div class="stat" style="border-color:#fecaca;background:var(--badbg)"><b style="color:var(--bad)">${d.removed_count}</b><span>檔案裡消失的品項</span></div>` : ""}
    </div>
    ${d.files && d.files.length > 1 ? `<div class="text-sm mb-1">檔案：${d.files.map(f => `<span class="badge bd-gray">${esc(f)}</span>`).join(" ")}</div>` : ""}
    <div class="text-sm mb-1">線別：${lines || "<span class='muted'>沒有線別欄</span>"} <span class="kbd">（檔案原始值：${raws}）</span></div>
    <div class="text-sm mb-2">月份：${months.map(x => `<span class="badge bd-gray">${esc(x)}</span>`).join(" ")}　交貨日 ${d.dates.length} 天</div>
    ${d.warnings.length ? `<div class="card p-3 mb-3 text-sm" style="background:var(--warnbg);border-color:#fcd34d"><b>提醒</b><ul class="list-disc pl-5">${d.warnings.slice(0, 12).map(w => `<li>${esc(w)}</li>`).join("")}${d.warnings.length > 12 ? `<li>…還有 ${d.warnings.length - 12} 則</li>` : ""}</ul></div>` : ""}
    ${d.removed && d.removed.length ? `<b class="text-sm" style="color:var(--bad)">這些 PO 在檔案裡，但底下這些品項不見了（${d.removed.length}）— 匯入後會留著、出貨數量歸 0</b><div class="tbl-wrap mb-3" style="max-height:200px"><table class="m"><thead><tr><th>PO</th><th>國條</th><th>品名</th><th class="num">下單</th><th class="num">目前出貨</th></tr></thead><tbody>${d.removed.map(x => `<tr><td class="mono">${esc(x.po_number)}</td><td class="mono">${esc(x.barcode)}</td><td>${esc(x.product_name)}</td><td class="num">${fmt(x.qty_coupang)}</td><td class="num">${fmt(x.qty_ship)}</td></tr>`).join("")}</tbody></table></div>` : ""}
    ${changes ? `<b class="text-sm">內容有變的品項</b><div class="tbl-wrap mb-3" style="max-height:240px"><table class="m"><thead><tr><th>PO</th><th>國條</th><th>品名</th><th>變動</th></tr></thead><tbody>${changes}</tbody></table></div>` : ""}
    ${missing ? `<b class="text-sm">總表裡還沒有的商品（${d.missing_products.length}）</b><div class="tbl-wrap" style="max-height:200px"><table class="m"><thead><tr><th>線別</th><th>國條</th><th>品名</th><th class="num">整合表箱入數</th></tr></thead><tbody>${missing}</tbody></table></div>` : `<div class="text-sm" style="color:var(--ok)"><i class="bi bi-check-circle"></i> 所有國條主檔都有</div>`}`;
  $("#btn-commit").classList.remove("hidden"); $("#imp-addmissing-wrap").classList.toggle("hidden", !missing);
}
$("#btn-commit").addEventListener("click", async () => {
  if (!pendingBatch) return; $("#btn-commit").disabled = true;
  try { const d = await api("/api/master/import/commit", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({ batch_id: pendingBatch, add_missing_products: $("#imp-addmissing").checked }) });
    toast(`匯入完成：新增 ${d.inserted}、更新 ${d.updated}、沒變 ${d.identical}${d.removed ? `，${d.removed} 筆品項在檔案裡消失（已歸 0 保留）` : ""}${d.products_added ? `，主檔新增 ${d.products_added} 筆` : ""}`);
    stagedFiles = []; renderStaged(); $("#dlg-import").close(); await loadMeta(); await loadImportScopes();
    // 匯完直接切到「依匯入批次篩選：剛這批」，第一眼就是剛匯進來的那幾張 PO，不用自己找
    if (importBatches[0]) { pickBatch(importBatches[0].id); toast(`剛匯進來：新增 ${importBatches[0].new_pos} 張 PO（${importBatches[0].new_now} 品項）。畫面已只剩這批，右上「清除批次篩選」回全部。`); checkMaster(); }
    else refresh(); }
  catch (e) { toast(e.message, "err"); } finally { $("#btn-commit").disabled = false; }
});

/* ════════════ ③ 總表 ════════════ */
/* 匯出總表的樣子：① 匯進主檔的那份業務總表（有 M/D交貨 日期欄的）就是底稿，匯出只填箱數、含順序。
   沒匯過總表的線別，匯出用系統自己排的格式。 */
async function loadTemplateInfo() {
  let t; try { t = await api(`/api/master/template?line=${encodeURIComponent(state.sumLine)}`); } catch (e) { return; }
  const info = $("#tpl-info");
  if (t.exists) {
    const months = Object.entries(t.months || {}).sort((a, b) => Number(a[0]) - Number(b[0])).map(([m, e]) => `${m}月 ${e.dates} 天${e.spare ? `＋${e.spare} 空欄` : ""}`).join("、");
    info.innerHTML = `<span class="badge bd-blue" title="${esc(t.uploaded_at)} 由 ${esc(t.uploaded_by)} 匯進主檔的那份。要換樣子就再匯一次新的總表。"><i class="bi bi-file-earmark-check"></i> 匯出照「${esc(t.filename)}」的樣子，含順序</span> <span class="kbd">${esc(months)}</span>`;
    $("#btn-export-2").title = `打開「${t.filename}」只填 ${state.sumMonth} 各交貨日箱數，其他一格不動；對不到的商品與沒欄位的日期寫在最後一個分頁「系統填入說明」`;
  } else {
    info.innerHTML = `<span class="kbd">${esc(state.sumLine)} 還沒匯過總表，匯出用系統格式；到 ① 把總表匯進主檔就會照它的樣子</span>`;
    $("#btn-export-2").title = "系統自己排的總表格式：一列一國條、往右各交貨日箱數";
  }
}
$("#sum-line").addEventListener("change", e => { state.sumLine = e.target.value; localStorage.setItem("mst_sum_line", state.sumLine); loadSummary(); });
$("#sum-month").addEventListener("change", e => { if (e.target.value) { state.sumMonth = e.target.value; $("#sum-exp-to").value = e.target.value; loadSummary(); } });
$("#sum-exp-to").addEventListener("change", loadSummary);
async function loadSummary() {
  if (!state.sumLine) { $("#sum-table").innerHTML = ""; $("#sum-stats").innerHTML = `<div class="muted p-4">還沒有訂單，先到訂單明細上傳。</div>`; return; }
  $("#btn-export-2").href = `/api/master/export?line=${encodeURIComponent(state.sumLine)}&month=${state.sumMonth}`;
  loadTemplateInfo();
  const sto = $("#sum-exp-to").value; const sumTo = (sto && sto > state.sumMonth) ? sto : state.sumMonth;
  $("#btn-export-daily-2").href = `/api/master/export/daily?month=${state.sumMonth}&month_to=${sumTo}&lines=${encodeURIComponent(state.sumLine)}`;
  let s; try { s = await api(`/api/master/summary?line=${encodeURIComponent(state.sumLine)}&month=${state.sumMonth}`); } catch (e) { toast(e.message, "err"); return; }
  const mm = Number(state.sumMonth.slice(5));
  $("#sum-stats").innerHTML = `<div class="stat"><b>${fmt(s.month_total)}</b><span>${mm} 月出貨總箱數</span></div><div class="stat"><b>${s.dates.length}</b><span>有出貨的交貨日</span></div><div class="stat"><b>${s.rows.length}</b><span>商品數</span></div>
    ${s.missing_box_rows ? `<div class="stat" style="border-color:#fcd34d;background:var(--warnbg)"><b style="color:var(--warn)">${s.missing_box_rows}</b><span>筆算不出箱數（缺箱入數）</span></div>` : ""}
    ${s.not_in_master ? `<div class="stat" style="border-color:#fcd34d;background:var(--warnbg)"><b style="color:var(--warn)">${s.not_in_master}</b><span>個國條不在主檔</span></div>` : ""}`;
  const md = d => { const [, m, dd] = d.split("-"); return `${Number(m)}/${Number(dd)}`; };
  let html = `<thead><tr><th class="sticky-l" style="z-index:3">國條</th><th>線別</th><th>品類</th><th>品名</th><th>品牌</th><th class="num">箱入數</th>${s.dates.map(d => `<th class="num">${md(d)}<br><span style="font-weight:400;font-size:11px">交貨</span></th>`).join("")}<th class="num" style="background:var(--b900)">${mm}月TTL</th><th>備註</th></tr></thead><tbody>`;
  for (const r of s.rows) html += `<tr><td class="mono sticky-l">${esc(r.barcode)} ${r.in_master ? "" : `<span class="badge bd-warn">未建檔</span>`}</td><td>${r.lines_raw.map(l => `<span class="line-tag">${esc(l) || "空白"}</span>`).join(" ")}</td><td>${esc(r.category || "")}</td>
      <td style="max-width:300px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(r.product_name)}">${esc(r.product_name)}</td><td>${esc(r.brand)}</td>
      <td class="num">${r.box_size ? fmt(r.box_size) : `<span class="neg">缺</span>`}</td>
      ${s.dates.map(d => `<td class="num">${r.by_date[d] != null ? fmt(r.by_date[d]) : `<span class="muted">·</span>`}</td>`).join("")}
      <td class="num" style="background:var(--b50)"><b>${fmt(r.month_total)}</b>${r.missing_box ? ` <span class="badge bd-warn" title="${r.missing_box} 筆沒算進去">!</span>` : ""}</td>
      <td class="muted" style="max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(r.note)}</td></tr>`;
  html += `</tbody><tfoot><tr style="font-weight:700;background:var(--b100)"><td class="sticky-l" style="background:var(--b100)">合計</td><td></td><td></td><td></td><td></td><td></td>${s.dates.map(d => `<td class="num">${fmt(s.totals_by_date[d])}</td>`).join("")}<td class="num">${fmt(s.month_total)}</td><td></td></tr></tfoot>`;
  $("#sum-table").innerHTML = html;
}

/* ════════════ ① 商品主檔 ════════════ */
let productsCache = [];
/* 總表也是先加進清單，按「匯入主檔」才一份一份寫進去（後面的會蓋前面的同國條） */
let stagedProd = [];
function renderStagedProd() {
  $("#prod-files").classList.toggle("hidden", !stagedProd.length);
  $("#prod-files-list").innerHTML = stagedProd.map((f, i) => `<span class="chip" style="cursor:default"><i class="bi bi-file-earmark-excel"></i> ${esc(f.name)} <span class="muted">${(f.size / 1024).toFixed(0)} KB</span> <button data-i="${i}" class="rm-staged-prod" title="移除這個檔" style="margin-left:4px;color:var(--bad)"><i class="bi bi-x-lg"></i></button></span>`).join("");
  $("#prod-files-list").querySelectorAll(".rm-staged-prod").forEach(b => b.addEventListener("click", () => { stagedProd.splice(Number(b.dataset.i), 1); renderStagedProd(); }));
}
dropzone($("#dz-prod"), $("#file-prod"), files => {
  let dup = 0;
  files.forEach(f => { if (stagedProd.some(x => x.name === f.name && x.size === f.size)) dup++; else stagedProd.push(f); });
  if (dup) toast(`${dup} 個檔案已經在清單裡，略過`, "err");
  $("#prod-imp-msg").innerHTML = ""; renderStagedProd();
});
$("#prod-files-clear").addEventListener("click", () => { stagedProd = []; renderStagedProd(); });
$("#prod-files-go").addEventListener("click", async () => {
  if (!stagedProd.length) return;
  const files = stagedProd.slice(); $("#prod-files-go").disabled = true;
  const out = []; $("#prod-imp-msg").innerHTML = `<span class="muted">匯入中…</span>`;
  for (const file of files) {
    const fd = new FormData(); fd.append("file", file);
    try { const d = await api("/api/master/products/import", { method: "POST", body: fd });
      out.push(`<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> ${files.length > 1 ? esc(file.name) + " " : ""}工作表「${esc(d.sheet)}」：新增 ${d.added}、更新 ${d.updated}、沒變 ${d.unchanged}。抓到的欄位：${d.columns_found.join("、")}${d.template_line ? `<br><i class="bi bi-file-earmark-check"></i> 這是${esc(d.template_line)}的總表，已記住它的樣子：之後「匯出總表」會長得跟這份一模一樣（含順序），只填箱數。` : ""}</span>`);
      toast(`主檔匯入完成：新增 ${d.added}、更新 ${d.updated}`); }
    catch (e) { out.push(`<span class="neg">${files.length > 1 ? esc(file.name) + "：" : ""}${esc(e.message)}</span>`); }
  }
  $("#prod-files-go").disabled = false; stagedProd = []; renderStagedProd();
  $("#prod-imp-msg").innerHTML = out.join("<br>"); loadProducts();
});
$("#pq").addEventListener("input", () => { clearTimeout(window._pq); window._pq = setTimeout(loadProducts, 200); });
$("#prod-line").addEventListener("change", loadProducts);
async function loadProducts() {
  let d; try { d = await api(`/api/master/products?line=${encodeURIComponent($("#prod-line").value)}&q=${encodeURIComponent($("#pq").value.trim())}`); } catch (e) { toast(e.message, "err"); return; }
  productsCache = d.products;
  const orph = $("#orphans");
  if (d.orphans.length) {
    orph.classList.remove("hidden");
    orph.innerHTML = `<b class="text-sm" style="color:var(--warn)"><i class="bi bi-exclamation-triangle-fill"></i> 總表裡還沒有的商品（${d.orphans.length}）</b><p class="kbd mb-2">訂單有下、但總表查不到這些商品，所以箱入數只能先用整合表自帶的，品類／Note／單價會是空的（就是以前 VLOOKUP 出 #N/A 的那些）。請補進總表後重匯，或按「加入主檔」先建起來。</p>
      <div class="tbl-wrap" style="max-height:220px"><table class="m"><thead><tr><th>線別</th><th>國條</th><th>品名</th><th>品牌</th><th class="num">整合表箱入數</th><th class="num">訂單筆數</th><th></th></tr></thead><tbody>
      ${d.orphans.map(o => `<tr><td><span class="line-tag">${esc(o.line_group)}</span></td><td class="mono">${esc(o.barcode)}</td><td>${esc(o.product_name)}</td><td>${esc(o.brand)}</td><td class="num">${o.box_size_file ?? `<span class="neg">空白</span>`}</td><td class="num">${o.order_rows}</td><td><button class="btn btn-o btn-sm add-orphan" data-o='${esc(JSON.stringify(o))}'><i class="bi bi-plus"></i> 加入主檔</button></td></tr>`).join("")}</tbody></table></div>`;
    orph.querySelectorAll(".add-orphan").forEach(b => b.addEventListener("click", () => { const o = JSON.parse(b.dataset.o); openProductDialog({ barcode: o.barcode, sku_id: o.sku_id, yf_sku: o.yf_sku, brand: o.brand, product_name: o.product_name, box_size: o.box_size_file }); }));
  } else orph.classList.add("hidden");
  $("#prod-count").textContent = `${d.total} 筆`;
  $("#prod-table").innerHTML = `<thead><tr><th>國條</th><th>線別</th><th>SKU ID</th><th>永豐料號</th><th>品類</th><th>品牌</th><th>品名</th><th>單位</th><th class="num">箱入數</th><th class="num">單價(含稅)</th><th>Note／報價備註</th><th class="num">訂單筆數</th><th>最後更新</th><th></th></tr></thead><tbody>
    ${d.products.length ? d.products.map(p => `<tr data-id="${p.id}" ${p.active === "N" ? 'style="opacity:.55"' : ""}>
      <td class="mono">${esc(p.barcode)}</td>
      <td>${p.line_groups.length ? p.line_groups.map(g => `<span class="line-tag">${esc(g)}</span>`).join(" ") : (p.order_rows ? `<span class="line-tag" style="color:var(--bad)">未分類</span>` : `<span class="kbd">尚未出現在訂單</span>`)}</td>
      <td class="mono">${esc(p.sku_id)}</td><td class="mono">${esc(p.yf_sku)}</td><td>${esc(p.category || "")}</td><td>${esc(p.brand)}</td>
      <td style="max-width:280px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${esc(p.product_name)}">${esc(p.product_name)}</td>
      <td>${esc(p.unit || "")}${p.active === "N" ? ` <span class="badge bd-gray" title="酷澎主檔標記為停用">停用</span>` : ""}</td>
      <td class="num">${p.box_size ? fmt(p.box_size) : `<span class="neg">缺</span>`}${p.auto_created ? ` <span class="badge bd-warn" title="匯入訂單時用整合表的箱入數自動建的，還沒人核對；存一次或重匯總表就解除">請核對</span>` : ""}</td>
      <td class="num">${p.cost_price != null ? fmt(p.cost_price) : `<span class="muted">—</span>`}</td>
      <td class="muted">${esc(p.note)}</td><td class="num">${p.order_rows}</td><td class="kbd">${esc(p.updated_at)}<br>${esc(p.updated_by)}</td>
      <td class="whitespace-nowrap"><button class="btn btn-g btn-sm edit" data-p='${esc(JSON.stringify(p))}'><i class="bi bi-pencil"></i></button> <button class="btn btn-danger btn-sm del"><i class="bi bi-trash"></i></button></td></tr>`).join("")
    : `<tr><td colspan="14" class="text-center muted p-8">還沒有主檔。把酷澎主檔（或寶僑總表）拖到上面的框裡，或匯入訂單時勾「自動建主檔」。</td></tr>`}</tbody>`;
  $("#prod-table").querySelectorAll(".edit").forEach(b => b.addEventListener("click", () => openProductDialog(JSON.parse(b.dataset.p))));
  $("#prod-table").querySelectorAll(".del").forEach(b => b.addEventListener("click", async () => { const id = b.closest("tr").dataset.id; if (!confirm("確定刪除這筆主檔？")) return;
    try { await api(`/api/master/products/${id}`, { method: "DELETE" }); toast("已刪除"); loadProducts(); } catch (e) { toast(e.message, "err"); } }));
}
function openProductDialog(p) {
  $("#prod-dlg-title").innerHTML = p && p.id ? `<i class="bi bi-pencil"></i> 編輯商品` : `<i class="bi bi-box-seam"></i> 新增商品`;
  $("#p-barcode").value = p?.barcode || ""; $("#p-barcode").readOnly = !!(p && p.id);
  $("#p-box").value = p?.box_size ?? ""; $("#p-sku").value = p?.sku_id || ""; $("#p-yf").value = p?.yf_sku || "";
  $("#p-cat").value = p?.category || ""; $("#p-cost").value = p?.cost_price ?? ""; $("#p-pg").value = p?.pgcode || "";
  $("#p-brand").value = p?.brand || ""; $("#p-name").value = p?.product_name || ""; $("#p-note").value = p?.note || "";
  $("#p-line").value = p?.master_line || ""; $("#p-unit").value = p?.unit || ""; $("#p-shelf").value = p?.shelf_days ?? ""; $("#p-active").value = p?.active === "N" ? "N" : "Y";
  $("#dlg-prod").showModal();
}
$("#btn-prod-add").addEventListener("click", () => openProductDialog(null));
$("#btn-prod-save").addEventListener("click", async () => {
  try { await api("/api/master/products", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({ barcode: $("#p-barcode").value, box_size: $("#p-box").value, sku_id: $("#p-sku").value, yf_sku: $("#p-yf").value, brand: $("#p-brand").value, product_name: $("#p-name").value, note: $("#p-note").value, category: $("#p-cat").value, cost_price: $("#p-cost").value, pgcode: $("#p-pg").value, master_line: $("#p-line").value, unit: $("#p-unit").value, shelf_days: $("#p-shelf").value, active: $("#p-active").value }) });
    toast("主檔已儲存"); $("#dlg-prod").close(); loadProducts(); } catch (e) { toast(e.message, "err"); }
});

/* ── 修改歷程 ── */
async function loadLogs() {
  const d = await api(`/api/master/logs?q=${encodeURIComponent($("#lg-q").value.trim())}&limit=300`);
  $("#lg-table").innerHTML = `<thead><tr><th>時間</th><th>誰</th><th>來源</th><th>PO</th><th>國條</th><th>欄位</th><th>改前</th><th>改後</th><th>說明</th></tr></thead><tbody>${d.logs.map(l => `<tr><td class="kbd">${esc(l.changed_at)}</td><td>${esc(l.operator)}</td><td><span class="badge ${l.source === "import" ? "bd-blue" : "bd-gray"}">${l.source === "import" ? "匯入" : l.source === "system" ? "系統" : "手動"}</span></td><td class="mono">${esc(l.po_number)}</td><td class="mono">${esc(l.barcode)}</td><td>${esc(l.field_label)}</td><td>${esc(fmt(l.old_value))}</td><td><b>${esc(fmt(l.new_value))}</b></td><td class="kbd">${esc(l.note)}</td></tr>`).join("") || `<tr><td colspan="9" class="muted p-6 text-center">沒有紀錄</td></tr>`}</tbody>`;
}
$("#btn-logs").addEventListener("click", () => { $("#dlg-logs").showModal(); loadLogs(); });

/* ════════════ 清除資料（只有管理員有按鈕） ════════════ */
if ($("#btn-reset")) {
  const resetBtnState = () => { $("#rs-go").disabled = $("#rs-confirm").value.trim() !== "清空資料" || !($("#rs-orders").checked || $("#rs-products").checked); };
  $("#btn-reset").addEventListener("click", () => { $("#rs-confirm").value = ""; resetBtnState(); $("#dlg-reset").showModal(); $("#rs-confirm").focus(); });
  ["#rs-confirm", "#rs-orders", "#rs-products"].forEach(sel => $(sel).addEventListener("input", resetBtnState));
  $("#rs-go").addEventListener("click", async () => {
    $("#rs-go").disabled = true;
    try {
      const d = await api("/api/master/reset", { method: "POST", headers: {"Content-Type":"application/json"}, body: JSON.stringify({
        confirm: $("#rs-confirm").value, orders: $("#rs-orders").checked,
        products: $("#rs-products").checked, keep_logs: $("#rs-keeplogs").checked }) });
      $("#dlg-reset").close(); $("#rs-confirm").value = "";
      toast(d.message);
      await loadMeta(); loadImportScopes(); refresh();
    } catch (e) { toast(e.message, "err"); resetBtnState(); }
  });
}
$("#lg-go").addEventListener("click", loadLogs); $("#lg-q").addEventListener("keydown", e => { if (e.key === "Enter") loadLogs(); });

/* ── 啟動 ── */
(async () => {
  await loadMeta();
  if (await checkMaster()) setTab("products");
  refresh();
})();
