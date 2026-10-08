/* 瑪氏出貨：商品總表上傳、拆單、回填 EIP 採購單號／約倉時間、產瑪氏採購單、採購單設定。後端在 mars/。 */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmt = n => n == null ? "" : Number(n).toLocaleString("en-US", { maximumFractionDigits: 4 });
function toast(msg, kind) { const t = document.createElement("div"); t.className = "toast" + (kind === "err" ? " err" : ""); t.textContent = msg; ($("#toasts") || document.body).appendChild(t); setTimeout(() => t.remove(), kind === "err" ? 5000 : 2500); }
async function api(url, opts = {}) {
  const res = await fetch(url, opts); let data = {}; try { data = await res.json(); } catch (e) {}
  if (!res.ok) throw Object.assign(new Error(data.error || `伺服器錯誤（${res.status}）`), { status: res.status, data });
  return data;
}
const md = d => { const [, m, dd] = d.split("-"); return `${Number(m)}/${Number(dd)}`; };
const today = () => new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
const STATUS = { new: "未產出", generated: "已產出，待填 EIP 採購單號", filled: "已回填", changed: "訂單已變更，請重新下載 EIP 採購單", changed_after_eip: "EIP 送出後訂單已變更", gone: "訂單已無此份" };
let VIEW = null, OPEN = new Set(), CHECKED = new Set();
let DATES = [], Q = "", NOEIP = false, SORT = "date";        // 有瑪氏訂單的到貨日（搜尋全部日期用）、搜尋字、只看還沒填 EIP 單號的

/* ── 商品總表 ── */
async function loadStatus() {
  const d = await api("/api/mars/status");
  DATES = d.dates || [];
  const u = d.last_upload;
  $("#mp-status").innerHTML = u ? `上次上傳 ${esc(u.uploaded_at.slice(0, 16))} · ${esc(u.operator)} · ${esc(u.filename)} · <b style="color:var(--m8)">${d.codes}</b> 個料號`
    : `<span class="neg">尚未上傳</span>：拆單需依商品總表區分品類與中標，請先上傳 Alice 的「自動化_Mars整合商品資料」。`;
  $("#od-status").innerHTML = d.dates.length ? `已匯入瑪氏訂單的到貨日：${d.dates.slice(-6).map(md).join("、")}${d.dates.length > 6 ? " …" : ""}${d.last_order_import ? `<br>最近一次匯入 ${esc(d.last_order_import.committed_at.slice(0, 16))} · ${esc(d.last_order_import.operator)} · ${esc(d.last_order_import.filename)}` : ""}`
    : `<span class="neg">尚無瑪氏訂單</span>：請上傳 Kate 系統匯出的訂單彙總表。`;
  const up = d.dates.filter(x => x >= today());
  if (!$("#d-from").value) { const first = up[0] || d.dates[d.dates.length - 1] || today(); $("#d-from").value = $("#d-to").value = first; }
  CAL_MONTH = CAL_MONTH || $("#d-from").value.slice(0, 7);
  await loadCalendar();
}

/* ── 月曆：有瑪氏訂單的天亮起來；點一天、Shift 點一段 ── */
let CAL_MONTH = null, CAL = null;
async function loadCalendar() {
  try { CAL = await api(`/api/mars/calendar?month=${CAL_MONTH}`); } catch (e) { toast(e.message, "err"); return; }
  renderCalendar();
}
function renderCalendar() {
  if (!CAL) return;
  const [Y, M] = CAL_MONTH.split("-").map(Number); const f = $("#d-from").value, t = $("#d-to").value, tod = today();
  $("#cal-title").textContent = `${Y} 年 ${M} 月到貨`;
  const first = new Date(Y, M - 1, 1), days = new Date(Y, M, 0).getDate();
  let h = ["日", "一", "二", "三", "四", "五", "六"].map(w => `<div class="wd">${w}</div>`).join("");
  for (let i = 0; i < first.getDay(); i++) h += `<div class="day blank"></div>`;
  for (let d = 1; d <= days; d++) {
    const key = `${Y}-${String(M).padStart(2, "0")}-${String(d).padStart(2, "0")}`, c = CAL.days[key];
    const on = f && t && key >= f && key <= t;
    h += `<div class="day ${c ? "has" : ""} ${on ? "on" : ""} ${key === tod ? "today" : ""}" data-date="${key}"><div class="n">${d}</div>` + (c ? `
      <span class="eip ${c.filled === c.files ? "ok" : c.filled ? "part" : ""}" title="已填 EIP 採購單號的份數">${c.filled}／${c.files}</span>
      <div class="c">${fmt(c.cases)} <span style="font-size:10.5px;font-weight:500">箱</span></div><div class="p">${c.pos} 張 PO · ${c.files} 份${c.gone ? ` · ${c.gone} 份失效` : ""}</div>
      ${c.unmatched || c.blocking ? `<span class="warn" title="${c.unmatched ? `${c.unmatched} 個品項在商品總表中找不到` : ""}${c.blocking ? ` ${c.blocking} 份箱數不是整數` : ""}"><i class="bi bi-exclamation-triangle-fill"></i></span>` : ""}` : "") + `</div>`;
  }
  $("#cal").innerHTML = h;
  $("#cal").querySelectorAll(".day.has").forEach(el => el.addEventListener("click", e => {
    const k = el.dataset.date;
    if (e.shiftKey && $("#d-from").value) { const a = $("#d-from").value; $("#d-from").value = a < k ? a : k; $("#d-to").value = a < k ? k : a; }
    else { $("#d-from").value = $("#d-to").value = k; }
    renderCalendar(); loadSplits();
  }));
}
const shiftMonth = n => { const [Y, M] = CAL_MONTH.split("-").map(Number); const dt = new Date(Y, M - 1 + n, 1); CAL_MONTH = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`; loadCalendar(); };
$("#cal-prev").addEventListener("click", () => shiftMonth(-1)); $("#cal-next").addEventListener("click", () => shiftMonth(1));
$("#cal-toggle").addEventListener("click", () => { const w = $("#cal-wrap"); w.classList.toggle("hidden"); try { localStorage.setItem("mars_cal", w.classList.contains("hidden") ? "0" : "1"); } catch (e) {} });
try { if (localStorage.getItem("mars_cal") === "0") $("#cal-wrap").classList.add("hidden"); } catch (e) {}
$("#mp-file").addEventListener("change", e => { const f = e.target.files[0]; e.target.value = ""; if (f) uploadMaster(f); });
async function uploadMaster(f) {
  const fd = new FormData(); fd.append("file", f); $("#mp-msg").innerHTML = `<span class="muted">上傳中…</span>`;
  try {
    const d = await api("/api/mars/products/import", { method: "POST", body: fd });
    const cats = Object.entries(d.by_category).map(([k, v]) => `${k} ${v}`).join("、");
    $("#mp-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 已讀取 ${d.codes} 個料號（${d.rows} 列，${esc(cats)}）${d.added || d.removed ? `，與上一版相比：新增 ${d.added} 個、移除 ${d.removed} 個` : ""}。</span>`
      + (d.warnings.length ? `<details class="kbd"><summary>${d.warnings.length} 列有問題（已略過或標示）</summary>${d.warnings.map(esc).join("<br>")}</details>` : "");
    toast("商品總表已更新"); await loadStatus(); loadSplits();
  } catch (err) { $("#mp-msg").innerHTML = `<span class="neg">${esc(err.message)}</span>`; }
}
/* 拖檔進卡片就等於按上傳。預設收 Excel；勇信那區收 PDF，用 opt 指定檔案類型和擋下時的提示 */
function dropzone(el, onFiles, opt = {}) {
  const re = opt.re || /\.xlsx?$/i, msg = opt.msg || "請拖曳 Excel 檔（.xlsx）";
  ["dragenter", "dragover"].forEach(ev => el.addEventListener(ev, e => { e.preventDefault(); el.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => el.addEventListener(ev, e => { e.preventDefault(); el.classList.remove("over"); }));
  el.addEventListener("drop", e => { const fs = [...e.dataTransfer.files].filter(f => re.test(f.name)); if (fs.length) onFiles(fs); else toast(msg, "err"); });
}
dropzone($("#dz-mp"), fs => uploadMaster(fs[0]));
dropzone($("#dz-od"), fs => previewOrders(fs));

/* ── 訂單彙總表：走 ② 訂單明細同一套匯入（預覽 → 確認），只是從這頁傳 ── */
let ODP = null;
$("#od-file").addEventListener("change", e => { const files = [...e.target.files]; e.target.value = ""; if (files.length) previewOrders(files); });
async function previewOrders(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  $("#od-msg").innerHTML = `<span class="muted">讀檔中…</span>`; $("#od-preview").classList.add("hidden");
  try { ODP = await api("/api/master/import/preview", { method: "POST", body: fd }); } catch (err) { $("#od-msg").innerHTML = `<span class="neg">${esc(err.message)}</span>`; return; }
  $("#od-msg").innerHTML = "";
  const mars = ODP.lines["瑪氏"] || 0, others = Object.entries(ODP.lines).filter(([k]) => k !== "瑪氏");
  if (!mars) {
    $("#od-preview").innerHTML = `<div class="alert al-bad"><b>此檔案沒有瑪氏訂單</b>（${others.map(([k, v]) => `${esc(k)} ${v} 列`).join("、") || "0 列"}），本頁只匯入瑪氏的訂單彙總表，此次未匯入。寶僑訂單請至商品主檔自動化的「訂單明細」分頁上傳。</div>`;
    $("#od-preview").classList.remove("hidden"); ODP = null; return;
  }
  $("#od-preview").innerHTML = `<div style="border:1px solid var(--m2);border-radius:10px;padding:10px 12px;background:var(--m0)">
    <div><b>${esc(ODP.filename)}</b>：${ODP.rows_total} 列、${ODP.po_count} 張 PO，瑪氏 <b style="color:var(--m8)">${mars}</b> 列${others.length ? `，另有 ${others.map(([k, v]) => `${esc(k)} ${v} 列`).join("、")}（會一併匯入商品主檔自動化的訂單明細）` : ""}。</div>
    <div class="mt-1">新增 <b>${ODP.new_count}</b> 筆、更新 <b>${ODP.updated_count}</b> 筆、未變動 ${ODP.identical_count} 筆${ODP.removed_count ? `、檔案中已無 <span class="neg">${ODP.removed_count}</span> 筆（出貨數量將歸 0）` : ""}；到貨日 ${ODP.dates.map(md).join("、")}。</div>
    ${ODP.warnings.length ? `<div class="mt-1" style="color:var(--warn)">${ODP.warnings.map(esc).join("<br>")}</div>` : ""}
    <div class="kbd mt-1">在商品主檔自動化的訂單明細中，我方修改過的出貨數量、到貨日不會被覆蓋；只更新檔案裡有的 PO。</div>
    <div class="flex gap-2 mt-2"><button id="od-commit" class="btn btn-p btn-sm"><i class="bi bi-check2"></i> 確認匯入</button><button id="od-cancel" class="btn btn-o btn-sm">取消</button></div></div>`;
  $("#od-preview").classList.remove("hidden");
  $("#od-cancel").onclick = () => { $("#od-preview").classList.add("hidden"); ODP = null; };
  $("#od-commit").onclick = async () => {
    $("#od-commit").disabled = true;
    try { const d = await api("/api/master/import/commit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ batch_id: ODP.batch_id, add_missing_products: true }) });
      $("#od-preview").classList.add("hidden"); $("#od-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 匯入完成：新增 ${d.inserted} 筆、更新 ${d.updated} 筆、未變動 ${d.identical} 筆${d.removed ? `、歸 0 ${d.removed} 筆` : ""}。</span>`;
      toast("訂單彙總表匯入完成，已切換至本批到貨日"); const ds = (ODP.dates || []).slice().sort(); ODP = null; CAL_MONTH = ds.length ? ds[0].slice(0, 7) : CAL_MONTH; await loadStatus();
      if (ds.length) { $("#d-from").value = ds[0]; $("#d-to").value = ds[ds.length - 1]; }     // 匯完直接跳到這批的到貨日
      loadSplits();
    } catch (err) { toast(err.message, "err"); $("#od-commit").disabled = false; }
  };
}

/* ── 拆單 ── */
["#d-from", "#d-to"].forEach(s => $(s).addEventListener("change", loadSplits));
/* 搜尋：在已載入的期間裡找，不用重新問伺服器；用空白隔開多個關鍵字，全部符合才列出。料號、品名在品項裡，符合的那份自動展開 */
$("#q").addEventListener("input", () => { Q = $("#q").value.trim().toLowerCase(); if (VIEW) render(); });
$("#q").addEventListener("keydown", e => { if (e.key === "Escape") { $("#q").value = ""; Q = ""; if (VIEW) render(); } });
$("#f-noeip").addEventListener("click", () => { NOEIP = !NOEIP; $("#f-noeip").classList.toggle("on", NOEIP); if (VIEW) render(); });
/* 排序：依到貨日（伺服器給的順序）或依倉別（同倉再照到貨日、酷澎 PO），同一張 PO 的幾份還是排在一起（Jerry 2026-10-08）。選過的記在這台電腦 */
try { SORT = localStorage.getItem("mars_sort") === "wh" ? "wh" : "date"; } catch (e) { /* 無痕視窗等讀不到就用預設 */ }
$("#f-sort").value = SORT;
$("#f-sort").addEventListener("change", () => { SORT = $("#f-sort").value; try { localStorage.setItem("mars_sort", SORT); } catch (e) { /* 存不了就算了 */ } if (VIEW) render(); });
const WH_CMP = new Intl.Collator("en", { numeric: true });   // TAO3 排在 TAO10 前面
function matchSplit(r) {
  if (NOEIP && r.eip_po) return { ok: false };
  if (!Q) return { ok: true, byItem: false };
  const words = Q.split(/\s+/).filter(Boolean);
  const head = [r.po_number, r.eip_po, r.warehouse, r.filename, r.category, r.unit, r.label === "V" ? "需貼中標" : "不貼中標"].join("\n").toLowerCase();
  const items = r.items.map(i => [i.yf_sku, i.purchase_code, i.product_name, i.mars_code].join("\n").toLowerCase());
  const ok = words.every(w => head.includes(w) || items.some(t => t.includes(w)));
  return { ok, byItem: ok && !words.every(w => head.includes(w)) };
}
async function loadSplits() {
  const f = $("#d-from").value, t = $("#d-to").value; if (!f || !t) return;
  if (CAL && f.slice(0, 7) !== CAL_MONTH && f.slice(0, 7) === t.slice(0, 7)) { CAL_MONTH = f.slice(0, 7); loadCalendar(); } else renderCalendar();
  try { VIEW = await api(`/api/mars/splits?from=${f}&to=${t}`); } catch (e) { toast(e.message, "err"); return; }
  render();
}
function render() {
  const v = VIEW, s = v.summary;
  $("#sum").innerHTML = `<span><b>${s.pos}</b> 張 PO</span><span><b>${s.files}</b> 份拆單表</span><span title="盒、包沒有相同料號時合成一張 EIP 採購單"><b>${s.eip_files}</b> 張 EIP 採購單</span><span><b>${fmt(s.cases)}</b> 箱</span><span><b>${s.items}</b> 個品項</span><span><b>${s.filled}</b>／${s.files} 已回填 EIP 採購單號</span><span><b>${s.po_ready}</b>／${s.files} 可產出瑪氏採購單</span>`;
  const al = [];
  if (!v.has_products) al.push(`<div class="alert al-bad"><i class="bi bi-exclamation-octagon-fill"></i> 尚未上傳瑪氏商品總表，無法區分品類與中標，無法拆單。</div>`);
  if (v.unmatched.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-exclamation-octagon-fill"></i> ${v.unmatched.length} 個品項在瑪氏商品總表中找不到</b>，不會列入拆單表，請更新商品總表後重新拆單：<br>${v.unmatched.map(u => `${esc(u.po_number)} · ${esc(u.yf_sku)}${u.quote_note && u.quote_note !== u.yf_sku ? `（報價備註 ${esc(u.quote_note)}）` : ""} · ${esc(u.product_name)} · ${esc(u.unit)} ${fmt(u.qty_ship)}`).join("<br>")}</div>`);
  const blocking = v.splits.flatMap(r => r.blocking || []);
  if (blocking.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-slash-circle"></i> 以下品項需先處理才能產出</b>（EIP 採購單只能填整數箱）：<br>${blocking.map(esc).join("<br>")}</div>`);
  if (v.zero_rows) al.push(`<div class="kbd">出貨數量為 0 的 ${v.zero_rows} 個品項未列入拆單。</div>`);
  if (v.wh_missing && v.wh_missing.length) al.push(`<div class="kbd"><i class="bi bi-geo-alt"></i> 倉庫資料未填齊：<b>${v.wh_missing.map(esc).join("、")}</b>。採購單仍可產出，缺少的欄位會留空，請至下方「採購單設定」補齊。</div>`);
  $("#ps-warn").classList.toggle("hidden", !(v.wh_missing && v.wh_missing.length)); $("#ps-warn").textContent = v.wh_missing && v.wh_missing.length ? `${v.wh_missing.join("、")} 未填齊` : "";
  $("#alerts").innerHTML = al.join("");
  $("#btn-gen").disabled = !v.splits.length || !v.has_products || blocking.length > 0;
  $("#btn-po").disabled = !s.po_ready;
  const usable = v.splits.filter(r => r.status !== "gone" && !(r.blocking && r.blocking.length)).map(r => r.split_key);
  CHECKED = new Set([...CHECKED].filter(k => usable.includes(k)));
  $("#btn-emma-all").disabled = !usable.length;

  const matched = new Map(v.splits.map(r => [r.split_key, matchSplit(r)]));
  const shown = v.splits.filter(r => matched.get(r.split_key).ok);
  if (SORT === "wh") shown.sort((a, b) => WH_CMP.compare(a.warehouse || "", b.warehouse || "") || (a.delivery_date || "").localeCompare(b.delivery_date || "") || (a.po_number || "").localeCompare(b.po_number || ""));
  if (Q || NOEIP) $("#sum").innerHTML += `<span>符合 <b style="color:var(--m8)">${shown.length}</b>／${v.splits.length} 份</span>`;

  let h = `<thead><tr><th><input type="checkbox" id="ck-all" title="全選"></th><th>狀態</th><th>拆單表檔名</th><th>到貨日</th><th>酷澎 PO</th><th>倉</th><th>品類</th>${v.split_by_unit ? "<th>單位</th>" : ""}<th>中標</th><th class="num">品項</th><th class="num">箱數</th><th style="min-width:150px">EIP 採購單號</th><th style="min-width:170px">約倉時間</th><th>下載</th></tr></thead><tbody>`;
  const ncol = v.split_by_unit ? 14 : 13;
  let lastPo = null, gi = 0;
  for (const r of shown) {
    const key = r.split_key, warn = r.items.some(i => i.issues && i.issues.length);
    // 還沒產出的列也直接能填、能按：第一次動到時先把這一份存起來（ensureSplit）。只有箱數算不出整數的那份不行
    const blocked = r.blocking && r.blocking.length, ok = !blocked && r.status !== "gone" || r.id;
    // 同一個 EIP 採購單號的幾份（盒、包同一張）：多一顆「合併」，產一張瑪氏採購單（Jerry 2026-10-08，同事要的）
    // 品類或中標不同的不能合（表頭只能寫一種），那種就不出現這顆
    const sameEip = r.eip_po ? v.splits.filter(x => x.eip_po === r.eip_po && x.po_number === r.po_number && x.delivery_date === r.delivery_date && x.warehouse === r.warehouse
      && x.category === r.category && (x.label || "") === (r.label || "")) : [];
    const dis = blocked && !r.id ? "disabled title=\"此份有品項需先處理（見上方紅色提示）\"" : "";
    if (r.po_number !== lastPo) {
      lastPo = r.po_number; gi++;
      const same = v.splits.filter(x => x.po_number === r.po_number);
      const filled = same.filter(x => x.eip_po).length, hit = same.filter(x => matched.get(x.split_key).ok).length;
      h += `<tr class="pog"><td colspan="${ncol}"><i class="bi bi-receipt"></i> PO ${esc(r.po_number)}<span class="kbd">${md(r.delivery_date)} 到貨 · ${esc(r.warehouse)} · 拆成 ${same.length} 份 · ${fmt(same.reduce((a, x) => a + (x.cases_total || 0), 0))} 箱 · 已回填 ${filled}／${same.length}${hit < same.length ? ` · 符合 ${hit} 份` : ""}</span></td></tr>`;
    }
    h += `<tr class="sp ${gi % 2 ? "g1" : "g0"}" data-k="${esc(key)}">
      <td>${ok ? `<input type="checkbox" class="ck" data-k="${esc(key)}" ${CHECKED.has(key) ? "checked" : ""} title="勾選後可合併下載 EMMA 匯入檔">` : ""}</td>
      <td><span class="badge st-${r.status}">${STATUS[r.status]}</span>${r.diff ? `<div class="kbd" style="max-width:220px;color:var(--bad)">${esc(r.diff)}</div>` : ""}</td>
      <td><span class="fn fname">${esc(r.filename)}</span> <button class="btn btn-o btn-sm cp" title="複製檔名" data-t="${esc(r.filename)}" style="padding:1px 6px"><i class="bi bi-clipboard"></i></button></td>
      <td class="muted">${md(r.delivery_date)}</td><td class="fn muted">${esc(r.po_number)}</td><td class="muted">${esc(r.warehouse)}</td>
      <td><span class="cat cat-${esc(r.category)}">${esc(r.category || "?")}</span></td>${v.split_by_unit ? `<td>${esc(r.unit)}</td>` : ""}
      <td><span class="cat lbl-${r.label}">${r.label === "V" ? "需貼中標" : "不貼中標"}</span></td>
      <td class="num">${r.item_count}${warn ? ` <i class="bi bi-exclamation-triangle-fill" style="color:var(--warn)" title="有品項需注意，點選展開查看"></i>` : ""}</td><td class="num"><b>${fmt(r.cases_total)}</b></td>
      <td><input class="inp eip" data-k="${esc(key)}" value="${esc(r.eip_po)}" placeholder="PO202609…" ${dis}>${r.eip_mates && r.eip_mates.length ? `<div class="kbd" title="${esc(r.eip_mates.map(m => m.filename).join("\n"))}">與「${r.eip_mates.map(m => esc(m.unit)).join("、")}」同一張 EIP 採購單</div>` : ""}</td>
      <td><input class="inp slot" data-k="${esc(key)}" value="${esc(r.slot_time)}" placeholder="例如 12:30~15:30（1台車）" ${dis}></td>
      <td class="whitespace-nowrap">${ok ? `<button class="btn btn-o btn-sm emma" data-k="${esc(key)}" data-eip="${r.eip_po ? 1 : 0}" title="${r.eip_po ? "下載 EMMA 匯入檔" : "下載 EMMA 匯入檔（尚未填 EIP 採購單號，出貨備註會留空）"}">EMMA${r.eip_po ? "" : ` <i class="bi bi-exclamation-circle" style="opacity:.8"></i>`}</button> <button class="btn btn-o btn-sm eipf" data-k="${esc(key)}" title="下載 EIP 採購單">EIP</button> <button class="btn btn-sm po ${r.po_missing.length ? "btn-o" : "btn-p"}" data-k="${esc(key)}" ${r.po_missing.length ? `disabled title="${esc(r.po_missing.join("；"))}"` : `title="${esc([r.po_filename, ...(r.po_hints || [])].join("\n"))}"`}>瑪氏採購單${r.po_hints && r.po_hints.length ? ` <i class="bi bi-exclamation-circle" style="opacity:.8"></i>` : ""}</button>${sameEip.length > 1 ? ` <button class="btn btn-o btn-sm po-merge" data-k="${esc(key)}" title="${esc(`與${sameEip.filter(x => x !== r).map(x => `「${x.unit}」`).join("、")}同一張 EIP 採購單（${r.eip_po}），合併成一張瑪氏採購單下載`)}"><i class="bi bi-union"></i> 合併</button>` : ""}` : `<span class="kbd" title="此份有品項需先處理（見上方紅色提示）">先處理紅色提示</span>`}</td></tr>`;
    if (OPEN.has(key) || matched.get(key).byItem) {
      h += `<tr class="sub"><td></td><td></td><td colspan="${v.split_by_unit ? 12 : 11}"><table class="t"><thead><tr><th>永豐料號</th><th>下採料號</th><th>品名</th><th class="num">出貨數量</th><th>單位</th><th class="num">箱入數</th><th class="num">箱數</th><th>瑪氏貨號</th><th>採購單箱備註</th><th>注意事項</th></tr></thead><tbody>
        ${r.items.map(i => `<tr><td class="fn">${esc(i.yf_sku)}</td><td class="fn">${esc(i.purchase_code)}${i.purchase_code !== i.yf_sku ? ` <span class="badge st-changed" title="報價備註與永豐料號不同（組出商品）">組出</span>` : ""}</td><td>${esc(i.product_name)}</td><td class="num">${fmt(i.qty_ship)}${i.qty_overridden ? ` <span class="badge st-generated" title="已在商品主檔自動化的訂單明細修改出貨數量">我方修改</span>` : ""}</td><td>${esc(i.unit)}</td><td class="num">${fmt(i.box_file)}</td><td class="num"><b>${fmt(i.cases)}</b></td><td class="fn">${esc(i.mars_code)}</td><td>${esc(i.po_case_note)}</td><td style="color:var(--warn)">${(i.issues || []).map(esc).join("<br>")}</td></tr>`).join("")}
      </tbody></table></td></tr>`;
    }
  }
  if (!v.splits.length) h += `<tr><td colspan="${ncol}" class="muted" style="padding:24px;text-align:center">此期間沒有瑪氏訂單。請改選到貨日，或將訂單彙總表拖曳至上方匯入。</td></tr>`;
  else if (!shown.length) {
    const whole = DATES.length && ($("#d-from").value > DATES[0] || $("#d-to").value < DATES[DATES.length - 1]);
    h += `<tr><td colspan="${ncol}" class="muted" style="padding:24px;text-align:center">此期間沒有符合的拆單表。${whole ? `<button class="btn btn-o btn-sm ml-2" id="q-all"><i class="bi bi-calendar-range"></i> 改成全部日期再找</button>` : ""}</td></tr>`;
  }
  $("#sp-table").innerHTML = h + "</tbody>";
  const qa = $("#q-all"); if (qa) qa.addEventListener("click", () => { $("#d-from").value = DATES[0]; $("#d-to").value = DATES[DATES.length - 1]; loadSplits(); });
  $("#sp-table").querySelectorAll("tr.sp").forEach(tr => tr.addEventListener("click", e => {
    if (e.target.closest("input, a, button")) return;
    const k = tr.dataset.k; OPEN.has(k) ? OPEN.delete(k) : OPEN.add(k); render();
  }));
  $("#sp-table").querySelectorAll("button.po").forEach(b => b.addEventListener("click", async () => {
    const id = await ensureSplit(b.dataset.k); if (!id) return;
    await downloadBlob(`/api/mars/splits/${id}/po`, {}, "瑪氏採購單.xlsx", "瑪氏採購單已下載");
  }));
  $("#sp-table").querySelectorAll("button.po-merge").forEach(b => b.addEventListener("click", async () => {
    const id = await ensureSplit(b.dataset.k); if (!id) return;
    await downloadBlob(`/api/mars/splits/${id}/po?merge=1`, {}, "瑪氏採購單.xlsx", "合併的瑪氏採購單已下載");
  }));
  $("#sp-table").querySelectorAll("button.emma").forEach(b => b.addEventListener("click", async () => {
    if (b.dataset.eip !== "1" && !confirm("此份尚未填 EIP 採購單號，EMMA 匯入檔的出貨備註會留空。確定要下載？")) return;
    const id = await ensureSplit(b.dataset.k); if (!id) return;
    await downloadBlob(`/api/mars/splits/${id}/emma?ack=1`, {}, "酷澎訂單匯入.xlsx", "EMMA 匯入檔已下載");
  }));
  $("#sp-table").querySelectorAll("button.eipf").forEach(b => b.addEventListener("click", async () => {
    const id = await ensureWithMates(b.dataset.k); if (!id) return;
    await downloadBlob(`/api/mars/splits/${id}/file?kind=eip`, {}, "EIP上傳.xls", "EIP 採購單已下載");
  }));
  const syncSel = () => { $("#btn-emma-sel").disabled = !CHECKED.size; $("#btn-emma-sel").innerHTML = `<i class="bi bi-check2-square"></i> 合併下載 EMMA 匯入檔${CHECKED.size ? `（勾選 ${CHECKED.size} 份）` : "（勾選）"}`; };
  $("#sp-table").querySelectorAll("input.ck").forEach(c => c.addEventListener("change", () => { const k = c.dataset.k; c.checked ? CHECKED.add(k) : CHECKED.delete(k); syncSel(); }));
  const all = $("#ck-all"); if (all) all.addEventListener("change", () => { $("#sp-table").querySelectorAll("input.ck").forEach(c => { c.checked = all.checked; const k = c.dataset.k; all.checked ? CHECKED.add(k) : CHECKED.delete(k); }); syncSel(); });
  syncSel();
  $("#sp-table").querySelectorAll(".cp").forEach(b => b.addEventListener("click", async () => { try { await navigator.clipboard.writeText(b.dataset.t); toast("檔名已複製"); } catch (e) { toast("瀏覽器不允許複製，請手動選取", "err"); } }));
  $("#sp-table").querySelectorAll("input.eip, input.slot").forEach(inp => {
    inp.dataset.orig = inp.value;
    const save = async () => {
      if (inp.value === inp.dataset.orig) return;
      const body = inp.classList.contains("eip") ? { eip_po: inp.value } : { slot_time: inp.value };
      const isEip = inp.classList.contains("eip");
      const id = await (isEip ? ensureWithMates(inp.dataset.k) : ensureSplit(inp.dataset.k)); if (!id) { inp.value = inp.dataset.orig; return; }
      try { const d = await api(`/api/mars/splits/${id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
        inp.classList.remove("bad"); inp.dataset.orig = isEip ? d.split.eip_po : d.split.slot_time; inp.value = inp.dataset.orig;
        const mates = d.mates_updated && d.mates_updated.length ? `（同一張 EIP 採購單的 ${d.mates_updated.length} 份一併${inp.value ? "填入" : "清除"}）` : "";
        toast(isEip ? (inp.value ? `已儲存 EIP 採購單號 ${inp.value}${mates}` : `已清除 EIP 採購單號${mates}`) : "約倉時間已儲存");
        if (d.warning) toast(d.warning, "err");
        loadSplits(); if (isEip) loadCalendar();
      } catch (e) { inp.classList.add("bad"); toast(e.message, "err"); }
    };
    inp.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); inp.blur(); } if (e.key === "Escape") { inp.value = inp.dataset.orig; inp.blur(); } });
    inp.addEventListener("blur", save);
  });
}

/* 還沒產出的那一份，第一次按它的按鈕或填 EIP 單號時先存起來（Jerry 2026-10-06：不用先按上面的大按鈕） */
async function ensureSplit(key, ack) {
  const r = VIEW && VIEW.splits.find(x => x.split_key === key);
  if (r && r.id) return r.id;
  try {
    const res = await fetch("/api/mars/splits/ensure", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ split_key: key, ack_unmatched: !!ack }) });
    const d = await res.json().catch(() => ({}));
    if (res.status === 409 && d.needs_ack) return confirm(`${d.error}\n\n${(d.details || []).join("\n")}\n\n確定要繼續？`) ? ensureSplit(key, true) : null;
    if (!res.ok) { toast([d.error || `伺服器錯誤（${res.status}）`, ...(d.details || []).slice(0, 5)].join("\n"), "err"); return null; }
    if (r) r.id = d.id;
    if (d.created) setTimeout(() => { loadSplits(); loadCalendar(); }, 0);   // 狀態變「已產出」，畫面重抓
    return d.id;
  } catch (e) { toast(e.message, "err"); return null; }
}
/* 同一張 EIP 採購單（盒、包合一張）的其他份也先存起來，下載 EIP 檔、填號碼時才會整張一起處理 */
async function ensureWithMates(key) {
  const r = VIEW && VIEW.splits.find(x => x.split_key === key);
  const ids = await ensureMany([key, ...((r && r.eip_mates) || []).map(m => m.split_key)]);
  return ids ? ids[0] : null;
}
async function ensureMany(keys) {
  const ids = [];
  for (const k of keys) { const id = await ensureSplit(k); if (!id) return null; ids.push(id); }
  return ids;
}

async function generate(ack) {
  const body = { from: $("#d-from").value, to: $("#d-to").value, ack_unmatched: !!ack };
  $("#btn-gen").disabled = true;
  try {
    const res = await fetch("/api/mars/splits/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!res.ok) {
      const d = await res.json().catch(() => ({}));
      if (d.needs_ack && confirm(`${d.error}\n\n${(d.details || []).join("\n")}\n\n其餘品項仍要產出？`)) return generate(true);
      if (!d.needs_ack) toast([d.error, ...(d.details || [])].join("　"), "err");
      return;
    }
    const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/); const name = m ? decodeURIComponent(m[1]) : "瑪氏EIP採購單.zip";
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    toast("EIP 採購單已下載。取得 EIP 採購單號後請回填，再下載 EMMA 匯入檔。"); loadSplits(); loadCalendar();
  } catch (e) { toast(e.message, "err"); }
  finally { $("#btn-gen").disabled = false; }
}
$("#btn-gen").addEventListener("click", () => generate(false));

/* ── ③ 瑪氏採購單：單份下載、整段期間打包 ── */
async function downloadBlob(url, opts, fallback, okMsg) {
  try {
    const res = await fetch(url, opts);
    if (!res.ok) { const d = await res.json().catch(() => ({})); toast([d.error || `伺服器錯誤（${res.status}）`, ...(d.details || []).slice(0, 5)].join("\n"), "err"); return null; }
    const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/); const name = m ? decodeURIComponent(m[1]) : fallback;
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    if (okMsg) toast(okMsg);
    return res;
  } catch (e) { toast(e.message, "err"); return null; }
}
$("#btn-po").addEventListener("click", async () => {
  $("#btn-po").disabled = true;
  try {
    const res = await downloadBlob("/api/mars/po/zip", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ from: $("#d-from").value, to: $("#d-to").value }) }, "瑪氏採購單.zip");
    if (res) { const n = res.headers.get("X-Mars-Po-Count"), sk = Number(res.headers.get("X-Mars-Po-Skipped") || 0); toast(`產出 ${n} 張瑪氏採購單` + (sk ? `，另有 ${sk} 份資料不足未產出（說明見 zip 內檔案）` : "")); }
  } finally { $("#btn-po").disabled = false; }
});

/* EMMA 合併：勾選的 / 全部。沒填 EIP 單號的先問一次 */
async function emmaMerge(body) {
  const send = async (ack) => {
    const res = await fetch("/api/mars/emma", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ...body, ack_missing_eip: ack }) });
    if (res.status === 409) { const d = await res.json(); if (confirm(`${d.error}\n\n${(d.details || []).join("\n")}\n\n確定要下載？`)) return send(true); return null; }
    if (!res.ok) { const d = await res.json().catch(() => ({})); toast(d.error || `伺服器錯誤（${res.status}）`, "err"); return null; }
    const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/); const name = m ? decodeURIComponent(m[1]) : "酷澎訂單匯入.xlsx";
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    const w = Number(res.headers.get("X-Mars-Emma-Warnings") || 0);
    toast(`EMMA 匯入檔已下載：${res.headers.get("X-Mars-Emma-Rows")} 份合併` + (w ? `，有 ${w} 處留空（電話、EIP 採購單號或單價未填）` : ""));
    return res;
  };
  try { await send(false); } catch (e) { toast(e.message, "err"); }
}
$("#btn-emma-sel").addEventListener("click", async () => { const ids = await ensureMany([...CHECKED]); if (ids) emmaMerge({ ids }); });
$("#btn-emma-all").addEventListener("click", async () => {
  // 全部：還沒產出的那幾份先存起來，再照期間合併
  const keys = (VIEW ? VIEW.splits : []).filter(r => !r.id && r.status !== "gone" && !(r.blocking && r.blocking.length)).map(r => r.split_key);
  if (keys.length && !(await ensureMany(keys))) return;
  emmaMerge({ from: $("#d-from").value, to: $("#d-to").value });
});

/* ── ⑥ 驗收單缺貨劃線（Jerry 2026-10-08）：先預覽會改哪幾品，按下載才產檔 ── */
let RC_FILES = [];
$("#rc-file").addEventListener("change", e => { const fs = [...e.target.files]; e.target.value = ""; if (fs.length) previewRc(fs); });
dropzone($("#dz-rc"), fs => previewRc(fs), { re: /\.pdf$/i, msg: "請拖曳驗收單 PDF" });
async function previewRc(files) {
  RC_FILES = files;
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  $("#rc-msg").innerHTML = `<span class="muted">讀取驗收單中…</span>`; $("#rc-result").classList.add("hidden");
  let d;
  try { d = await api("/api/mars/receipts/preview", { method: "POST", body: fd }); } catch (e) { $("#rc-msg").innerHTML = `<span class="neg">${esc(e.message)}</span>`; return; }
  $("#rc-msg").innerHTML = "";
  const marked = d.files.filter(f => f.note).length;
  let h = d.files.map(f => {
    const items = f.items.map(i => `<div class="kbd" style="padding-left:18px">No.${i.no}　${esc(i.name)}　<b>${fmt(i.pdf_qty)}</b> → ${i.kind === "strike" ? `<b style="color:var(--bad)">整列劃掉</b>` : `<b style="color:var(--bad)">${fmt(i.new_qty)}</b>`}</div>`).join("");
    const status = f.error ? `<span class="neg">${esc(f.error)}，原檔不變</span>`
      : !f.applies ? `<span class="muted">不是瑪氏的驗收單（或系統中沒有此 PO），原檔不變</span>`
      : f.note ? `<b style="color:var(--bad)">${esc(f.note)}</b>` : `<span style="color:var(--ok)">數量與系統一致，不需劃線</span>`;
    const miss = f.missing && f.missing.length ? `<div class="kbd" style="padding-left:18px;color:var(--warn)">有 ${f.missing.length} 個 SKU 在系統中找不到，不會修改：${f.missing.map(esc).join("、")}</div>` : "";
    return `<div class="mb-2"><span class="fn">${esc(f.filename)}</span>${f.po_number ? ` · <span class="fn">${esc(f.po_number)}</span>` : ""} · ${status}${items}${miss}</div>`;
  }).join("");
  h += `<div class="flex items-center gap-2 mt-3 flex-wrap"><button id="rc-go" class="btn btn-p" ${marked ? "" : "disabled"}><i class="bi bi-download"></i> 下載劃好線的驗收單（${marked} 份要修改${d.files.length > marked ? `，其他 ${d.files.length - marked} 份原檔照附` : ""}）</button>
    <button id="rc-clear" class="btn btn-o">清除並重新上傳</button></div>`;
  $("#rc-result").innerHTML = h; $("#rc-result").classList.remove("hidden");
  $("#rc-clear").addEventListener("click", () => { RC_FILES = []; $("#rc-result").classList.add("hidden"); $("#rc-result").innerHTML = ""; });
  $("#rc-go").addEventListener("click", async () => {
    const fd2 = new FormData(); RC_FILES.forEach(f => fd2.append("file", f));
    await downloadBlob("/api/mars/receipts/mark", { method: "POST", body: fd2 }, RC_FILES.length > 1 ? "驗收單缺貨劃線.zip" : RC_FILES[0].name, `已下載 ${marked} 份劃好線的驗收單`);
  });
}

/* ── ⑤ 勇信缺貨 ── */
let YX = null, YX_FILES = [], YX_SKIP = new Set();       // YX_SKIP＝取消勾選（這次不出）的酷澎 PO
const YX_ST = { ok: ["沒缺", "st-filled"], partial: ["部分缺", "st-changed"], none: ["全缺", "st-gone"], over: ["出貨超量（異常）", "st-changed_after_eip"] };
$("#yx-file").addEventListener("change", e => { const fs = [...e.target.files]; e.target.value = ""; if (fs.length) compareYx(fs); });
dropzone($("#dz-yx"), fs => compareYx(fs), { re: /\.pdf$/i, msg: "請拖曳勇信配送明細表 PDF" });
async function compareYx(files, keepSkip) {
  if (!files.length) { toast("請拖曳勇信配送明細表 PDF", "err"); return; }
  YX_FILES = files; if (!keepSkip) YX_SKIP = new Set();
  const fd = new FormData(); files.forEach(f => fd.append("file", f)); fd.append("skip_pos", [...YX_SKIP].join(","));
  // 勾勾重算時不要把結果區先藏起來：藏起來頁面會變短、瀏覽器就跳回上面（Jerry 2026-10-08）。畫面位置、清單捲到哪都記住，重畫後放回去
  const y = window.scrollY, inner = [...$("#yx-result").querySelectorAll(".tbl-wrap")].map(w => w.scrollTop);
  $("#yx-msg").innerHTML = `<span class="muted">讀取 PDF 並比對中…</span>`; if (!keepSkip) $("#yx-result").classList.add("hidden");
  try { YX = await api("/api/mars/shortage/compare", { method: "POST", body: fd }); } catch (e) { $("#yx-msg").innerHTML = `<span class="neg">${esc(e.message)}</span>`; return; }
  $("#yx-msg").innerHTML = ""; renderYx();
  if (keepSkip) { $("#yx-result").querySelectorAll(".tbl-wrap").forEach((w, i) => { w.scrollTop = inner[i] || 0; }); window.scrollTo(0, y); }
}
function renderYx() {
  const r = YX, s = r.summary;
  let h = `<div class="sum text-sm mb-2"><span><b>${s.pages}</b> 頁</span><span>對到 <b>${s.pos}</b> 張 PO</span><span><b>${s.items}</b> 個品項</span><span>沒缺 <b style="color:var(--ok)">${s.ok}</b> 項</span><span>部分缺 <b style="color:var(--warn)">${s.partial}</b> 項</span><span>全缺 <b style="color:var(--bad)">${s.none}</b> 項</span>${s.over ? `<span>出貨超量 <b style="color:var(--bad)">${s.over}</b> 項</span>` : ""}<span>短缺 <b>${fmt(s.short_cases)}</b> 箱</span>${s.full_pos ? `<span>整張不出 <b style="color:var(--bad)">${s.full_pos}</b> 張</span>` : ""}</div>`;
  const al = [];
  if (r.unknown.length) al.push(`<div class="alert al-bad"><b>勇信配送明細表中有 ${r.unknown.length} 張採購單在系統中找不到</b>（可能未回填 EIP 採購單號，或不是本系統拆的單）：${r.unknown.map(u => `${esc(u.eip_po)}（第 ${u.pages.join("、")} 頁，${u.cases} 箱）`).join("、")}</div>`);
  // 這次要出的訂單：PDF 那天那個倉的每一張酷澎 PO，預設全勾；核不到就是缺（主管 2026-10-08）
  if (r.scope.length) {
    const where = `${r.ship_dates.map(md).join("、")} · ${r.warehouses.map(esc).join("、") || "倉別不明"}`;
    al.push(`<div class="alert" style="background:var(--m0);border:1px solid var(--m2)"><b>這次要出的訂單</b>（${where}，共 ${r.scope.length} 張酷澎 PO）：勾選的才核對，<b>勇信配送明細表上核不到的算缺貨</b>；這次不出的請取消勾選。
      <div class="tbl-wrap mt-2" style="max-height:30vh"><table class="t"><thead><tr><th style="width:36px"><input type="checkbox" id="yx-scope-all" ${r.scope.every(x => !x.skipped) ? "checked" : ""} title="全選"></th><th>酷澎 PO</th><th>EIP 採購單號</th><th class="num">份數</th><th class="num">箱數</th><th>核對結果</th></tr></thead><tbody>
      ${r.scope.map(x => `<tr${x.skipped ? ' class="muted"' : ""}><td><input type="checkbox" class="yx-scope" data-po="${esc(x.po_number)}" ${x.skipped ? "" : "checked"}></td><td class="fn">${esc(x.po_number)}</td><td class="fn">${x.eip_pos.map(esc).join("、") || '<span class="neg">未填</span>'}</td><td class="num">${x.files}</td><td class="num">${fmt(x.cases)}</td>
        <td>${x.skipped ? `<span class="kbd">這次不核</span>` : [x.in_pdf ? `<span class="badge st-filled">表上有 ${x.in_pdf} 份</span>` : "", x.not_in_pdf ? `<span class="badge st-gone">表上沒有 ${x.not_in_pdf} 份，算全缺</span>` : "", x.no_eip ? `<span class="badge st-changed">${x.no_eip} 份還沒填 EIP 單號，無法核對</span>` : ""].filter(Boolean).join(" ")}</td></tr>`).join("")}
      </tbody></table></div></div>`);
  }
  if (r.no_eip.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-exclamation-octagon-fill"></i> ${r.no_eip.length} 份拆單表還沒填 EIP 採購單號，無法跟勇信配送明細表核對</b>（這幾份不修改、不列入下修檔）。請先到上方拆單表補上號碼再重新比對：<br>${r.no_eip.map(x => `${esc(x.po_number)} <span class="fn">${esc(x.filename)}</span>`).join("<br>")}</div>`);
  const extras = r.pos.flatMap(p => p.splits.flatMap(y => y.extra.map(x => `${esc(y.eip_po)}：勇信多出 ${esc(x.mars_code)} ${esc(x.name)} ${x.cases} 箱（未訂購）`)));
  if (extras.length) al.push(`<div class="alert al-warn"><b>勇信配送明細表中有未訂購的料號</b>（系統不處理，請人工確認）：<br>${extras.join("<br>")}</div>`);
  if (r.page_warnings.length) al.push(`<div class="alert al-warn">${r.page_warnings.map(esc).join("<br>")}</div>`);
  if (s.full_pos) al.push(`<div class="alert al-bad"><i class="bi bi-exclamation-octagon-fill"></i> 有 <b>${s.full_pos}</b> 張 PO 整張不出：確認後，該 PO 的出貨數量全部改為 <b>0</b>；依酷澎規定，酷澎下修檔第一個品項填 1、其他填 0，評論填「此單不出」。</div>`);
  h += al.join("");
  h += `<div class="tbl-wrap" style="max-height:52vh"><table class="t"><thead><tr><th>酷澎 PO</th><th>EIP 採購單號</th><th>永豐料號</th><th>品名</th><th class="num">訂購（箱）</th><th class="num">勇信出貨（箱）</th><th class="num">短缺（箱）</th><th>狀態</th><th class="num">出貨數量 → 改成</th></tr></thead><tbody>`;
  for (const p of r.pos) {
    h += `<tr class="pog"><td colspan="9"><i class="bi bi-receipt"></i> PO ${esc(p.po_number)}<span class="kbd">${md(p.delivery_date)} 到貨 · ${esc(p.warehouse)}${p.full ? ` · <b style="color:var(--bad)">整張不出</b>` : ""}${p.full_note && !p.full ? ` · ${esc(p.full_note)}` : ""}${p.missing_splits.length ? ` · <span style="color:var(--bad)">另有 ${p.missing_splits.length} 份還沒填 EIP 單號、無法核對</span>` : ""}</span></td></tr>`;
    for (const y of p.splits) for (const it of y.items) {
      const st = YX_ST[it.status]; const target = p.full ? 0 : it.new_qty;
      h += `<tr><td class="fn muted">${esc(p.po_number)}</td><td class="fn">${esc(y.eip_po)}${y.not_in_pdf ? `<div class="kbd" style="color:var(--bad)">表上沒有</div>` : ""}</td><td class="fn">${esc(it.yf_sku)}</td><td>${esc(it.product_name)}</td><td class="num">${fmt(it.ordered)}</td><td class="num"><b>${fmt(it.shipped)}</b></td><td class="num">${it.short ? `<span class="neg">${fmt(it.short)}</span>` : ""}</td><td><span class="badge ${st[1]}">${st[0]}</span></td><td class="num">${it.status === "over" ? `<span class="kbd">不調整</span>` : (it.status === "ok" && !p.full) ? `<span class="kbd">${fmt(it.qty_ship)}</span>` : `${fmt(it.qty_ship)} → <b>${fmt(target)}</b>`}</td></tr>`;
    }
  }
  if (!r.pos.length) h += `<tr><td colspan="9" class="muted" style="padding:16px;text-align:center">這些 PDF 中的採購單在系統裡都找不到對應的拆單表。</td></tr>`;
  h += `</tbody></table></div>`;
  const nChange = r.pos.reduce((a, p) => a + p.splits.reduce((b, y) => b + y.items.filter(it => p.full ? it.status !== "over" : (it.status === "partial" || it.status === "none")).length, 0), 0);
  h += `<div class="flex items-center gap-2 mt-3 flex-wrap">
    <button id="yx-apply" class="btn btn-p" ${!r.pos.length ? "disabled" : ""}><i class="bi bi-check2-circle"></i> 確認修改 ${nChange} 筆出貨數量並下載酷澎下修檔（${r.downgrade_rows} 列）</button>
    <button id="yx-file-only" class="btn btn-o" ${!r.downgrade_rows ? "disabled" : ""} title="不修改系統數量，只下載酷澎下修檔">只下載酷澎下修檔</button>
    <button id="yx-clear" class="btn btn-o">清除並重新上傳</button>
    <span class="kbd">修改會記入歷程紀錄（原因：缺貨）；已送 EIP 的拆單也會一併更新。</span></div>`;
  $("#yx-result").innerHTML = h; $("#yx-result").classList.remove("hidden");
  $("#yx-clear").addEventListener("click", () => { YX = null; YX_FILES = []; YX_SKIP = new Set(); $("#yx-result").classList.add("hidden"); $("#yx-result").innerHTML = ""; });
  $("#yx-result").querySelectorAll("input.yx-scope").forEach(c => c.addEventListener("change", () => { c.checked ? YX_SKIP.delete(c.dataset.po) : YX_SKIP.add(c.dataset.po); compareYx(YX_FILES, true); }));
  const sa = $("#yx-scope-all"); if (sa) sa.addEventListener("change", () => { r.scope.forEach(x => sa.checked ? YX_SKIP.delete(x.po_number) : YX_SKIP.add(x.po_number)); compareYx(YX_FILES, true); });
  $("#yx-file-only").addEventListener("click", () => downloadBlob("/api/mars/shortage/apply", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ result: YX, only_file: true }) }, "酷澎下修.xlsx", "酷澎下修檔已下載（系統數量未修改）"));
  $("#yx-apply").addEventListener("click", async () => {
    if (!confirm(`確定將 ${nChange} 筆出貨數量改為勇信實際出貨量？${s.full_pos ? `\n其中 ${s.full_pos} 張 PO 整張不出貨，將全部改為 0。` : ""}`)) return;
    $("#yx-apply").disabled = true;
    try {
      const res = await fetch("/api/mars/shortage/apply", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ result: YX }) });
      const ct = res.headers.get("Content-Type") || "";
      if (!res.ok || ct.includes("json")) { const d = await res.json().catch(() => ({})); if (!res.ok) { toast(d.error || `伺服器錯誤（${res.status}）`, "err"); $("#yx-apply").disabled = false; return; } toast(`已修改 ${d.changed} 筆出貨數量。${d.error || ""}`); }
      else {
        const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || ""; const m = cd.match(/filename\*=UTF-8''([^;]+)/);
        const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "酷澎下修.xlsx"; document.body.appendChild(a); a.click(); a.remove();
        toast(`已修改 ${res.headers.get("X-Mars-Changed")} 筆出貨數量，酷澎下修檔 ${res.headers.get("X-Mars-Rows")} 列已下載`);
      }
      YX = null; $("#yx-result").classList.add("hidden"); $("#yx-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 比對結果已確認，出貨數量已更新，請至上方拆單區查看。</span>`; loadSplits(); loadCalendar();
    } catch (e) { toast(e.message, "err"); $("#yx-apply").disabled = false; }
  });
}

/* ── 採購單設定：倉庫資料、固定文字、假日 ── */
let PS = null;
async function loadPoSettings() {
  PS = await api("/api/mars/po/settings");
  const st = PS.settings;
  $("#ps-lead").value = st.lead_days; $("#ps-contact").value = st.contact_default;
  ["CHO", "GUM", "PET"].forEach(k => $(`#ps-shelf-${k}`).value = st.shelf_req_cat[k]);
  $("#ps-em-recipient").value = st.emma_recipient; $("#ps-em-customer").value = st.emma_customer; $("#ps-em-billto").value = st.emma_billto; $("#ps-em-shipto").value = st.emma_shipto; $("#ps-em-payment").value = st.emma_payment;
  $("#ps-label").value = st.label_line; $("#ps-white").value = st.white_line; $("#ps-special").value = st.special_text; $("#ps-holidays").value = st.holidays.join("\n");
  const yrs = PS.builtin_holiday_years || [], thisY = new Date().getFullYear();
  $("#ps-builtin-years").textContent = yrs.join("、") + " 年";
  const missing = [thisY, thisY + 1].filter(y => !yrs.includes(y));
  $("#ps-builtin-warn").classList.toggle("hidden", !missing.length); $("#ps-builtin-warn").textContent = missing.length ? `${missing.join("、")} 年國定假日尚未內建，請通知系統管理員` : "";
  $("#ps-builtin-list").innerHTML = Object.entries(PS.builtin_holidays || {}).sort().map(([d, n]) => `${d.slice(5).replace("-", "/")} ${esc(n)}`).join("<br>");
  renderWarehouses(PS.warehouses);
}
function renderWarehouses(list) {
  const F = [["name", "入倉倉別（C6）", 160], ["address", "地址（C7）", 220], ["phone", "電話（C8）", 115], ["ship_to", "ship-to（F7）", 95], ["contact", "聯絡人（F8）", 85], ["special_note", "特殊需求加註", 200]];
  let h = `<thead><tr><th>倉別</th>${F.map(f => `<th style="min-width:${f[2]}px">${f[1]}</th>`).join("")}<th></th></tr></thead><tbody>`;
  for (const w of list) {
    h += `<tr data-code="${esc(w.code)}"><td class="whitespace-nowrap"><b>${esc(w.code)}</b><br>${w.missing.length ? `<span class="badge st-changed" title="缺 ${esc(w.missing.join("、"))}">缺 ${esc(w.missing.join("、"))}</span>` : `<span class="badge st-filled" title="${esc(w.updated_by)} ${esc((w.updated_at || "").slice(5, 16))}">已填齊</span>`}</td>${F.map(f => `<td><input class="inp wh ${w.missing.includes(f[1].split("（")[0]) ? "bad" : ""}" data-f="${f[0]}" value="${esc(w[f[0]])}" ${f[0] === "address" && w.address_from_orders ? 'title="此為訂單上的地址，未另外填寫時使用"' : ""}></td>`).join("")}
      <td>${w.in_orders ? `<span class="kbd" title="訂單中有此倉，無法刪除">訂單使用中</span>` : `<button class="btn btn-o btn-sm wh-del" data-code="${esc(w.code)}" title="刪除此倉（限訂單中未出現的倉）" style="padding:2px 7px"><i class="bi bi-trash"></i></button>`}</td></tr>`;
  }
  if (!list.length) h += `<tr><td colspan="8" class="muted" style="padding:12px">尚無瑪氏訂單。匯入訂單彙總表後會自動列出倉別，也可在下方直接新增。</td></tr>`;
  h += `<tr id="wh-new"><td><input class="inp" id="wh-new-code" placeholder="倉別，例如 TAO8" style="width:90px;text-transform:uppercase"></td>${F.map(f => `<td><input class="inp" data-f="${f[0]}" placeholder="${esc(f[1].split("（")[0])}"></td>`).join("")}<td><button class="btn btn-p btn-sm" id="wh-add"><i class="bi bi-plus-lg"></i> 新增</button></td></tr>`;
  $("#wh-table").innerHTML = h + "</tbody>";
  $("#wh-add").addEventListener("click", async () => {
    const body = { code: $("#wh-new-code").value }; $("#wh-new").querySelectorAll("input[data-f]").forEach(i => body[i.dataset.f] = i.value);
    try { const d = await api("/api/mars/warehouses", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      toast(`已新增 ${d.code}`); renderWarehouses(d.warehouses); loadSplits();
    } catch (e) { toast(e.message, "err"); $("#wh-new-code").classList.add("bad"); }
  });
  $("#wh-new").querySelectorAll("input").forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); $("#wh-add").click(); } }));
  $("#wh-table").querySelectorAll(".wh-del").forEach(b => b.addEventListener("click", async () => {
    if (!confirm(`確定刪除 ${b.dataset.code} 的倉庫資料？`)) return;
    try { const d = await api(`/api/mars/warehouses/${encodeURIComponent(b.dataset.code)}`, { method: "DELETE" }); toast(`已刪除 ${b.dataset.code}`); renderWarehouses(d.warehouses); loadSplits(); }
    catch (e) { toast(e.message, "err"); }
  }));
  $("#wh-table").querySelectorAll("tr[data-code]").forEach(tr => {
    const inputs = [...tr.querySelectorAll("input.wh")]; inputs.forEach(i => i.dataset.orig = i.value);
    const save = async () => {
      if (!inputs.some(i => i.value !== i.dataset.orig)) return;
      const body = {}; inputs.forEach(i => body[i.dataset.f] = i.value);
      try { const d = await api(`/api/mars/warehouses/${encodeURIComponent(tr.dataset.code)}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
        toast(`${tr.dataset.code} 的倉庫資料已儲存`); renderWarehouses(d.warehouses); loadSplits();
      } catch (e) { toast(e.message, "err"); }
    };
    tr.addEventListener("focusout", e => { if (!tr.contains(e.relatedTarget)) save(); });   // 一列填完、離開這列才存一次
    inputs.forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); i.blur(); } }));
  });
}
async function uploadWarehouses(f) {
  const fd = new FormData(); fd.append("file", f); $("#wh-msg").textContent = "讀檔中…";
  try {
    const d = await api("/api/mars/warehouses/import", { method: "POST", body: fd });
    $("#wh-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> ${esc(f.name)}：${d.rows} 個倉：新增 ${d.added} 個、更新 ${d.updated} 個、未變動 ${d.same} 個。</span>`;
    renderWarehouses(d.warehouses); toast("倉庫資料已更新"); loadSplits();
  } catch (e) { $("#wh-msg").innerHTML = `<span class="neg">${esc(e.message)}</span>`; }
}
$("#wh-file").addEventListener("change", e => { const f = e.target.files[0]; e.target.value = ""; if (f) uploadWarehouses(f); });
dropzone($("#dz-wh"), fs => uploadWarehouses(fs[0]));
$("#ps-builtin-show").addEventListener("click", e => { e.preventDefault(); $("#ps-builtin-list").classList.toggle("hidden"); });
$("#ps-save").addEventListener("click", async () => {
  const body = { lead_days: $("#ps-lead").value, contact_default: $("#ps-contact").value, label_line: $("#ps-label").value,
    white_line: $("#ps-white").value, special_text: $("#ps-special").value, holidays: $("#ps-holidays").value,
    shelf_req_cat: { CHO: $("#ps-shelf-CHO").value, GUM: $("#ps-shelf-GUM").value, PET: $("#ps-shelf-PET").value },
    emma_recipient: $("#ps-em-recipient").value, emma_customer: $("#ps-em-customer").value, emma_billto: $("#ps-em-billto").value, emma_shipto: $("#ps-em-shipto").value, emma_payment: $("#ps-em-payment").value };
  $("#ps-msg").textContent = "儲存中…";
  try { await api("/api/mars/po/settings", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }); $("#ps-msg").textContent = ""; toast("採購單設定已儲存"); await loadPoSettings(); loadSplits(); }
  catch (e) { $("#ps-msg").textContent = ""; toast(e.message, "err"); }
});

loadStatus().then(loadSplits).catch(e => toast(e.message, "err"));
loadPoSettings().catch(e => toast(e.message, "err"));
