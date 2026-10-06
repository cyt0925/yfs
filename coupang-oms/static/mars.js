/* 瑪氏出貨：商品總表上傳、拆單、回填 EIP 採購單號／約倉時間、產瑪氏採購單、採購單設定。後端在 mars/。 */
const $ = s => document.querySelector(s);
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmt = n => n == null ? "" : Number(n).toLocaleString("en-US", { maximumFractionDigits: 4 });
function toast(msg, kind) { const t = document.createElement("div"); t.className = "toast" + (kind === "err" ? " err" : ""); t.textContent = msg; ($("#toasts") || document.body).appendChild(t); setTimeout(() => t.remove(), kind === "err" ? 5000 : 2500); }
async function api(url, opts = {}) {
  const res = await fetch(url, opts); let data = {}; try { data = await res.json(); } catch (e) {}
  if (!res.ok) throw Object.assign(new Error(data.error || `伺服器錯誤 (${res.status})`), { status: res.status, data });
  return data;
}
const md = d => { const [, m, dd] = d.split("-"); return `${Number(m)}/${Number(dd)}`; };
const today = () => new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
const STATUS = { new: "還沒產出", generated: "已產出，等 EIP 單號", filled: "已回填", changed: "訂單有變，重按產出", changed_after_eip: "EIP 送出後訂單有變", gone: "訂單已沒有這份" };
let VIEW = null, OPEN = new Set(), CHECKED = new Set();

/* ── 商品總表 ── */
async function loadStatus() {
  const d = await api("/api/mars/status");
  const u = d.last_upload;
  $("#mp-status").innerHTML = u ? `上次上傳 ${esc(u.uploaded_at.slice(0, 16))} · ${esc(u.operator)} · ${esc(u.filename)} · <b style="color:var(--m8)">${d.codes}</b> 個料號`
    : `<span class="neg">還沒上傳</span>：拆單要靠它分品類、中標，先把 Alice 的「自動化_Mars整合商品資料」傳上來`;
  $("#od-status").innerHTML = d.dates.length ? `② 裡有瑪氏訂單的到貨日：${d.dates.slice(-6).map(md).join("、")}${d.dates.length > 6 ? " …" : ""}${d.last_order_import ? `<br>最近一次匯入 ${esc(d.last_order_import.committed_at.slice(0, 16))} · ${esc(d.last_order_import.operator)} · ${esc(d.last_order_import.filename)}` : ""}`
    : `<span class="neg">還沒有瑪氏的訂單</span>：把 Kate 系統匯出的出貨彙總表傳上來`;
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
  $("#cal-title").textContent = `${Y} 年 ${M} 月 到貨`;
  const first = new Date(Y, M - 1, 1), days = new Date(Y, M, 0).getDate();
  let h = ["日", "一", "二", "三", "四", "五", "六"].map(w => `<div class="wd">${w}</div>`).join("");
  for (let i = 0; i < first.getDay(); i++) h += `<div class="day blank"></div>`;
  for (let d = 1; d <= days; d++) {
    const key = `${Y}-${String(M).padStart(2, "0")}-${String(d).padStart(2, "0")}`, c = CAL.days[key];
    const on = f && t && key >= f && key <= t;
    h += `<div class="day ${c ? "has" : ""} ${on ? "on" : ""} ${key === tod ? "today" : ""}" data-date="${key}"><div class="n">${d}</div>` + (c ? `
      <span class="eip ${c.filled === c.files ? "ok" : c.filled ? "part" : ""}" title="填了 EIP 單號的份數">${c.filled}／${c.files}</span>
      <div class="c">${fmt(c.cases)} <span style="font-size:10.5px;font-weight:500">箱</span></div><div class="p">${c.pos} 張 PO · ${c.files} 份${c.gone ? ` · ${c.gone} 份舊` : ""}</div>
      ${c.unmatched || c.blocking ? `<span class="warn" title="${c.unmatched ? `${c.unmatched} 個品項對不到商品總表` : ""}${c.blocking ? ` ${c.blocking} 份箱數不是整數` : ""}"><i class="bi bi-exclamation-triangle-fill"></i></span>` : ""}` : "") + `</div>`;
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
    $("#mp-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 讀到 ${d.codes} 個料號（${d.rows} 列，${esc(cats)}）${d.added || d.removed ? `，跟上一版比：新增 ${d.added}、拿掉 ${d.removed}` : ""}。</span>`
      + (d.warnings.length ? `<details class="kbd"><summary>${d.warnings.length} 列有問題（跳過或標出來）</summary>${d.warnings.map(esc).join("<br>")}</details>` : "");
    toast("商品總表已更新"); await loadStatus(); loadSplits();
  } catch (err) { $("#mp-msg").innerHTML = `<span class="neg">${esc(err.message)}</span>`; }
}
/* 拖檔進卡片就等於按上傳 */
function dropzone(el, onFiles) {
  ["dragenter", "dragover"].forEach(ev => el.addEventListener(ev, e => { e.preventDefault(); el.classList.add("over"); }));
  ["dragleave", "drop"].forEach(ev => el.addEventListener(ev, e => { e.preventDefault(); el.classList.remove("over"); }));
  el.addEventListener("drop", e => { const fs = [...e.dataTransfer.files].filter(f => /\.xlsx?$/i.test(f.name)); if (fs.length) onFiles(fs); else toast("要拖 Excel 檔（.xlsx）", "err"); });
}
dropzone($("#dz-mp"), fs => uploadMaster(fs[0]));
dropzone($("#dz-od"), fs => previewOrders(fs));

/* ── 出貨彙總表：走 ② 訂單明細同一套匯入（預覽 → 確認），只是從這頁傳 ── */
let ODP = null;
$("#od-file").addEventListener("change", e => { const files = [...e.target.files]; e.target.value = ""; if (files.length) previewOrders(files); });
async function previewOrders(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  $("#od-msg").innerHTML = `<span class="muted">讀檔中…</span>`; $("#od-preview").classList.add("hidden");
  try { ODP = await api("/api/master/import/preview", { method: "POST", body: fd }); } catch (err) { $("#od-msg").innerHTML = `<span class="neg">${esc(err.message)}</span>`; return; }
  $("#od-msg").innerHTML = "";
  const mars = ODP.lines["瑪氏"] || 0, others = Object.entries(ODP.lines).filter(([k]) => k !== "瑪氏");
  if (!mars) {
    $("#od-preview").innerHTML = `<div class="alert al-bad"><b>這份檔裡沒有瑪氏的單</b>（${others.map(([k, v]) => `${esc(k)} ${v} 列`).join("、") || "0 列"}），這頁只收瑪氏的出貨彙總表，沒有匯入。寶僑的單請到商品主檔自動化 ② 傳。</div>`;
    $("#od-preview").classList.remove("hidden"); ODP = null; return;
  }
  $("#od-preview").innerHTML = `<div style="border:1px solid var(--m2);border-radius:10px;padding:10px 12px;background:var(--m0)">
    <div><b>${esc(ODP.filename)}</b>：${ODP.rows_total} 列、${ODP.po_count} 張 PO，瑪氏 <b style="color:var(--m8)">${mars}</b> 列${others.length ? `，另有 ${others.map(([k, v]) => `${esc(k)} ${v} 列`).join("、")}（會一起進 ②，那邊的人也看得到）` : ""}。</div>
    <div class="mt-1">新增 <b>${ODP.new_count}</b> 筆、有變 <b>${ODP.updated_count}</b> 筆、沒變 ${ODP.identical_count} 筆${ODP.removed_count ? `、檔案裡消失 <span class="neg">${ODP.removed_count}</span> 筆（出貨會歸 0）` : ""}；到貨日 ${ODP.dates.map(md).join("、")}。</div>
    ${ODP.warnings.length ? `<div class="mt-1" style="color:var(--warn)">${ODP.warnings.map(esc).join("<br>")}</div>` : ""}
    <div class="kbd mt-1">在 ② 人工改過的出貨數量、交貨日不會被蓋掉；只動檔案裡有的 PO。</div>
    <div class="flex gap-2 mt-2"><button id="od-commit" class="btn btn-p btn-sm"><i class="bi bi-check2"></i> 確認匯入</button><button id="od-cancel" class="btn btn-o btn-sm">取消</button></div></div>`;
  $("#od-preview").classList.remove("hidden");
  $("#od-cancel").onclick = () => { $("#od-preview").classList.add("hidden"); ODP = null; };
  $("#od-commit").onclick = async () => {
    $("#od-commit").disabled = true;
    try { const d = await api("/api/master/import/commit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ batch_id: ODP.batch_id, add_missing_products: true }) });
      $("#od-preview").classList.add("hidden"); $("#od-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 匯入完成：新增 ${d.inserted}、更新 ${d.updated}、沒變 ${d.identical}${d.removed ? `、消失歸 0 ${d.removed}` : ""}。</span>`;
      toast("出貨彙總表匯入完成，已跳到這批的到貨日"); const ds = (ODP.dates || []).slice().sort(); ODP = null; CAL_MONTH = ds.length ? ds[0].slice(0, 7) : CAL_MONTH; await loadStatus();
      if (ds.length) { $("#d-from").value = ds[0]; $("#d-to").value = ds[ds.length - 1]; }     // 匯完直接跳到這批的到貨日
      loadSplits();
    } catch (err) { toast(err.message, "err"); $("#od-commit").disabled = false; }
  };
}

/* ── 拆單 ── */
["#d-from", "#d-to"].forEach(s => $(s).addEventListener("change", loadSplits));
async function loadSplits() {
  const f = $("#d-from").value, t = $("#d-to").value; if (!f || !t) return;
  if (CAL && f.slice(0, 7) !== CAL_MONTH && f.slice(0, 7) === t.slice(0, 7)) { CAL_MONTH = f.slice(0, 7); loadCalendar(); } else renderCalendar();
  try { VIEW = await api(`/api/mars/splits?from=${f}&to=${t}`); } catch (e) { toast(e.message, "err"); return; }
  render();
}
function render() {
  const v = VIEW, s = v.summary;
  $("#sum").innerHTML = `<span><b>${s.pos}</b> 張 PO</span><span><b>${s.files}</b> 份拆單表</span><span><b>${fmt(s.cases)}</b> 箱</span><span><b>${s.items}</b> 個品項</span><span><b>${s.filled}</b>／${s.files} 已回填 EIP 單號</span><span><b>${s.po_ready}</b>／${s.files} 可產瑪氏採購單</span>`;
  const al = [];
  if (!v.has_products) al.push(`<div class="alert al-bad"><i class="bi bi-exclamation-octagon-fill"></i> 還沒上傳瑪氏商品總表，分不出品類和中標，拆不了。</div>`);
  if (v.unmatched.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-exclamation-octagon-fill"></i> ${v.unmatched.length} 個品項對不到瑪氏商品總表</b>，這些不會進拆單表，請更新商品總表再重拆：<br>${v.unmatched.map(u => `${esc(u.po_number)} · ${esc(u.yf_sku)}${u.quote_note && u.quote_note !== u.yf_sku ? `（報價備註 ${esc(u.quote_note)}）` : ""} · ${esc(u.product_name)} · ${esc(u.unit)} ${fmt(u.qty_ship)}`).join("<br>")}</div>`);
  const blocking = v.splits.flatMap(r => r.blocking || []);
  if (blocking.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-slash-circle"></i> 這幾個品項要先處理才能產出</b>（EIP 採購表只能填整數箱）：<br>${blocking.map(esc).join("<br>")}</div>`);
  if (v.zero_rows) al.push(`<div class="kbd">出貨數量是 0 的 ${v.zero_rows} 個品項沒有拆進來。</div>`);
  if (v.wh_missing && v.wh_missing.length) al.push(`<div class="kbd"><i class="bi bi-geo-alt"></i> 倉庫資料還沒填齊：<b>${v.wh_missing.map(esc).join("、")}</b>。採購單照樣能產，缺的格子會留空；到下面「採購單設定」填一次就不用每次補。</div>`);
  $("#ps-warn").classList.toggle("hidden", !(v.wh_missing && v.wh_missing.length)); $("#ps-warn").textContent = v.wh_missing && v.wh_missing.length ? `${v.wh_missing.join("、")} 還沒填齊` : "";
  $("#alerts").innerHTML = al.join("");
  $("#btn-gen").disabled = !v.splits.length || !v.has_products || blocking.length > 0;
  $("#btn-po").disabled = !s.po_ready;
  const saved = v.splits.filter(r => r.id).map(r => r.id); CHECKED = new Set([...CHECKED].filter(id => saved.includes(id)));
  $("#btn-emma-all").disabled = !saved.length;

  let h = `<thead><tr><th><input type="checkbox" id="ck-all" title="全選（只選已產出的）"></th><th>狀態</th><th>拆單表檔名</th><th>到貨日</th><th>酷澎 PO</th><th>倉</th><th>品類</th>${v.split_by_unit ? "<th>單位</th>" : ""}<th>中標</th><th class="num">品項</th><th class="num">箱數</th><th style="min-width:150px">EIP 採購單號</th><th style="min-width:170px">約倉時間</th><th>下載</th></tr></thead><tbody>`;
  const ncol = v.split_by_unit ? 14 : 13;
  let lastPo = null, gi = 0;
  for (const r of v.splits) {
    const key = r.split_key, warn = r.items.some(i => i.issues && i.issues.length);
    const dis = r.id ? "" : "disabled title=\"先按「產出 EMMA 與 EIP 檔」才能填\"";
    if (r.po_number !== lastPo) {
      lastPo = r.po_number; gi++;
      const same = v.splits.filter(x => x.po_number === r.po_number);
      const filled = same.filter(x => x.eip_po).length;
      h += `<tr class="pog"><td colspan="${ncol}"><i class="bi bi-receipt"></i> PO ${esc(r.po_number)}<span class="kbd">${md(r.delivery_date)} 到貨 · ${esc(r.warehouse)} · 拆成 ${same.length} 份 · ${fmt(same.reduce((a, x) => a + (x.cases_total || 0), 0))} 箱 · 已回填 ${filled}／${same.length}</span></td></tr>`;
    }
    h += `<tr class="sp ${gi % 2 ? "g1" : "g0"}" data-k="${esc(key)}">
      <td>${r.id ? `<input type="checkbox" class="ck" data-id="${r.id}" ${CHECKED.has(r.id) ? "checked" : ""} title="勾起來合併成 EMMA 檔">` : ""}</td>
      <td><span class="badge st-${r.status}">${STATUS[r.status]}</span>${r.diff ? `<div class="kbd" style="max-width:220px;color:var(--bad)">${esc(r.diff)}</div>` : ""}</td>
      <td><span class="fn fname">${esc(r.filename)}</span> <button class="btn btn-o btn-sm cp" title="複製檔名" data-t="${esc(r.filename)}" style="padding:1px 6px"><i class="bi bi-clipboard"></i></button></td>
      <td class="muted">${md(r.delivery_date)}</td><td class="fn muted">${esc(r.po_number)}</td><td class="muted">${esc(r.warehouse)}</td>
      <td><span class="cat cat-${esc(r.category)}">${esc(r.category || "?")}</span></td>${v.split_by_unit ? `<td>${esc(r.unit)}</td>` : ""}
      <td><span class="cat lbl-${r.label}">${r.label === "V" ? "需貼中標" : "不貼中標"}</span></td>
      <td class="num">${r.item_count}${warn ? ` <i class="bi bi-exclamation-triangle-fill" style="color:var(--warn)" title="有品項要注意，點開看"></i>` : ""}</td><td class="num"><b>${fmt(r.cases_total)}</b></td>
      <td><input class="inp eip" data-id="${r.id || ""}" value="${esc(r.eip_po)}" placeholder="PO202609…" ${dis}></td>
      <td><input class="inp slot" data-id="${r.id || ""}" value="${esc(r.slot_time)}" placeholder="例如 12:30~15:30（1台車）" ${dis}></td>
      <td class="whitespace-nowrap">${r.id ? `<button class="btn btn-o btn-sm emma" data-id="${r.id}" data-eip="${r.eip_po ? 1 : 0}" title="${r.eip_po ? "EMMA 匯入檔" : "EMMA 匯入檔（EIP 採購單號還沒填，出貨備註會空著）"}">EMMA${r.eip_po ? "" : ` <i class="bi bi-exclamation-circle" style="opacity:.8"></i>`}</button> <a class="btn btn-o btn-sm" href="/api/mars/splits/${r.id}/file?kind=eip" title="EIP 上傳用採購表">EIP</a> <button class="btn btn-sm po ${r.po_missing.length ? "btn-o" : "btn-p"}" data-id="${r.id}" ${r.po_missing.length ? `disabled title="${esc(r.po_missing.join("；"))}"` : `title="${esc([r.po_filename, ...(r.po_hints || [])].join("\n"))}"`}>採購單${r.po_hints && r.po_hints.length ? ` <i class="bi bi-exclamation-circle" style="opacity:.8"></i>` : ""}</button>` : `<span class="kbd">—</span>`}</td></tr>`;
    if (OPEN.has(key)) {
      h += `<tr class="sub"><td></td><td></td><td colspan="${v.split_by_unit ? 12 : 11}"><table class="t"><thead><tr><th>永豐料號</th><th>下採料號</th><th>品名</th><th class="num">出貨數量</th><th>單位</th><th class="num">箱入數</th><th class="num">箱數</th><th>瑪氏貨號</th><th>採購單箱備註</th><th>要注意</th></tr></thead><tbody>
        ${r.items.map(i => `<tr><td class="fn">${esc(i.yf_sku)}</td><td class="fn">${esc(i.purchase_code)}${i.purchase_code !== i.yf_sku ? ` <span class="badge st-changed" title="報價備註跟永豐料號不同（組出商品）">組出</span>` : ""}</td><td>${esc(i.product_name)}</td><td class="num">${fmt(i.qty_ship)}${i.qty_overridden ? ` <span class="badge st-generated" title="在 ② 訂單明細人工改過">改過</span>` : ""}</td><td>${esc(i.unit)}</td><td class="num">${fmt(i.box_file)}</td><td class="num"><b>${fmt(i.cases)}</b></td><td class="fn">${esc(i.mars_code)}</td><td>${esc(i.po_case_note)}</td><td style="color:var(--warn)">${(i.issues || []).map(esc).join("<br>")}</td></tr>`).join("")}
      </tbody></table></td></tr>`;
    }
  }
  if (!v.splits.length) h += `<tr><td colspan="${ncol}" class="muted" style="padding:24px;text-align:center">這段期間沒有瑪氏的訂單。換個到貨日，或把出貨彙總表拖到上面的框匯進來。</td></tr>`;
  $("#sp-table").innerHTML = h + "</tbody>";
  $("#sp-table").querySelectorAll("tr.sp").forEach(tr => tr.addEventListener("click", e => {
    if (e.target.closest("input, a, button")) return;
    const k = tr.dataset.k; OPEN.has(k) ? OPEN.delete(k) : OPEN.add(k); render();
  }));
  $("#sp-table").querySelectorAll("button.po").forEach(b => b.addEventListener("click", () => downloadBlob(`/api/mars/splits/${b.dataset.id}/po`, {}, "瑪氏採購單.xlsx", "瑪氏採購單已下載")));
  $("#sp-table").querySelectorAll("button.emma").forEach(b => b.addEventListener("click", () => {
    if (b.dataset.eip !== "1" && !confirm("這份還沒填 EIP 採購單號，EMMA 檔的出貨備註會空著。還是要下載？")) return;
    downloadBlob(`/api/mars/splits/${b.dataset.id}/emma?ack=1`, {}, "酷澎訂單匯入.xlsx", "EMMA 匯入檔已下載");
  }));
  const syncSel = () => { $("#btn-emma-sel").disabled = !CHECKED.size; $("#btn-emma-sel").innerHTML = `<i class="bi bi-check2-square"></i> 勾選的合併成 EMMA 檔${CHECKED.size ? `（${CHECKED.size}）` : ""}`; };
  $("#sp-table").querySelectorAll("input.ck").forEach(c => c.addEventListener("change", () => { const id = Number(c.dataset.id); c.checked ? CHECKED.add(id) : CHECKED.delete(id); syncSel(); }));
  const all = $("#ck-all"); if (all) all.addEventListener("change", () => { $("#sp-table").querySelectorAll("input.ck").forEach(c => { c.checked = all.checked; const id = Number(c.dataset.id); all.checked ? CHECKED.add(id) : CHECKED.delete(id); }); syncSel(); });
  syncSel();
  $("#sp-table").querySelectorAll(".cp").forEach(b => b.addEventListener("click", async () => { try { await navigator.clipboard.writeText(b.dataset.t); toast("檔名已複製"); } catch (e) { toast("瀏覽器不讓複製，請手動選取", "err"); } }));
  $("#sp-table").querySelectorAll("input.eip, input.slot").forEach(inp => {
    inp.dataset.orig = inp.value;
    const save = async () => {
      if (!inp.dataset.id || inp.value === inp.dataset.orig) return;
      const body = inp.classList.contains("eip") ? { eip_po: inp.value } : { slot_time: inp.value };
      try { const d = await api(`/api/mars/splits/${inp.dataset.id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
        inp.classList.remove("bad"); inp.dataset.orig = inp.classList.contains("eip") ? d.split.eip_po : d.split.slot_time; inp.value = inp.dataset.orig;
        toast(inp.classList.contains("eip") ? (inp.value ? `已存 ${inp.value}` : "EIP 採購單號已清掉") : "約倉時間已存"); loadSplits(); if (inp.classList.contains("eip")) loadCalendar();
      } catch (e) { inp.classList.add("bad"); toast(e.message, "err"); }
    };
    inp.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); inp.blur(); } if (e.key === "Escape") { inp.value = inp.dataset.orig; inp.blur(); } });
    inp.addEventListener("blur", save);
  });
}

async function generate(ack) {
  const body = { from: $("#d-from").value, to: $("#d-to").value, ack_unmatched: !!ack };
  $("#btn-gen").disabled = true;
  try {
    const res = await fetch("/api/mars/splits/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!res.ok) {
      const d = await res.json().catch(() => ({}));
      if (d.needs_ack && confirm(`${d.error}\n\n${(d.details || []).join("\n")}\n\n其他的照樣產出？`)) return generate(true);
      if (!d.needs_ack) toast([d.error, ...(d.details || [])].join("　"), "err");
      return;
    }
    const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/); const name = m ? decodeURIComponent(m[1]) : "瑪氏拆單.zip";
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    toast("已產出 EMMA 與 EIP 檔，EIP 拿到單號後填回來"); loadSplits(); loadCalendar();
  } catch (e) { toast(e.message, "err"); }
  finally { $("#btn-gen").disabled = false; }
}
$("#btn-gen").addEventListener("click", () => generate(false));

/* ── ③ 瑪氏採購單：單份下載、整段期間打包 ── */
async function downloadBlob(url, opts, fallback, okMsg) {
  try {
    const res = await fetch(url, opts);
    if (!res.ok) { const d = await res.json().catch(() => ({})); toast([d.error || `伺服器錯誤 (${res.status})`, ...(d.details || []).slice(0, 5)].join("\n"), "err"); return null; }
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
    if (res) { const n = res.headers.get("X-Mars-Po-Count"), sk = Number(res.headers.get("X-Mars-Po-Skipped") || 0); toast(`產出 ${n} 張瑪氏採購單` + (sk ? `，另外 ${sk} 份還差東西沒產（zip 裡有說明）` : "")); }
  } finally { $("#btn-po").disabled = false; }
});

/* EMMA 合併：勾選的 / 全部。沒填 EIP 單號的先問一次 */
async function emmaMerge(body) {
  const send = async (ack) => {
    const res = await fetch("/api/mars/emma", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ ...body, ack_missing_eip: ack }) });
    if (res.status === 409) { const d = await res.json(); if (confirm(`${d.error}\n\n${(d.details || []).join("\n")}\n\n還是要下載？`)) return send(true); return null; }
    if (!res.ok) { const d = await res.json().catch(() => ({})); toast(d.error || `伺服器錯誤 (${res.status})`, "err"); return null; }
    const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || "";
    const m = cd.match(/filename\*=UTF-8''([^;]+)/); const name = m ? decodeURIComponent(m[1]) : "酷澎訂單匯入.xlsx";
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove();
    const w = Number(res.headers.get("X-Mars-Emma-Warnings") || 0);
    toast(`EMMA 匯入檔已下載：${res.headers.get("X-Mars-Emma-Rows")} 份合併` + (w ? `，有 ${w} 個地方留空（電話或單號沒填）` : ""));
    return res;
  };
  try { await send(false); } catch (e) { toast(e.message, "err"); }
}
$("#btn-emma-sel").addEventListener("click", () => emmaMerge({ ids: [...CHECKED] }));
$("#btn-emma-all").addEventListener("click", () => emmaMerge({ from: $("#d-from").value, to: $("#d-to").value }));

/* ── ⑤ 勇信缺貨 ── */
let YX = null, YX_FILES = [], YX_NONE = new Set();
const YX_ST = { ok: ["沒缺", "st-filled"], partial: ["部分缺", "st-changed"], none: ["全缺", "st-gone"], over: ["出太多，異常", "st-changed_after_eip"] };
$("#yx-file").addEventListener("change", e => { const fs = [...e.target.files]; e.target.value = ""; if (fs.length) compareYx(fs); });
dropzone($("#dz-yx"), fs => compareYx(fs.filter(f => /\.pdf$/i.test(f.name))));
async function compareYx(files, keepNone) {
  if (!files.length) { toast("要拖勇信的 PDF 檔", "err"); return; }
  YX_FILES = files; if (!keepNone) YX_NONE = new Set();
  const fd = new FormData(); files.forEach(f => fd.append("file", f)); fd.append("none_ids", [...YX_NONE].join(","));
  $("#yx-msg").innerHTML = `<span class="muted">讀 PDF、比對中…</span>`; $("#yx-result").classList.add("hidden");
  try { YX = await api("/api/mars/shortage/compare", { method: "POST", body: fd }); } catch (e) { $("#yx-msg").innerHTML = `<span class="neg">${esc(e.message)}</span>`; return; }
  $("#yx-msg").innerHTML = ""; renderYx();
}
function renderYx() {
  const r = YX, s = r.summary;
  let h = `<div class="sum text-sm mb-2"><span><b>${s.pages}</b> 頁</span><span><b>${s.pos}</b> 張 PO 對到</span><span><b>${s.items}</b> 個品項</span><span><b style="color:var(--ok)">${s.ok}</b> 沒缺</span><span><b style="color:var(--warn)">${s.partial}</b> 部分缺</span><span><b style="color:var(--bad)">${s.none}</b> 全缺</span>${s.over ? `<span><b style="color:var(--bad)">${s.over}</b> 出太多</span>` : ""}<span>缺 <b>${fmt(s.short_cases)}</b> 箱</span>${s.full_pos ? `<span><b style="color:var(--bad)">${s.full_pos}</b> 張整張不出</span>` : ""}</div>`;
  const al = [];
  if (r.unknown.length) al.push(`<div class="alert al-bad"><b>勇信表裡有 ${r.unknown.length} 張採購單系統裡沒有</b>（EIP 單號沒填回來，或不是這系統拆的）：${r.unknown.map(u => `${esc(u.eip_po)}（第 ${u.pages.join("、")} 頁，${u.cases} 箱）`).join("、")}</div>`);
  if (r.not_in_pdf.length) al.push(`<div class="alert al-warn"><b>${r.not_in_pdf.length} 份同倉同日的拆單表沒在這次的勇信表裡</b>，先不判缺貨。勇信確定整份沒出的，勾起來就當 0 箱算：<br>${r.not_in_pdf.map(x => `<label class="block"><input type="checkbox" class="yx-none" data-id="${x.split_id}"> ${esc(x.eip_po)} <span class="fn">${esc(x.filename)}</span></label>`).join("")}</div>`);
  if (r.none_ids && r.none_ids.length) al.push(`<div class="kbd">你勾了 ${r.none_ids.length} 份「勇信沒出」，已當 0 箱算進去（下面標全缺的那些）。</div>`);
  const extras = r.pos.flatMap(p => p.splits.flatMap(y => y.extra.map(x => `${esc(y.eip_po)}：勇信多了 ${esc(x.mars_code)} ${esc(x.name)} ${x.cases} 箱，我們沒訂`)));
  if (extras.length) al.push(`<div class="alert al-warn"><b>勇信表裡有我們沒訂的料號</b>（不處理，請人看）：<br>${extras.join("<br>")}</div>`);
  if (r.page_warnings.length) al.push(`<div class="alert al-warn">${r.page_warnings.map(esc).join("<br>")}</div>`);
  if (s.full_pos) al.push(`<div class="alert al-bad"><i class="bi bi-exclamation-octagon-fill"></i> 有 <b>${s.full_pos}</b> 張 PO 整張不出：按確認後系統裡那張的出貨數量全部改 <b>0</b>；下修檔照酷澎的規矩第一支填 1、其他 0，評論寫「此單不出」。</div>`);
  h += al.join("");
  h += `<div class="tbl-wrap" style="max-height:52vh"><table class="t"><thead><tr><th>酷澎 PO</th><th>EIP 單號</th><th>永豐料號</th><th>品名</th><th class="num">訂（箱）</th><th class="num">勇信出（箱）</th><th class="num">缺（箱）</th><th>狀態</th><th class="num">出貨數量 → 改成</th></tr></thead><tbody>`;
  for (const p of r.pos) {
    h += `<tr class="pog"><td colspan="9"><i class="bi bi-receipt"></i> PO ${esc(p.po_number)}<span class="kbd">${md(p.delivery_date)} 到貨 · ${esc(p.warehouse)}${p.full ? ` · <b style="color:var(--bad)">整張不出</b>` : ""}${p.full_note && !p.full ? ` · ${esc(p.full_note)}` : ""}${p.missing_splits.length ? ` · 還有 ${p.missing_splits.length} 份沒在表裡` : ""}</span></td></tr>`;
    for (const y of p.splits) for (const it of y.items) {
      const st = YX_ST[it.status]; const target = p.full ? 0 : it.new_qty;
      h += `<tr><td class="fn muted">${esc(p.po_number)}</td><td class="fn">${esc(y.eip_po)}</td><td class="fn">${esc(it.yf_sku)}</td><td>${esc(it.product_name)}</td><td class="num">${fmt(it.ordered)}</td><td class="num"><b>${fmt(it.shipped)}</b></td><td class="num">${it.short ? `<span class="neg">${fmt(it.short)}</span>` : ""}</td><td><span class="badge ${st[1]}">${st[0]}</span></td><td class="num">${it.status === "over" ? `<span class="kbd">不動</span>` : (it.status === "ok" && !p.full) ? `<span class="kbd">${fmt(it.qty_ship)}</span>` : `${fmt(it.qty_ship)} → <b>${fmt(target)}</b>`}</td></tr>`;
    }
  }
  if (!r.pos.length) h += `<tr><td colspan="9" class="muted" style="padding:16px;text-align:center">這幾份 PDF 裡的採購單都對不到系統的拆單表。</td></tr>`;
  h += `</tbody></table></div>`;
  const nChange = r.pos.reduce((a, p) => a + p.splits.reduce((b, y) => b + y.items.filter(it => p.full ? it.status !== "over" : (it.status === "partial" || it.status === "none")).length, 0), 0);
  h += `<div class="flex items-center gap-2 mt-3 flex-wrap">
    <button id="yx-apply" class="btn btn-p" ${!r.pos.length ? "disabled" : ""}><i class="bi bi-check2-circle"></i> 確認：系統出貨數量改 ${nChange} 筆，並下載下修檔（${r.downgrade_rows} 列）</button>
    <button id="yx-file-only" class="btn btn-o" ${!r.downgrade_rows ? "disabled" : ""} title="不改系統，只拿下修檔">只下載下修檔</button>
    <button id="yx-clear" class="btn btn-o">清掉重傳</button>
    <span class="kbd">改數量會記修改歷程（原因：缺貨），已送 EIP 的那幾份也會跟著改，之後 EMMA 檔、採購單都是新數字。</span></div>`;
  $("#yx-result").innerHTML = h; $("#yx-result").classList.remove("hidden");
  $("#yx-clear").addEventListener("click", () => { YX = null; YX_FILES = []; YX_NONE = new Set(); $("#yx-result").classList.add("hidden"); $("#yx-result").innerHTML = ""; });
  $("#yx-result").querySelectorAll("input.yx-none").forEach(c => c.addEventListener("change", () => { const id = Number(c.dataset.id); c.checked ? YX_NONE.add(id) : YX_NONE.delete(id); compareYx(YX_FILES, true); }));
  $("#yx-file-only").addEventListener("click", () => downloadBlob("/api/mars/shortage/apply", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ result: YX, only_file: true }) }, "酷澎下修.xlsx", "下修檔已下載（系統沒改）"));
  $("#yx-apply").addEventListener("click", async () => {
    if (!confirm(`要把系統裡 ${nChange} 筆訂單的出貨數量改成勇信實際出的？${s.full_pos ? `\n其中 ${s.full_pos} 張 PO 整張不出，會全部改 0。` : ""}\n改了會記歷程，之後 EMMA 檔、採購單都是新數字。`)) return;
    $("#yx-apply").disabled = true;
    try {
      const res = await fetch("/api/mars/shortage/apply", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ result: YX }) });
      const ct = res.headers.get("Content-Type") || "";
      if (!res.ok || ct.includes("json")) { const d = await res.json().catch(() => ({})); if (!res.ok) { toast(d.error || `伺服器錯誤 (${res.status})`, "err"); $("#yx-apply").disabled = false; return; } toast(`系統出貨數量改了 ${d.changed} 筆。${d.error || ""}`); }
      else {
        const blob = await res.blob(); const cd = res.headers.get("Content-Disposition") || ""; const m = cd.match(/filename\*=UTF-8''([^;]+)/);
        const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "酷澎下修.xlsx"; document.body.appendChild(a); a.click(); a.remove();
        toast(`系統出貨數量改了 ${res.headers.get("X-Mars-Changed")} 筆，下修檔 ${res.headers.get("X-Mars-Rows")} 列已下載`);
      }
      YX = null; $("#yx-result").classList.add("hidden"); $("#yx-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 這批比對已確認，系統數量已更新。要看結果到上面拆單表那區。</span>`; loadSplits(); loadCalendar();
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
  $("#ps-builtin-warn").classList.toggle("hidden", !missing.length); $("#ps-builtin-warn").textContent = missing.length ? `${missing.join("、")} 年的還沒加，請跟系統負責人說` : "";
  $("#ps-builtin-list").innerHTML = Object.entries(PS.builtin_holidays || {}).sort().map(([d, n]) => `${d.slice(5).replace("-", "/")} ${esc(n)}`).join("<br>");
  renderWarehouses(PS.warehouses);
}
function renderWarehouses(list) {
  const F = [["name", "入倉倉別（C6）", 160], ["address", "地址（C7）", 220], ["phone", "電話（C8）", 115], ["ship_to", "ship-to（F7）", 95], ["contact", "聯絡人（F8）", 85], ["special_note", "特殊需求加註", 200]];
  let h = `<thead><tr><th>倉</th>${F.map(f => `<th style="min-width:${f[2]}px">${f[1]}</th>`).join("")}<th></th></tr></thead><tbody>`;
  for (const w of list) {
    h += `<tr data-code="${esc(w.code)}"><td class="whitespace-nowrap"><b>${esc(w.code)}</b><br>${w.missing.length ? `<span class="badge st-changed" title="缺 ${esc(w.missing.join("、"))}">缺 ${esc(w.missing.join("、"))}</span>` : `<span class="badge st-filled" title="${esc(w.updated_by)} ${esc((w.updated_at || "").slice(5, 16))}">齊了</span>`}</td>${F.map(f => `<td><input class="inp wh ${w.missing.includes(f[1].split("（")[0]) ? "bad" : ""}" data-f="${f[0]}" value="${esc(w[f[0]])}" ${f[0] === "address" && w.address_from_orders ? 'title="這是訂單上的地址，沒另外填就用它"' : ""}></td>`).join("")}
      <td>${w.in_orders ? `<span class="kbd" title="訂單裡有這個倉，不能刪">訂單有</span>` : `<button class="btn btn-o btn-sm wh-del" data-code="${esc(w.code)}" title="刪掉這個倉（訂單裡沒出現過的才能刪）" style="padding:2px 7px"><i class="bi bi-trash"></i></button>`}</td></tr>`;
  }
  if (!list.length) h += `<tr><td colspan="8" class="muted" style="padding:12px">還沒有瑪氏訂單，倉會在匯入出貨彙總表後自動列出來；也可以在下面直接加。</td></tr>`;
  h += `<tr id="wh-new"><td><input class="inp" id="wh-new-code" placeholder="倉別 TAO8" style="width:90px;text-transform:uppercase"></td>${F.map(f => `<td><input class="inp" data-f="${f[0]}" placeholder="${esc(f[1].split("（")[0])}"></td>`).join("")}<td><button class="btn btn-p btn-sm" id="wh-add"><i class="bi bi-plus-lg"></i> 加入</button></td></tr>`;
  $("#wh-table").innerHTML = h + "</tbody>";
  $("#wh-add").addEventListener("click", async () => {
    const body = { code: $("#wh-new-code").value }; $("#wh-new").querySelectorAll("input[data-f]").forEach(i => body[i.dataset.f] = i.value);
    try { const d = await api("/api/mars/warehouses", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      toast(`已加入 ${d.code}`); renderWarehouses(d.warehouses); loadSplits();
    } catch (e) { toast(e.message, "err"); $("#wh-new-code").classList.add("bad"); }
  });
  $("#wh-new").querySelectorAll("input").forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") { e.preventDefault(); $("#wh-add").click(); } }));
  $("#wh-table").querySelectorAll(".wh-del").forEach(b => b.addEventListener("click", async () => {
    if (!confirm(`要刪掉 ${b.dataset.code} 的倉庫資料？`)) return;
    try { const d = await api(`/api/mars/warehouses/${encodeURIComponent(b.dataset.code)}`, { method: "DELETE" }); toast(`${b.dataset.code} 已刪`); renderWarehouses(d.warehouses); loadSplits(); }
    catch (e) { toast(e.message, "err"); }
  }));
  $("#wh-table").querySelectorAll("tr[data-code]").forEach(tr => {
    const inputs = [...tr.querySelectorAll("input.wh")]; inputs.forEach(i => i.dataset.orig = i.value);
    const save = async () => {
      if (!inputs.some(i => i.value !== i.dataset.orig)) return;
      const body = {}; inputs.forEach(i => body[i.dataset.f] = i.value);
      try { const d = await api(`/api/mars/warehouses/${encodeURIComponent(tr.dataset.code)}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
        toast(`${tr.dataset.code} 的倉庫資料已存`); renderWarehouses(d.warehouses); loadSplits();
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
    $("#wh-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> ${esc(f.name)}：${d.rows} 個倉，新增 ${d.added}、更新 ${d.updated}、沒變 ${d.same}。</span>`;
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
  $("#ps-msg").textContent = "存檔中…";
  try { await api("/api/mars/po/settings", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }); $("#ps-msg").textContent = ""; toast("採購單設定已存"); await loadPoSettings(); loadSplits(); }
  catch (e) { $("#ps-msg").textContent = ""; toast(e.message, "err"); }
});

loadStatus().then(loadSplits).catch(e => toast(e.message, "err"));
loadPoSettings().catch(e => toast(e.message, "err"));
