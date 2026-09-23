/* 瑪氏出貨：商品總表上傳、拆單、回填 EIP 採購單號／約倉時間。後端在 mars/。 */
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
let VIEW = null, OPEN = new Set();

/* ── 商品總表 ── */
async function loadStatus() {
  const d = await api("/api/mars/status");
  const u = d.last_upload;
  $("#mp-status").innerHTML = u ? `上次上傳 ${esc(u.uploaded_at.slice(0, 16))} · ${esc(u.operator)} · ${esc(u.filename)} · <b style="color:var(--m8)">${d.codes}</b> 個料號`
    : `<span class="neg">還沒上傳</span>：拆單要靠它分品類、中標，先把 Alice 的「自動化_Mars整合商品資料」傳上來`;
  const up = d.dates.filter(x => x >= today()).slice(0, 8), recent = d.dates.filter(x => x < today()).slice(-3);
  const chips = [...recent, ...up];
  $("#d-chips").innerHTML = chips.map(x => `<span class="chip" data-d="${x}" title="只看這天到貨的">${md(x)}</span>`).join("") || `<span class="kbd">② 訂單明細裡還沒有瑪氏的訂單</span>`;
  $("#d-chips").querySelectorAll(".chip").forEach(c => c.addEventListener("click", () => { $("#d-from").value = $("#d-to").value = c.dataset.d; loadSplits(); }));
  if (!$("#d-from").value) { const first = up[0] || d.dates[d.dates.length - 1] || today(); $("#d-from").value = $("#d-to").value = first; }
}
$("#mp-file").addEventListener("change", async e => {
  const f = e.target.files[0]; e.target.value = ""; if (!f) return;
  const fd = new FormData(); fd.append("file", f); $("#mp-msg").innerHTML = `<span class="muted">上傳中…</span>`;
  try {
    const d = await api("/api/mars/products/import", { method: "POST", body: fd });
    const cats = Object.entries(d.by_category).map(([k, v]) => `${k} ${v}`).join("、");
    $("#mp-msg").innerHTML = `<span style="color:var(--ok)"><i class="bi bi-check-circle"></i> 讀到 ${d.codes} 個料號（${d.rows} 列，${esc(cats)}）${d.added || d.removed ? `，跟上一版比：新增 ${d.added}、拿掉 ${d.removed}` : ""}。</span>`
      + (d.warnings.length ? `<details class="kbd"><summary>${d.warnings.length} 列有問題（跳過或標出來）</summary>${d.warnings.map(esc).join("<br>")}</details>` : "");
    toast("商品總表已更新"); await loadStatus(); loadSplits();
  } catch (err) { $("#mp-msg").innerHTML = `<span class="neg">${esc(err.message)}</span>`; }
});

/* ── 拆單 ── */
["#d-from", "#d-to"].forEach(s => $(s).addEventListener("change", loadSplits));
async function loadSplits() {
  const f = $("#d-from").value, t = $("#d-to").value; if (!f || !t) return;
  document.querySelectorAll("#d-chips .chip").forEach(c => c.classList.toggle("on", f === t && c.dataset.d === f));
  try { VIEW = await api(`/api/mars/splits?from=${f}&to=${t}`); } catch (e) { toast(e.message, "err"); return; }
  render();
}
function render() {
  const v = VIEW, s = v.summary;
  $("#sum").innerHTML = `<span><b>${s.pos}</b> 張 PO</span><span><b>${s.files}</b> 份拆單表</span><span><b>${fmt(s.cases)}</b> 箱</span><span><b>${s.items}</b> 個品項</span><span><b>${s.filled}</b>／${s.files} 已回填 EIP 單號</span>`;
  const al = [];
  if (!v.has_products) al.push(`<div class="alert al-bad"><i class="bi bi-exclamation-octagon-fill"></i> 還沒上傳瑪氏商品總表，分不出品類和中標，拆不了。</div>`);
  if (v.unmatched.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-exclamation-octagon-fill"></i> ${v.unmatched.length} 個品項對不到瑪氏商品總表</b>，這些不會進拆單表，請更新商品總表再重拆：<br>${v.unmatched.map(u => `${esc(u.po_number)} · ${esc(u.yf_sku)}${u.quote_note && u.quote_note !== u.yf_sku ? `（報價備註 ${esc(u.quote_note)}）` : ""} · ${esc(u.product_name)} · ${esc(u.unit)} ${fmt(u.qty_ship)}`).join("<br>")}</div>`);
  const blocking = v.splits.flatMap(r => r.blocking || []);
  if (blocking.length) al.push(`<div class="alert al-bad"><b><i class="bi bi-slash-circle"></i> 這幾個品項要先處理才能產出</b>（EIP 採購表只能填整數箱）：<br>${blocking.map(esc).join("<br>")}</div>`);
  if (v.zero_rows) al.push(`<div class="kbd">出貨數量是 0 的 ${v.zero_rows} 個品項沒有拆進來。</div>`);
  $("#alerts").innerHTML = al.join("");
  $("#btn-gen").disabled = !v.splits.length || !v.has_products || blocking.length > 0;

  let h = `<thead><tr><th>狀態</th><th>拆單表檔名</th><th>到貨日</th><th>酷澎 PO</th><th>倉</th><th>品類</th>${v.split_by_unit ? "<th>單位</th>" : ""}<th>中標</th><th class="num">品項</th><th class="num">箱數</th><th style="min-width:150px">EIP 採購單號</th><th style="min-width:170px">約倉時間</th><th>下載</th></tr></thead><tbody>`;
  for (const r of v.splits) {
    const key = r.split_key, warn = r.items.some(i => i.issues && i.issues.length);
    const dis = r.id ? "" : "disabled title=\"先按「產出拆單表」才能填\"";
    h += `<tr class="sp" data-k="${esc(key)}">
      <td><span class="badge st-${r.status}">${STATUS[r.status]}</span>${r.diff ? `<div class="kbd" style="max-width:220px;color:var(--bad)">${esc(r.diff)}</div>` : ""}</td>
      <td><span class="fn fname">${esc(r.filename)}</span> <button class="btn btn-o btn-sm cp" title="複製檔名" data-t="${esc(r.filename)}" style="padding:1px 6px"><i class="bi bi-clipboard"></i></button></td>
      <td>${md(r.delivery_date)}</td><td class="fn">${esc(r.po_number)}</td><td>${esc(r.warehouse)}</td>
      <td><span class="cat cat-${esc(r.category)}">${esc(r.category || "?")}</span></td>${v.split_by_unit ? `<td>${esc(r.unit)}</td>` : ""}
      <td><span class="cat lbl-${r.label}">${r.label === "V" ? "需貼中標" : "不貼中標"}</span></td>
      <td class="num">${r.item_count}${warn ? ` <i class="bi bi-exclamation-triangle-fill" style="color:var(--warn)" title="有品項要注意，點開看"></i>` : ""}</td><td class="num"><b>${fmt(r.cases_total)}</b></td>
      <td><input class="inp eip" data-id="${r.id || ""}" value="${esc(r.eip_po)}" placeholder="PO202609…" ${dis}></td>
      <td><input class="inp slot" data-id="${r.id || ""}" value="${esc(r.slot_time)}" placeholder="例如 12:30~15:30（1台車）" ${dis}></td>
      <td class="whitespace-nowrap">${r.id ? `<a class="btn btn-o btn-sm" href="/api/mars/splits/${r.id}/file?kind=split" title="拆單表"><i class="bi bi-file-earmark-excel"></i></a> <a class="btn btn-o btn-sm" href="/api/mars/splits/${r.id}/file?kind=eip" title="EIP 上傳用採購表">EIP</a>` : `<span class="kbd">—</span>`}</td></tr>`;
    if (OPEN.has(key)) {
      h += `<tr class="sub"><td></td><td colspan="${v.split_by_unit ? 12 : 11}"><table class="t"><thead><tr><th>永豐料號</th><th>下採料號</th><th>品名</th><th class="num">出貨數量</th><th>單位</th><th class="num">箱入數</th><th class="num">箱數</th><th>瑪氏貨號</th><th>採購單箱備註</th><th>要注意</th></tr></thead><tbody>
        ${r.items.map(i => `<tr><td class="fn">${esc(i.yf_sku)}</td><td class="fn">${esc(i.purchase_code)}${i.via === "報價備註" ? ` <span class="badge st-changed">報價備註</span>` : ""}</td><td>${esc(i.product_name)}</td><td class="num">${fmt(i.qty_ship)}${i.qty_overridden ? ` <span class="badge st-generated" title="在 ② 訂單明細人工改過">改過</span>` : ""}</td><td>${esc(i.unit)}</td><td class="num">${fmt(i.box_file)}</td><td class="num"><b>${fmt(i.cases)}</b></td><td class="fn">${esc(i.mars_code)}</td><td>${esc(i.po_case_note)}</td><td style="color:var(--warn)">${(i.issues || []).map(esc).join("<br>")}</td></tr>`).join("")}
      </tbody></table></td></tr>`;
    }
  }
  if (!v.splits.length) h += `<tr><td colspan="13" class="muted" style="padding:24px;text-align:center">這段期間沒有瑪氏的訂單。訂單在「商品主檔自動化 ② 訂單明細」匯入。</td></tr>`;
  $("#sp-table").innerHTML = h + "</tbody>";
  $("#sp-table").querySelectorAll("tr.sp").forEach(tr => tr.addEventListener("click", e => {
    if (e.target.closest("input, a, button")) return;
    const k = tr.dataset.k; OPEN.has(k) ? OPEN.delete(k) : OPEN.add(k); render();
  }));
  $("#sp-table").querySelectorAll(".cp").forEach(b => b.addEventListener("click", async () => { try { await navigator.clipboard.writeText(b.dataset.t); toast("檔名已複製"); } catch (e) { toast("瀏覽器不讓複製，請手動選取", "err"); } }));
  $("#sp-table").querySelectorAll("input.eip, input.slot").forEach(inp => {
    inp.dataset.orig = inp.value;
    const save = async () => {
      if (!inp.dataset.id || inp.value === inp.dataset.orig) return;
      const body = inp.classList.contains("eip") ? { eip_po: inp.value } : { slot_time: inp.value };
      try { const d = await api(`/api/mars/splits/${inp.dataset.id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
        inp.classList.remove("bad"); inp.dataset.orig = inp.classList.contains("eip") ? d.split.eip_po : d.split.slot_time; inp.value = inp.dataset.orig;
        toast(inp.classList.contains("eip") ? (inp.value ? `已存 ${inp.value}` : "EIP 採購單號已清掉") : "約倉時間已存"); loadSplits();
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
    toast("拆單表已產出，EIP 拿到單號後填回來"); loadSplits();
  } catch (e) { toast(e.message, "err"); }
  finally { $("#btn-gen").disabled = false; }
}
$("#btn-gen").addEventListener("click", () => generate(false));

loadStatus().then(loadSplits).catch(e => toast(e.message, "err"));
