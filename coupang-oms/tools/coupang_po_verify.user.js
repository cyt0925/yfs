// ==UserScript==
// @name         酷澎 PO 批次驗收比對器
// @namespace    yfycpg.kate
// @version      7.4
// @description  兩層驗收：第一層 清單頁收貨數量 vs detail可交貨總數(confirmedQty加總)；不符才下鑽逐SKU比 receivedQty vs confirmedQty。可貼上訂單系統複製的 PO 單號指定要驗哪幾張，或掃描目前清單頁。多狀態勾選(預設已確認+已關閉)、匯出Excel、同步「實際驗入數量」與「驗收金額(訂單金額稅後)」到訂單系統。
// @match        https://supplier.tw.coupang.com/pom/purchase-order/*
// @run-at       document-idle
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

// 這份腳本裝在 Tampermonkey 裡、跑在酷澎供應商後台，跟訂單系統
// （coupang-oms）透過 /api/sync/verified-qty 對話。放在 repo 裡是為了
// 版本跟得上後端：後端多收一個欄位，腳本就在這裡一起改、一起看歷史。
// 改完要請使用者把整份貼回 Tampermonkey 才會生效，後端部署不會自動更新它。

(function () {
  'use strict';

  const API = (poId) => `https://supplier.tw.coupang.com/pom/po/detail/${poId}`;
  const PO_STATUS_TXT = { CREATED: '已建立PO', CONFIRMED: '已確認PO', CANCELED: '已取消PO', CLOSED: '已關閉' };
  const toNum = (t) => { const m = String(t ?? '').replace(/,/g, '').match(/-?\d+(\.\d+)?/); return m ? parseFloat(m[0]) : null; };

  // ── 驗收金額：酷澎後台「訂單金額(稅後)」──
  // 2026-09-04 從真實 API 回應對出來的欄位名（poSkuList 每個 SKU 物件）：
  //   purchasePriceAfterTax   訂單金額(稅後) ＝ unitPriceAfterTax × orderedQty
  //                           （畫面上 88,176 = 167.00 × 528）← 用這個
  //   purchasePriceBeforeTax / purchasePriceTax   訂單金額 稅前／稅金
  //   receivingPriceAfterTax  實際收貨的稅後金額（依 receivedQty 算），
  //                           要改成「以收到的算錢」就換成這個
  //   unitPriceAfterTax / unitPriceBeforeTax      單價 稅後／稅前
  // ⚠ 這些金額欄位 API 回的是「分」（×100 的整數）：後台畫面顯示
  //   178,854.00 的那張，API 給的是 17885400。一律除以 100 才是畫面上的
  //   元。第一版沒除，同步進訂單系統的金額全部大了 100 倍，被使用者抓到。
  // 主欄位抓不到就退回「稅後單價 × PO 數量」自己算；還是算不出來就回
  // null——訂單系統那邊 null 代表「這次沒抓到」，不會把原本的金額清掉，
  // 也不會被當成 0 元，並在面板上把該 SKU 的原始資料攤出來給人看。
  const AMOUNT_FIELDS = ['purchasePriceAfterTax'];
  const UNIT_PRICE_FIELDS = ['unitPriceAfterTax'];
  const MINOR_UNITS = 100;   // API 金額單位：分 → 元
  // 找不到金額欄位時留一份樣本，直接顯示在面板上讓使用者複製給維護的
  // 人，不用去翻 Console。
  let amountSample = null;
  function pickAmount(s) {
    for (const k of AMOUNT_FIELDS) {
      const v = toNum(s[k]);
      if (v != null) return Math.round(v) / MINOR_UNITS;
    }
    const qty = toNum(s.orderedQty);
    for (const k of UNIT_PRICE_FIELDS) {
      const p = toNum(s[k]);
      if (p != null && qty != null) return Math.round(p * qty) / MINOR_UNITS;
    }
    if (!amountSample) {
      amountSample = s;
      console.warn('[PO 批次驗收] 找不到「訂單金額(稅後)」欄位，這個 SKU 的欄位有：',
        Object.keys(s).join(', '), s);
    }
    return null;
  }

  // ── 訂單系統同步設定（存在這台電腦、這個瀏覽器裡，不會傳去別的地方）──
  // 網址跟通行碼在訂單系統網頁右上角「⚙ 設定 → 驗收同步」那頁可以看到，
  // 貼過來這裡存一次，之後每次跑批次驗收都能一鍵同步，不用再手動存
  // Excel、上傳。
  const CFG_KEY = 'kpv_sync_cfg';
  function loadCfg() {
    try { return JSON.parse(localStorage.getItem(CFG_KEY) || '{}'); }
    catch { return {}; }
  }
  function saveCfg(cfg) {
    localStorage.setItem(CFG_KEY, JSON.stringify(cfg));
  }

  // 從訂單系統複製過來的 PO 單號清單（一行一個，也接受逗號／空白
  // 分隔）。有貼東西就用貼的這份當驗收對象，不用管目前這頁清單裡
  // 掃得到什麼——這樣不管誰上傳整合表、要驗哪幾張，一律照訂單系統
  // 畫面上勾了什麼算，不用自己去猜規則。
  function parsePastedPoList(text) {
    const ids = (text || '').split(/[\s,，、]+/).map(s => s.trim()).filter(Boolean);
    const seen = new Set();
    return ids.filter(id => (seen.has(id) ? false : seen.add(id)))
      .map(id => ({ poId: id, listRecv: null }));
  }

  function scrapeList() {
    const tables = [...document.querySelectorAll('table')];
    let target = null, colRecv = -1;
    for (const tb of tables) {
      const heads = [...tb.querySelectorAll('thead th, thead td, tr:first-child th, tr:first-child td')].map(c => c.innerText.replace(/\s+/g, ''));
      const idx = heads.findIndex(h => h.includes('收貨數量'));
      if (idx !== -1) { target = tb; colRecv = idx; break; }
    }
    const rows = [];
    const anchors = (target || document).querySelectorAll('a[href*="purchase-order/detail/"]');
    anchors.forEach(a => {
      const id = (a.getAttribute('href').match(/detail\/(\d+)/) || [])[1];
      if (!id) return;
      const tr = a.closest('tr');
      const cells = tr ? [...tr.children] : [];
      const recvCell = (colRecv !== -1 && cells[colRecv]) ? cells[colRecv].innerText : '';
      rows.push({ poId: id, listRecv: toNum(recvCell), statusText: tr ? tr.innerText.replace(/\s+/g, ' ') : '' });
    });
    const seen = new Set();
    return rows.filter(r => (seen.has(r.poId) ? false : seen.add(r.poId)));
  }

  async function fetchPO(poId) {
    const r = await fetch(API(poId), { credentials: 'include', headers: { Accept: 'application/json' } });
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const j = await r.json();
    const b = j.body || {};
    return { poId, status: b.status || '', skus: Array.isArray(b.poSkuList) ? b.poSkuList : [], edd: b.edd || '' };
  }

  function sums(skus) {
    const active = skus.filter(s => s.status !== 'CANCELED');
    let conf = 0, recv = 0;
    active.forEach(s => { conf += Number(s.confirmedQty ?? 0); recv += Number(s.receivedQty ?? 0); });
    return { conf, recv, active };
  }
  function drill(active) {
    return active.filter(s => Number(s.receivedQty ?? 0) !== Number(s.confirmedQty ?? 0))
      .map(s => ({ skuId: s.skuId, name: s.skuName || '', recv: Number(s.receivedQty ?? 0), conf: Number(s.confirmedQty ?? 0), ord: Number(s.orderedQty ?? 0), st: s.status }));
  }

  const S = document.createElement('style');
  S.textContent = `
  #kpv{position:fixed;top:70px;right:16px;z-index:99999;width:410px;font:13px/1.55 -apple-system,"Microsoft JhengHei",sans-serif;
    background:#fff;border:1px solid #d1d5db;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,.18);overflow:hidden}
  #kpv .hd{padding:11px 14px;background:#1f2937;color:#fff;display:flex;justify-content:space-between;align-items:center;cursor:move}
  #kpv .hd b{font-size:14px}#kpv .hd .x{cursor:pointer;opacity:.7;font-size:16px}
  #kpv .bd{padding:14px;max-height:70vh;overflow:auto}
  #kpv .chks{display:flex;gap:14px;margin-bottom:10px;flex-wrap:wrap}
  #kpv .chks label{font-size:12px;display:flex;align-items:center;gap:4px;cursor:pointer}
  #kpv .ctl{display:flex;gap:8px;align-items:center;margin-bottom:10px;flex-wrap:wrap}
  #kpv .go{flex:2;padding:9px;border:0;border-radius:8px;background:#2563eb;color:#fff;font-weight:700;cursor:pointer}
  #kpv .go:disabled{opacity:.5;cursor:wait}
  #kpv .exp{padding:9px 12px;border:0;border-radius:8px;background:#059669;color:#fff;font-weight:600;cursor:pointer}
  #kpv .sync{padding:9px 12px;border:0;border-radius:8px;background:#7c3aed;color:#fff;font-weight:600;cursor:pointer}
  #kpv .sync:disabled{opacity:.5;cursor:wait}
  #kpv .cfg{padding:8px 10px;border:0;border-radius:8px;background:#f3f4f6;color:#374151;font-weight:600;cursor:pointer}
  #kpv .sum{display:flex;gap:8px;margin:10px 0}
  #kpv .sum div{flex:1;text-align:center;padding:8px;border-radius:8px;font-weight:700}
  #kpv .sp{background:#dcfce7;color:#166534}#kpv .sf{background:#fee2e2;color:#991b1b}
  #kpv .card{border:1px solid #eee;border-radius:8px;padding:8px 10px;margin-bottom:8px}
  #kpv .t{display:flex;justify-content:space-between;font-weight:700;align-items:center}
  #kpv .badge{padding:1px 8px;border-radius:99px;font-size:11px}
  #kpv .bp{background:#dcfce7;color:#166534}#kpv .bf{background:#fee2e2;color:#991b1b}
  #kpv .l1{font-size:12px;color:#374151;margin-top:4px}
  #kpv .fl{font-size:12px;margin-top:5px}
  #kpv .fl div{display:flex;justify-content:space-between;border-top:1px dashed #f0f0f0;padding:2px 0;gap:8px}
  #kpv .fl .r{color:#991b1b;white-space:nowrap}
  #kpv .nm{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
  #kpv .prog{font-size:12px;color:#6b7280;margin:6px 0;min-height:16px}
  #kpv .note{font-size:11px;color:#9ca3af;margin-top:2px}
  #kpv .warn{font-size:11px;color:#b45309;background:#fef3c7;border-radius:6px;padding:4px 8px;margin-top:6px}
  #kpv .warn textarea{width:100%;box-sizing:border-box;margin-top:4px;font:10px/1.3 monospace;height:110px;border:1px solid #fcd34d;border-radius:4px;background:#fff;color:#374151}
  #kpv .warn button{margin-top:4px;padding:3px 8px;border:0;border-radius:4px;background:#b45309;color:#fff;font-size:11px;cursor:pointer}
  #kpv .cfgbox{display:none;border:1px solid #e5e7eb;border-radius:8px;padding:8px;margin-bottom:10px;background:#fafafa}
  #kpv .cfgbox.open{display:block}
  #kpv .cfgbox input{width:100%;box-sizing:border-box;padding:6px 8px;margin-bottom:6px;border:1px solid #d1d5db;border-radius:6px;font-size:12px}
  #kpv .cfgbox .save{padding:6px 10px;border:0;border-radius:6px;background:#1f2937;color:#fff;font-size:12px;cursor:pointer}
  #kpv .polist{width:100%;box-sizing:border-box;padding:6px 8px;margin-bottom:8px;border:1px solid #d1d5db;border-radius:6px;font-size:12px;font-family:monospace;resize:vertical;min-height:52px}
  #kpv .polist:focus{border-color:#2563eb;outline:none}
  #kpv .pocount{font-size:11px;color:#7c3aed;margin:-4px 0 8px 2px;min-height:14px}
  `;
  document.head.appendChild(S);

  const box = document.createElement('div');
  box.id = 'kpv';
  box.innerHTML = `
    <div class="hd"><b>PO 批次驗收 <span class="note" style="color:#9ca3af">v7.4</span></b><span class="x">×</span></div>
    <div class="bd">
      <textarea class="polist" id="kpv-po-list" placeholder="從訂單系統勾選 PO → 按「複製 PO 單號給驗收工具」→ 貼在這裡（一行一個）。留空則改成掃描目前這頁清單。"></textarea>
      <div class="pocount" id="kpv-pocount"></div>
      <div class="chks">
        <label><input type="checkbox" value="CONFIRMED" checked> 已確認PO</label>
        <label><input type="checkbox" value="CLOSED" checked> 已關閉</label>
        <label><input type="checkbox" value="CREATED"> 已建立PO</label>
      </div>
      <div class="ctl">
        <button class="go" id="kpv-go">▶ 開始批次驗收</button>
        <button class="exp" id="kpv-exp">Excel</button>
      </div>
      <div class="ctl">
        <button class="sync" id="kpv-sync" disabled>↑ 同步到訂單系統</button>
        <button class="cfg" id="kpv-cfg">⚙ 同步設定</button>
      </div>
      <div class="cfgbox" id="kpv-cfgbox">
        <input id="kpv-cfg-url" placeholder="同步網址（訂單系統 設定→驗收同步 裡的那串）">
        <input id="kpv-cfg-token" placeholder="通行碼" type="password">
        <button class="save" id="kpv-cfg-save">儲存設定</button>
      </div>
      <div class="note">第一層：清單「收貨數量」 vs detail 可交貨總數。不符才下鑽逐 SKU 比對。<br>※有貼 PO 單號就只驗貼的那幾張（不受下面狀態勾選限制）；沒貼才會掃描目前清單頁、且只驗有勾選狀態的。<br>※同步時會一併送每個品項的「訂單金額(稅後)」，訂單系統那邊加總成整張單的驗收金額。</div>
      <div class="prog" id="kpv-prog"></div>
      <div id="kpv-amount-warn"></div>
      <div id="kpv-out"></div>
    </div>`;
  document.body.appendChild(box);
  box.querySelector('.x').onclick = () => box.remove();

  // 設定面板：填過的網址、通行碼下次開頁面會自動帶出來
  const cfg0 = loadCfg();
  box.querySelector('#kpv-cfg-url').value = cfg0.url || '';
  box.querySelector('#kpv-cfg-token').value = cfg0.token || '';
  box.querySelector('#kpv-cfg').onclick = () => {
    box.querySelector('#kpv-cfgbox').classList.toggle('open');
  };
  box.querySelector('#kpv-cfg-save').onclick = () => {
    const url = box.querySelector('#kpv-cfg-url').value.trim();
    const token = box.querySelector('#kpv-cfg-token').value.trim();
    saveCfg({ url, token });
    box.querySelector('#kpv-prog').textContent = '同步設定已儲存。';
  };

  let results = [];

  // 貼上的內容即時顯示解析出幾張，貼完馬上看得出來有沒有貼對，
  // 不用按下去才發現數字不對。
  box.querySelector('#kpv-po-list').addEventListener('input', e => {
    const n = parsePastedPoList(e.target.value).length;
    box.querySelector('#kpv-pocount').textContent = n ? `解析到 ${n} 張 PO 單號` : '';
  });

  box.querySelector('#kpv-go').onclick = async () => {
    const btn = box.querySelector('#kpv-go'), prog = box.querySelector('#kpv-prog'), out = box.querySelector('#kpv-out');
    const pasted = parsePastedPoList(box.querySelector('#kpv-po-list').value);
    const usePasted = pasted.length > 0;

    const wantStatuses = [...box.querySelectorAll('.chks input:checked')].map(c => c.value);
    if (!usePasted && !wantStatuses.length) { prog.textContent = '請至少勾一個 PO 狀態'; return; }

    let list = usePasted ? pasted : scrapeList();
    if (!list.length) {
      prog.textContent = usePasted ? '貼的內容解析不出任何 PO 單號'
        : '清單頁抓不到 PO（往下捲載入、或確認在清單頁）';
      return;
    }

    btn.disabled = true; results = []; out.innerHTML = '';
    box.querySelector('#kpv-amount-warn').innerHTML = '';
    amountSample = null;
    box.querySelector('#kpv-sync').disabled = true;
    let done = 0, pass = 0, fail = 0, skip = 0, amountMissing = 0, skuTotal = 0;

    for (const r of list) {
      prog.textContent = `處理中… ${done + 1}/${list.length}（${r.poId}）`;
      try {
        const po = await fetchPO(r.poId);
        // 貼上模式：這幾張是使用者在訂單系統裡明確勾出來要驗的，
        // 不套用狀態勾選過濾（酷澎的狀態代碼跟訂單系統的狀態欄位本來
        // 就是兩套不相干的詞彙，硬套只會莫名其妙漏驗）。
        if (!usePasted && !wantStatuses.includes(po.status)) { skip++; done++; continue; }
        const { conf, recv, active } = sums(po.skus);
        const listRecv = (r.listRecv != null) ? r.listRecv : recv;
        const layer1Match = listRecv === conf;
        let fails = [];
        if (!layer1Match) fails = drill(active);
        // skuData 保留「每一個品項」的實際驗入數量（收貨數量 receivedQty，
        // 不是供應商可交貨數量 confirmedQty——後者幾乎都等於出貨數量，
        // 存過去會看起來永遠一樣）跟訂單金額(稅後)。不是只有不符的都要送，
        // 同步到訂單系統時要逐項送過去，不能只送有問題的那幾個。
        const skuData = active.map(s => {
          const amount = pickAmount(s);
          skuTotal++;
          if (amount == null) amountMissing++;
          return { skuId: String(s.skuId), verifiedQty: Number(s.receivedQty ?? 0), verifiedAmount: amount };
        });
        results.push({ poId: r.poId, status: po.status, edd: po.edd, listRecv, confTotal: conf, detailRecv: recv, pass: layer1Match, fails, skuData });
        layer1Match ? pass++ : fail++;
      } catch (e) {
        results.push({ poId: r.poId, status: '?', reason: 'API失敗:' + e.message, pass: false, fails: [], skuData: [] });
        fail++;
      }
      done++;
      await new Promise(s => setTimeout(s, 150));
    }

    prog.textContent = `完成：驗 ${pass + fail} 張` + (skip ? `（略過非勾選狀態 ${skip} 張）` : '');
    if (amountMissing) {
      // 金額欄位名字猜不到時明講，不要讓使用者以為有同步到金額。
      // 把第一個抓不到的 SKU 原始資料直接攤在面板上，複製貼給維護的人就
      // 能對出正確的欄位名，不用教人開 Console。
      const sample = amountSample ? JSON.stringify(amountSample, null, 1) : '';
      box.querySelector('#kpv-amount-warn').innerHTML =
        `<div class="warn">⚠ ${amountMissing}/${skuTotal} 個品項抓不到「訂單金額(稅後)」，同步時這些品項不會帶金額。<br>`
        + `下面是其中一個品項的原始資料，請整段複製貼給維護的人：`
        + `<textarea id="kpv-sample" readonly>${sample.replace(/</g, '&lt;')}</textarea>`
        + `<button id="kpv-sample-copy">複製這段</button></div>`;
      box.querySelector('#kpv-sample-copy').onclick = async () => {
        try { await navigator.clipboard.writeText(sample); box.querySelector('#kpv-sample-copy').textContent = '已複製'; }
        catch { box.querySelector('#kpv-sample').select(); document.execCommand('copy'); }
      };
    }
    const stTxt = s => PO_STATUS_TXT[s] || s || '?';
    out.innerHTML = `<div class="sum"><div class="sp">相符 ${pass}</div><div class="sf">短少 ${fail}</div></div>` +
      results.map(r => {
        if (r.reason) return `<div class="card"><div class="t"><span>${r.poId}</span><span class="badge bf">ERROR</span></div><div class="fl"><div class="r">${r.reason}</div></div></div>`;
        const amountSum = (r.skuData || []).reduce((a, s) => a + (s.verifiedAmount ?? 0), 0);
        const hasAmount = (r.skuData || []).some(s => s.verifiedAmount != null);
        const l1 = `<div class="l1">收貨 ${r.listRecv} ／ 可交貨總數 ${r.confTotal}${r.edd ? '　EDD ' + r.edd : ''}`
          + (hasAmount ? `　金額(稅後) ${amountSum.toLocaleString()}` : '') + `</div>`;
        const failHtml = (r.fails || []).map(f =>
          `<div><span class="nm">${f.skuId} ${f.name}</span><span class="r">收 ${f.recv} / 可交 ${f.conf}（少 ${f.conf - f.recv}）</span></div>`).join('');
        return `<div class="card">
          <div class="t"><span>${r.poId} <span class="note">${stTxt(r.status)}</span></span>
            <span class="badge ${r.pass ? 'bp' : 'bf'}">${r.pass ? '相符' : '短少'}</span></div>
          ${l1}${failHtml ? `<div class="fl">${failHtml}</div>` : ''}
        </div>`;
      }).join('');
    btn.disabled = false;
    box.querySelector('#kpv-sync').disabled = !results.some(r => !r.reason);
  };

  box.querySelector('#kpv-exp').onclick = () => {
    if (!results.length) { alert('先跑批次驗收'); return; }
    if (typeof XLSX === 'undefined') { alert('Excel 套件尚未載入，請稍等幾秒重按'); return; }
    const stTxt = s => PO_STATUS_TXT[s] || s || '?';
    const aoa = [['PO單號', '狀態', 'EDD', '結果', '清單收貨', '可交貨總數', '金額(稅後)', '短少SKU', 'SKU名稱', 'SKU收貨', 'SKU可交貨', '少']];
    for (const r of results) {
      if (r.reason) { aoa.push([r.poId, stTxt(r.status), '', 'ERROR', '', '', '', r.reason, '', '', '', '']); continue; }
      const hasAmount = (r.skuData || []).some(s => s.verifiedAmount != null);
      const amountSum = hasAmount ? (r.skuData || []).reduce((a, s) => a + (s.verifiedAmount ?? 0), 0) : '';
      if (r.fails && r.fails.length) {
        r.fails.forEach(f => aoa.push([r.poId, stTxt(r.status), r.edd || '', '短少', r.listRecv, r.confTotal, amountSum, f.skuId, f.name, f.recv, f.conf, f.conf - f.recv]));
      } else {
        aoa.push([r.poId, stTxt(r.status), r.edd || '', '相符', r.listRecv, r.confTotal, amountSum, '', '', '', '', '']);
      }
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 18 }, { wch: 10 }, { wch: 11 }, { wch: 7 }, { wch: 9 }, { wch: 11 }, { wch: 12 }, { wch: 18 }, { wch: 32 }, { wch: 9 }, { wch: 10 }, { wch: 6 }];
    // 把長 ID 欄（A=PO單號, H=短少SKU）強制成文字，避免 Excel 轉科學記號
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (const C of [0, 7]) { // A欄、H欄
        const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
        if (cell && cell.v !== '' && cell.v != null) { cell.t = 's'; cell.v = String(cell.v); cell.z = '@'; }
      }
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'PO驗收');
    XLSX.writeFile(wb, `PO驗收_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  // 把這次跑出來的「每個品項收貨數量＋訂單金額(稅後)」送進訂單系統，
  // 系統那邊會存成「實際驗入數量」「驗收金額」，並且自動判斷短驗、補進
  // 驗收註記——這一步取代了「匯出 Excel 再手動上傳」，按一次全部同步完。
  box.querySelector('#kpv-sync').onclick = async () => {
    const cfg = loadCfg();
    const prog = box.querySelector('#kpv-prog');
    if (!cfg.url || !cfg.token) {
      alert('請先按「⚙ 同步設定」，貼上訂單系統給的同步網址跟通行碼。');
      return;
    }
    if (!results.length) { alert('先跑批次驗收'); return; }

    const items = [];
    for (const r of results) {
      if (r.reason) continue;
      for (const s of (r.skuData || [])) {
        const item = { po_number: r.poId, sku_id: s.skuId, verified_qty: s.verifiedQty };
        // 抓不到金額就不送這個鍵——訂單系統那邊「沒送」代表保留原值，
        // 送 null／0 反而會把之前抓到的金額洗掉或當成 0 元。
        if (s.verifiedAmount != null) item.verified_amount = s.verifiedAmount;
        items.push(item);
      }
    }
    if (!items.length) { alert('沒有可以同步的品項。'); return; }

    const btn = box.querySelector('#kpv-sync');
    btn.disabled = true;
    prog.textContent = `同步中…（共 ${items.length} 個品項）`;
    try {
      const res = await fetch(cfg.url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'X-Sync-Token': cfg.token },
        body: JSON.stringify({ operator: '酷澎驗收工具', items }),
      });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) {
        prog.textContent = `同步失敗：${data.error || res.status}`;
      } else {
        // 後端有回 message 就直接用它（含金額筆數、自動判定結果），舊版後端
        // 沒有的話退回自己拼。
        prog.textContent = data.message || (`同步完成：比對到 ${data.matched} 個品項` +
          (data.not_found && data.not_found.length ? `，${data.not_found.length} 個系統裡還沒有` : ''));
      }
    } catch (e) {
      prog.textContent = '同步失敗：' + e.message;
    } finally {
      btn.disabled = false;
    }
  };

  (function () {
    const hd = box.querySelector('.hd'); let d = false, sx, sy, ox, oy;
    hd.onmousedown = e => { if (e.target.classList.contains('x')) return; d = true; sx = e.clientX; sy = e.clientY; const r = box.getBoundingClientRect(); ox = r.left; oy = r.top; };
    document.onmousemove = e => { if (!d) return; box.style.left = (ox + e.clientX - sx) + 'px'; box.style.top = (oy + e.clientY - sy) + 'px'; box.style.right = 'auto'; };
    document.onmouseup = () => d = false;
  })();
})();
