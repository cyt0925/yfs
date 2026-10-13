/* 今日運彩總覽：讀 data/today.json，純前端渲染 */
(function () {
  "use strict";

  const $ = (s, el = document) => el.querySelector(s);
  const params = new URLSearchParams(location.search);
  const DATA_URL = params.get("data") === "sample" ? "data/sample.json" : "data/today.json";
  const SPORT_ICON = { all: "i-grid", basketball: "i-basketball", soccer: "i-soccer", hockey: "i-hockey", tennis: "i-tennis", football: "i-football" };

  const state = {
    data: null,
    sport: localStorage.getItem("sport") || "all",
    view: "games",
    open: new Set(),
    standingsLeague: localStorage.getItem("standingsLeague") || null,
    showAllInj: false,
    showAllHot: false,
  };

  // ---------- 小工具 ----------
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const pct = (p) => (p == null ? "" : Math.round(p * 100) + "%");
  const tname = (t) => t.name_zh || t.short || t.name || "";
  const initials = (t) => (t.abbr || (t.name || "?").slice(0, 3)).toUpperCase();

  function todayTW() {
    return new Date().toLocaleDateString("en-CA", { timeZone: "Asia/Taipei" });
  }
  function dayLabel(day) {
    const t = todayTW();
    const toUTC = (s) => { const [y, m, dd] = s.split("-").map(Number); return Date.UTC(y, m - 1, dd); };
    const diff = Math.round((toUTC(day) - toUTC(t)) / 86400000);
    const wk = ["日", "一", "二", "三", "四", "五", "六"][new Date(toUTC(day)).getUTCDay()];
    const md = day.slice(5).replace("-", "/");
    const name = diff === 0 ? "今天" : diff === 1 ? "明天" : diff === -1 ? "昨天" : diff === 2 ? "後天" : "";
    return { name: name || md, sub: `${md}（${wk}）` };
  }
  function ago(ts) {
    const m = Math.max(0, Math.round((Date.now() / 1000 - ts) / 60));
    if (m < 1) return "剛剛更新";
    if (m < 60) return `${m} 分鐘前更新`;
    const h = Math.floor(m / 60);
    if (h < 24) return `${h} 小時前更新`;
    return `${Math.floor(h / 24)} 天前更新`;
  }
  function logoImg(t, cls = "logo") {
    if (t.logo) return `<img class="${cls}" src="${esc(t.logo)}" alt="" loading="lazy" data-fb="${esc(initials(t))}">`;
    return `<span class="${cls} fallback">${esc(initials(t))}</span>`;
  }
  function avatarImg(p) {
    const fb = esc((p.player || "?").split(" ").map((w) => w[0]).join("").slice(0, 2).toUpperCase());
    if (p.headshot) return `<img class="avatar" src="${esc(p.headshot)}" alt="" loading="lazy" data-fb="${fb}">`;
    return `<span class="avatar fallback">${fb}</span>`;
  }
  const icon = (id) => `<svg><use href="#${id}"/></svg>`;

  // 圖片載入失敗就換成縮寫
  document.addEventListener("error", (e) => {
    const el = e.target;
    if (el.tagName !== "IMG") return;
    if (!el.dataset.fb) { el.remove(); return; }
    const span = document.createElement("span");
    span.className = el.className + " fallback";
    span.textContent = el.dataset.fb;
    el.replaceWith(span);
  }, true);

  // ---------- 資料 ----------
  async function load() {
    try {
      const r = await fetch(`${DATA_URL}?t=${Date.now()}`, { cache: "no-store" });
      if (!r.ok) throw new Error("HTTP " + r.status);
      state.data = await r.json();
    } catch (err) {
      $("#games").innerHTML = `<div class="empty">資料載入失敗（${esc(err.message)}）。排程可能還沒跑過，稍後再試。</div>`;
      return;
    }
    render();
  }

  function gamesForSport() {
    const g = state.data.games || [];
    return state.sport === "all" ? g : g.filter((x) => x.sport === state.sport);
  }

  // ---------- 渲染 ----------
  function render() {
    const d = state.data;
    document.body.dataset.sport = state.sport;
    const t = todayTW();
    const dl = dayLabel(t);
    $("#hdr-date").textContent = `${dl.sub} · ${ago(d.generated_ts)}`;
    $("#hdr-updated").textContent = "台灣時間";
    $("#foot-src").textContent = `資料來源：${d.source || "ESPN"}　｜　更新時間 ${String(d.generated_at || "").replace("T", " ").slice(0, 16)}`;
    renderChips();
    renderPicks();
    renderGames();
    renderInjuries();
    renderHot();
    renderStandings();
  }

  function renderChips() {
    const all = state.data.games || [];
    const counts = {};
    all.forEach((g) => { counts[g.sport] = (counts[g.sport] || 0) + 1; });
    $("#chips").innerHTML = (state.data.sports || []).map((s) => {
      const n = s.key === "all" ? all.length : counts[s.key] || 0;
      return `<button class="chip ${state.sport === s.key ? "on" : ""}" data-sport="${s.key}">${icon(SPORT_ICON[s.key] || "i-grid")}${esc(s.label)}<span class="n">${n}</span></button>`;
    }).join("");
  }

  function topPicks() {
    const order = { 高: 0, 中: 1, 低: 2 };
    return gamesForSport()
      .filter((g) => g.suggestion && g.suggestion.pick && g.status.state !== "post")
      .sort((a, b) => {
        const c = order[a.suggestion.confidence] - order[b.suggestion.confidence];
        if (c) return c;
        const pa = Math.max(a.odds?.prob_home || 0, a.odds?.prob_away || 0);
        const pb = Math.max(b.odds?.prob_home || 0, b.odds?.prob_away || 0);
        return pb - pa;
      });
  }

  function renderPicks() {
    const picks = topPicks();
    const el = $("#picks");
    if (!picks.length) {
      el.innerHTML = `<div class="empty">目前沒有信心足夠的建議。盤口或資料不足時，我們寧可寫「觀望」也不裝懂。</div>`;
      return;
    }
    const [first, ...rest] = picks;
    const s = first.suggestion;
    const pickT = first[s.pick];
    const prob = s.pick === "home" ? first.odds?.prob_home : first.odds?.prob_away;
    const dl = dayLabel(first.day_tw);
    const hero = `
      <div class="hero">
        <span class="kicker">${icon("i-star")}焦點賽事 · ${esc(first.league_zh)} · ${esc(dl.name)} ${esc(first.time_tw)}</span>
        <div class="hteams">
          <div class="hteam">${logoImg(first.away)}<b>${esc(tname(first.away))}</b><small>${esc(first.away.record || "客隊")}</small></div>
          <div class="hmid">VS<small>${esc(first.status.detail)}</small></div>
          <div class="hteam">${logoImg(first.home)}<b>${esc(tname(first.home))}</b><small>${esc(first.home.record || "主隊")}</small></div>
        </div>
        <div class="pickline">傾向 <b>${esc(tname(pickT))}</b>
          <span class="badge lime big">${esc(s.play)}</span>
          <span class="badge vio big" style="background:rgba(255,255,255,.18);color:#fff">信心 ${esc(s.confidence)}</span>
          ${prob ? `<span class="badge big" style="background:rgba(255,255,255,.18);color:#fff">勝率 ${pct(prob)}</span>` : ""}
          ${s.totals && (s.totals.lean === "大" || s.totals.lean === "小") ? `<span class="badge big" style="background:rgba(255,255,255,.18);color:#fff">大小分看${esc(s.totals.lean)}</span>` : ""}
        </div>
        <div class="why">${esc((s.reasons || []).slice(0, 2).join("；") || s.summary)}</div>
        <button class="go" data-goto="${esc(first.id)}">${icon("i-list")}看完整分析</button>
      </div>`;
    const grid = rest.slice(0, 6).map((g) => {
      const sg = g.suggestion;
      const p = sg.pick === "home" ? g.odds?.prob_home : g.odds?.prob_away;
      const dl2 = dayLabel(g.day_tw);
      return `<button class="pick" data-goto="${esc(g.id)}">
        <span class="logos">${logoImg(g.away, "")}<span class="x">vs</span>${logoImg(g.home, "")}</span>
        <span class="pt"><b>傾向 ${esc(tname(g[sg.pick]))}　<span class="badge lime">${esc(sg.play)}</span></b>
          <small>${esc(g.league_zh)} · ${esc(dl2.name)} ${esc(g.time_tw)} · ${esc(tname(g.away))} 對 ${esc(tname(g.home))}</small></span>
        <span class="pr"><b>${p ? pct(p) : ""}</b><small class="badge ${sg.confidence === "高" ? "ok" : "vio"}">信心${esc(sg.confidence)}</small></span>
      </button>`;
    }).join("");
    el.innerHTML = hero + (grid ? `<div class="pick-grid">${grid}</div>` : "");
  }

  function formDots(t, n = 5) {
    const f = (t.form || []).slice(-n);
    if (!f.length && t.form_str) {
      return `<div class="dots">${t.form_str.slice(-n).split("").map((r) => `<span class="dot ${r}" title="${r}">${r === "W" ? "勝" : r === "L" ? "敗" : "和"}</span>`).join("")}</div>`;
    }
    if (!f.length) return `<div class="dots"><span class="badge muted">近況暫無</span></div>`;
    return `<div class="dots">${f.map((r) => `<span class="dot ${r.result}" title="${esc(r.date)} ${r.home ? "主" : "客"} vs ${esc(r.opp_zh || r.opp)} ${r.pf}-${r.pa}">${r.result === "W" ? "勝" : r.result === "L" ? "敗" : "和"}</span>`).join("")}</div>`;
  }

  function sideBadges(t) {
    const b = [];
    if (t.key_out && t.key_out.length) b.push(`<span class="badge out">${icon("i-med")}主力缺陣 ${t.key_out.length}</span>`);
    else if (t.out_count) b.push(`<span class="badge out">${icon("i-med")}缺陣 ${t.out_count}</span>`);
    if (t.maybe_count) b.push(`<span class="badge maybe">${icon("i-warn")}存疑 ${t.maybe_count}</span>`);
    (t.hot || []).forEach((p) => {
      if (p.label === "火燙") b.push(`<span class="badge hot">${icon("i-fire")}${esc(p.player.split(" ").pop())} 火燙</span>`);
      if (p.label === "冰冷") b.push(`<span class="badge cold">${icon("i-snow")}${esc(p.player.split(" ").pop())} 冰冷</span>`);
    });
    if (t.streak) b.push(`<span class="badge ${t.streak.startsWith("連勝") ? "ok" : "muted"}">${esc(t.streak)}</span>`);
    return b.join("");
  }

  function teamBlock(t, g) {
    const en = t.name_zh && t.name && t.name !== t.name_zh ? `<span class="ten">${esc(t.name)}</span>` : "";
    const rec = g.tennis ? (t.rank ? `<span class="rec">排名 <b>${esc(t.rank)}</b></span>` : "") : (t.record ? `<span class="rec">戰績 <b>${esc(t.record)}</b>${t.last10 ? `　近10 <b>${esc(t.last10)}</b>` : ""}</span>` : "");
    return `<div class="team">
      ${g.tennis ? (t.headshot ? `<img class="logo" src="${esc(t.headshot)}" alt="" loading="lazy" data-fb="${esc(initials(t))}">` : logoImg(t)) : logoImg(t)}
      <span class="tname">${esc(tname(t))}</span>${en}${rec}
      ${g.tennis ? "" : formDots(t)}
      <div class="side-badges">${sideBadges(t)}</div>
    </div>`;
  }

  function oddsRow(g) {
    const o = g.odds;
    if (!o || (o.dec_home == null && o.dec_away == null)) return `<div class="oline">目前沒有盤口資料</div>`;
    const cell = (label, dec, prob, fav) => `<div class="obtn ${fav ? "fav" : ""}">
      <div class="ol">${label}</div><div class="ov">${dec != null ? dec.toFixed(2) : "–"}</div>
      <div class="op">${prob != null ? "勝率 " + pct(prob) : ""}</div>
      <div class="bar"><span style="width:${prob ? Math.round(prob * 100) : 0}%"></span></div></div>`;
    const favHome = o.prob_home != null && o.prob_away != null && o.prob_home >= o.prob_away;
    const hasDraw = o.dec_draw != null;
    const ah = g.tennis ? esc(tname(g.away)) : "客 " + esc(tname(g.away));
    const hh = g.tennis ? esc(tname(g.home)) : "主 " + esc(tname(g.home));
    let html = `<div class="odds ${hasDraw ? "" : "two"}">`;
    html += cell(ah, o.dec_away, o.prob_away, !favHome && o.prob_away != null);
    if (hasDraw) html += cell("和局", o.dec_draw, o.prob_draw, false);
    html += cell(hh, o.dec_home, o.prob_home, favHome);
    html += `</div>`;
    const line = [];
    if (o.spread != null) line.push(`讓分 ${esc(o.details || (o.spread > 0 ? "+" : "") + o.spread)}`);
    if (o.total != null) line.push(`大小分 ${o.total}`);
    if (o.provider) line.push(esc(o.provider));
    if (line.length) html += `<div class="oline">${line.join("　·　")}</div>`;
    return html;
  }

  function sugRow(g) {
    const s = g.suggestion;
    if (!s) return "";
    if (!s.pick) return `<div class="sug watch"><div class="st">${icon("i-warn")}<b>觀望</b><span>${esc(s.summary)}</span></div></div>`;
    const lean = s.totals && (s.totals.lean === "大" || s.totals.lean === "小") ? `<span class="badge vio">大小分看${esc(s.totals.lean)}</span>` : "";
    return `<div class="sug"><div class="st">${icon("i-star")}傾向 <b>${esc(s.pick_label)}</b><span class="badge lime">${esc(s.play)}</span>
      <span class="badge ${s.confidence === "高" ? "ok" : "vio"} conf">信心 ${esc(s.confidence)}</span>${lean}</div></div>`;
  }

  const sevClass = (s) => (s === "out" ? "out" : s === "maybe" ? "maybe" : s === "ir" ? "ir" : "muted");

  function playerRow(p, extra = "") {
    const comment = p.comment ? `<div class="pc">${esc(String(p.comment).slice(0, 180))}${p.comment.length > 180 ? "…" : ""}</div>` : "";
    return `<div class="prow">${avatarImg(p)}<div class="grow"><div class="pn">${esc(p.player)}</div><div class="ps">${esc(p.position_zh || p.position || "")}${p.detail ? " · " + esc(p.detail) : ""}${p.return_date ? " · 預計回歸 " + esc(String(p.return_date).slice(0, 10)) : ""}</div>${comment}</div>${extra}</div>`;
  }

  function injuryCol(t) {
    const list = (t.injuries || []).slice(0, 10);
    if (!list.length) return `<div class="plist"><span class="badge ok">${icon("i-check")}沒有傷兵回報</span></div>`;
    return `<div class="plist">${list.map((x) => playerRow(x, `<span class="badge ${sevClass(x.severity)}">${esc(x.status_zh)}</span>`)).join("")}</div>`;
  }

  function hotRow(p) {
    const max = Math.max(...p.games.map((x) => x.v), 1);
    const mini = `<div class="mini" title="近 5 場">${p.games.map((x) => `<i style="height:${Math.max(8, Math.round((x.v / max) * 100))}%" title="${esc(x.date)} vs ${esc(x.opp || "")}：${x.v}"></i>`).join("")}</div>`;
    const cls = p.label === "火燙" ? "hot" : p.label === "冰冷" ? "cold" : "muted";
    const ic = p.label === "火燙" ? icon("i-fire") : p.label === "冰冷" ? icon("i-snow") : "";
    return `<div class="prow">${avatarImg(p)}<div class="grow"><div class="pn">${esc(p.player)} <span class="badge ${cls}">${ic}${esc(p.label)}</span></div>
      <div class="ps">${esc(p.position || "")} · 近 5 場 ${p.last5} ${esc(p.stat)}，季平均 ${p.season}（${p.delta_pct > 0 ? "+" : ""}${p.delta_pct}%）</div></div>
      <div class="hotnum"><b>${p.last5}</b><br><small>季均 ${p.season}</small></div>${mini}</div>`;
  }

  function hotCol(t) {
    const list = t.hot || [];
    if (!list.length) return `<div class="plist"><span class="badge muted">球員近況暫無資料</span></div>`;
    return `<div class="plist">${list.map(hotRow).join("")}</div>`;
  }

  function formCol(t) {
    const f = (t.form || []).slice(-10).reverse();
    if (!f.length) return `<div class="rows"><span class="badge muted">${t.form_str ? "近況：" + esc(t.form_str) : "本季尚無完成的比賽"}</span></div>`;
    return `<div class="rows">${f.map((r) => `<div class="frow"><span class="dot ${r.result}">${r.result === "W" ? "勝" : r.result === "L" ? "敗" : "和"}</span><span class="d">${esc(r.date)}</span><span>${esc(r.opp_zh || r.opp || "")}<span class="ha">${r.home ? "主" : "客"}${r.comp ? " · " + esc(r.comp) : ""}</span></span><span class="s">${r.pf}-${r.pa}</span></div>`).join("")}</div>
      ${t.avg_for5 != null ? `<div class="meta"><span>近 5 場平均得 ${t.avg_for5}</span><span>失 ${t.avg_against5}</span><span>總分 ${t.avg_total5}</span>${t.form_note ? `<span>${esc(t.form_note)}</span>` : ""}</div>` : ""}`;
  }

  function newsList(g) {
    if (!g.news || !g.news.length) return `<span class="badge muted">24 小時內沒有相關新聞</span>`;
    return `<div class="news">${g.news.map((n) => `<a href="${esc(n.url || "#")}" target="_blank" rel="noopener">${n.image ? `<img src="${esc(n.image)}" alt="" loading="lazy">` : ""}<div><div class="nt">${esc(n.title)}</div><div class="nd">${esc(n.team === "home" ? tname(g.home) : tname(g.away))} · ${esc(String(n.published || "").slice(0, 10))}</div></div></a>`).join("")}</div>`;
  }

  function details(g) {
    const s = g.suggestion || {};
    const reasons = [...(s.reasons || []).map((r) => `<li class="plus">${icon("i-check")}<span>${esc(r)}</span></li>`), ...(s.cautions || []).map((r) => `<li class="minus">${icon("i-warn")}<span>${esc(r)}</span></li>`)];
    const colHead = (t) => `<div class="col-title">${t.logo ? `<img src="${esc(t.logo)}" alt="">` : ""}${esc(tname(t))}</div>`;
    const twoCol = (fn) => `<div class="two-col"><div>${colHead(g.away)}${fn(g.away)}</div><div>${colHead(g.home)}${fn(g.home)}</div></div>`;
    const meta = [g.venue ? `<span>${icon("i-pin")}${esc(g.venue)}</span>` : "", g.broadcast ? `<span>${icon("i-tv")}${esc(g.broadcast)}</span>` : "", g.tournament ? `<span>${icon("i-trophy")}${esc(g.tournament)}${g.round ? " · " + esc(g.round) : ""}</span>` : ""].filter(Boolean).join("");
    return `<div class="details">
      <div class="dt">${icon("i-star")}建議與理由</div>
      <ul class="reasons">${reasons.length ? reasons.join("") : `<li>${icon("i-warn")}<span>${esc(s.summary || "資料不足")}</span></li>`}</ul>
      ${s.totals ? `<div class="meta"><span>大小分盤口 ${s.totals.line}</span><span>兩隊近 5 場平均總分 ${s.totals.expected}</span><span>判斷：${esc(s.totals.lean)}</span></div>` : ""}
      ${g.tennis ? "" : `<div class="dt">${icon("i-med")}傷兵與出賽狀態</div>${twoCol(injuryCol)}`}
      ${g.tennis ? "" : `<div class="dt">${icon("i-fire")}球員手感</div>${twoCol(hotCol)}`}
      ${g.tennis ? "" : `<div class="dt">${icon("i-clock")}近 10 場</div>${twoCol(formCol)}`}
      ${g.tennis ? "" : `<div class="dt">${icon("i-news")}相關新聞</div>${newsList(g)}`}
      ${g.headline ? `<div class="meta"><span>${esc(g.headline)}</span></div>` : ""}
      ${meta ? `<div class="meta">${meta}</div>` : ""}
    </div>`;
  }

  function gameCard(g) {
    const live = g.status.state === "in";
    const post = g.status.state === "post";
    const open = state.open.has(g.id);
    const s = g.suggestion;
    const edge = s && s.pick && s.confidence === "高";
    const hasScore = (live || post) && (g.home.score != null || g.away.score != null);
    const mid = hasScore
      ? `<div class="score ${g.tennis ? "sets" : ""}">${g.tennis ? esc(g.home.score || "") + "<br>" + esc(g.away.score || "") : esc(g.away.score ?? "") + " : " + esc(g.home.score ?? "")}</div><div class="sub">${esc(g.status.detail)}</div>`
      : `<div class="time">${esc(g.time_tw)}</div><div class="vs">VS</div><div class="sub">${esc(dayLabel(g.day_tw).name)}</div>`;
    return `<article class="card ${live ? "is-live" : ""} ${post ? "is-post" : ""} ${edge ? "has-edge" : ""} ${open ? "open" : ""}" id="g-${esc(g.id)}" data-id="${esc(g.id)}">
      <div class="card-head"><span class="lg">${icon(SPORT_ICON[g.sport] || "i-grid")}${esc(g.league_zh)}${g.round ? ` · ${esc(g.round)}` : ""}</span>
        <span class="status ${live ? "live" : ""}">${live ? esc(g.status.detail) : post ? "已結束" : esc(g.time_tw) + " 開賽"}</span></div>
      <div class="teams">${teamBlock(g.away, g)}<div class="mid">${mid}</div>${teamBlock(g.home, g)}</div>
      ${oddsRow(g)}
      ${sugRow(g)}
      <button class="expand" data-toggle="${esc(g.id)}">${open ? "收起" : "展開傷兵、手感、近況、新聞"}${icon("i-chev")}</button>
      ${details(g)}
    </article>`;
  }

  function tennisSummary(tennis) {
    if (!tennis.length) return "";
    const by = {};
    tennis.forEach((g) => { const k = `${g.league_zh}｜${g.tournament || ""}`; (by[k] = by[k] || []).push(g); });
    const items = Object.keys(by).map((k) => {
      const list = by[k];
      const live = list.filter((g) => g.status.state === "in").length;
      const up = list.filter((g) => g.status.state === "pre").length;
      const [lg, tn] = k.split("｜");
      return `<button class="pick" data-sport-go="tennis"><span class="logos">${icon("i-tennis")}</span><span class="pt"><b>${esc(tn || lg)}</b><small>${esc(lg)} · ${live ? `進行中 ${live} 場 · ` : ""}未開賽 ${up} 場 · 共 ${list.length} 場單打</small></span><span class="pr"><span class="badge vio">看網球</span></span></button>`;
    }).join("");
    return `<div class="day-title">網球<small>點進去看各場次</small></div><div class="pick-grid">${items}</div>`;
  }

  function renderGames() {
    const all = gamesForSport();
    const isAll = state.sport === "all";
    const games = isAll ? all.filter((g) => !g.tennis) : all;
    const tennis = isAll ? all.filter((g) => g.tennis) : [];
    const live = games.filter((g) => g.status.state === "in");
    const liveSec = $("#sec-live");
    liveSec.classList.toggle("none", !live.length);
    $("#live").innerHTML = live.map(gameCard).join("");
    const rest = games.filter((g) => g.status.state !== "in");
    if (!rest.length && !live.length && !tennis.length) {
      $("#games").innerHTML = `<div class="empty">這個球種目前沒有賽事（往後 40 小時內）。</div>`;
      return;
    }
    const sortGames = (list) => list.sort((a, b) => (a.status.state === "post") - (b.status.state === "post") || a.start.localeCompare(b.start));
    let html = "";
    if (state.sport === "tennis") {
      const byT = {};
      rest.forEach((g) => { const k = `${g.league_zh}｜${g.tournament || ""}`; (byT[k] = byT[k] || []).push(g); });
      html = Object.keys(byT).map((k) => {
        const [lg, tn] = k.split("｜");
        const list = sortGames(byT[k]);
        return `<div class="day-title">${esc(tn || lg)}<small>${esc(lg)} · ${list.length} 場單打</small></div><div class="cards">${list.map(gameCard).join("")}</div>`;
      }).join("");
    } else {
      const byDay = {};
      rest.forEach((g) => { (byDay[g.day_tw] = byDay[g.day_tw] || []).push(g); });
      html = Object.keys(byDay).sort().map((day) => {
        const dl = dayLabel(day);
        const list = sortGames(byDay[day]);
        return `<div class="day-title">${esc(dl.name)}<small>${esc(dl.sub)} · ${list.length} 場</small></div><div class="cards">${list.map(gameCard).join("")}</div>`;
      }).join("");
    }
    $("#games").innerHTML = html + tennisSummary(tennis);
  }

  function renderInjuries() {
    const feed = (state.data.injury_feed || []).filter((x) => state.sport === "all" || x.sport === state.sport);
    const el = $("#injuries");
    if (!feed.length) { el.innerHTML = `<div class="empty">今明兩天的比賽沒有傷兵回報，或該球種沒有傷兵資料來源。</div>`; return; }
    const list = state.showAllInj ? feed : feed.slice(0, 10);
    el.innerHTML = list.map((x) => `<button class="aitem" data-goto="${esc(x.game_id)}">${avatarImg(x)}<div class="grow"><div class="pn">${esc(x.player)} ${x.key ? `<span class="badge vio">主力</span>` : ""}</div>
      <div class="ps">${x.team_logo ? `<img class="tl" src="${esc(x.team_logo)}" alt="">` : ""}${esc(x.team)} · ${esc(x.position || "")}${x.detail ? " · " + esc(x.detail) : ""} · ${esc(x.time_tw)}</div></div>
      <span class="badge ${sevClass(x.severity)}">${esc(x.status_zh)}</span></button>`).join("") +
      (feed.length > 10 ? `<button class="more" data-more="inj">${state.showAllInj ? "收起" : `還有 ${feed.length - 10} 筆，展開全部`}</button>` : "");
  }

  function renderHot() {
    const hot = (state.data.hot_players || []).filter((x) => state.sport === "all" || x.sport === state.sport);
    const el = $("#hot");
    if (!hot.length) { el.innerHTML = `<div class="empty">目前沒有明顯火燙或冰冷的球員（需要本季至少 3 場資料）。</div>`; return; }
    const list = state.showAllHot ? hot : hot.slice(0, 8);
    el.innerHTML = list.map((p) => {
      const cls = p.label === "火燙" ? "hot" : "cold";
      return `<button class="aitem" data-goto="${esc(p.game_id)}">${avatarImg(p)}<div class="grow"><div class="pn">${esc(p.player)}</div>
        <div class="ps">${p.team_logo ? `<img class="tl" src="${esc(p.team_logo)}" alt="">` : ""}${esc(p.team)} · 近 5 場 ${p.last5} ${esc(p.stat)}，季均 ${p.season}</div></div>
        <span class="badge ${cls}">${icon(cls === "hot" ? "i-fire" : "i-snow")}${esc(p.label)} ${p.delta_pct > 0 ? "+" : ""}${p.delta_pct}%</span></button>`;
    }).join("") + (hot.length > 8 ? `<button class="more" data-more="hot">${state.showAllHot ? "收起" : `還有 ${hot.length - 8} 人，展開全部`}</button>` : "");
  }

  function renderStandings() {
    const st = state.data.standings || {};
    const keys = Object.keys(st).filter((k) => state.sport === "all" || st[k].sport === state.sport);
    const el = $("#standings");
    if (!keys.length) { el.innerHTML = `<div class="empty">沒有排名資料。</div>`; return; }
    if (!state.standingsLeague || !keys.includes(state.standingsLeague)) state.standingsLeague = keys[0];
    const cur = st[state.standingsLeague];
    const sel = `<select class="sel" id="st-sel">${keys.map((k) => `<option value="${esc(k)}" ${k === state.standingsLeague ? "selected" : ""}>${esc(st[k].label || k)}</option>`).join("")}</select>`;
    const soccer = cur.sport === "soccer";
    const hockey = cur.sport === "hockey";
    const tables = (cur.groups || []).map((grp) => `${grp.name ? `<div class="grp">${esc(grp.name)}</div>` : ""}<table class="table"><thead><tr><th class="pos">#</th><th>球隊</th>${soccer ? `<th class="n">場</th><th class="n">勝-和-敗</th><th class="n">積分</th>` : hockey ? `<th class="n">勝-敗-延敗</th><th class="n">積分</th>` : `<th class="n">勝-敗</th><th class="n">勝率</th><th class="n">近況</th>`}</tr></thead><tbody>
      ${grp.rows.slice(0, 20).map((r) => `<tr><td class="pos">${r.pos}</td><td><span class="tm">${r.logo ? `<img src="${esc(r.logo)}" alt="">` : ""}${esc(r.name_zh || r.team || "")}</span></td>
        ${soccer ? `<td class="n">${r.gp ?? ""}</td><td class="n">${r.w ?? 0}-${r.d ?? 0}-${r.l ?? 0}</td><td class="n"><b>${r.points ?? ""}</b></td>`
          : hockey ? `<td class="n">${r.w ?? 0}-${r.l ?? 0}-${r.otl ?? 0}</td><td class="n"><b>${r.points ?? ""}</b></td>`
          : `<td class="n">${r.w ?? 0}-${r.l ?? 0}</td><td class="n">${esc(r.pct || "")}</td><td class="n">${esc(r.streak || "")}</td>`}</tr>`).join("")}
    </tbody></table>`).join("");
    el.innerHTML = sel + tables;
  }

  // ---------- 互動 ----------
  function gotoGame(id) {
    state.open.add(id);
    if (document.body.dataset.view !== "games" && window.innerWidth < 1100) setView("games");
    renderGames();
    const el = document.getElementById("g-" + id);
    if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
  }
  function setView(v) {
    state.view = v;
    document.body.dataset.view = v;
    document.querySelectorAll(".bottom-nav button").forEach((b) => b.classList.toggle("on", b.dataset.view === v));
    window.scrollTo({ top: 0 });
  }
  document.addEventListener("click", (e) => {
    const chip = e.target.closest(".chip, [data-sport-go]");
    if (chip) { state.sport = chip.dataset.sport || chip.dataset.sportGo; localStorage.setItem("sport", state.sport); render(); window.scrollTo({ top: 0 }); return; }
    const tg = e.target.closest("[data-toggle]");
    if (tg) {
      const id = tg.dataset.toggle;
      state.open.has(id) ? state.open.delete(id) : state.open.add(id);
      const card = tg.closest(".card");
      card.classList.toggle("open", state.open.has(id));
      tg.firstChild.textContent = state.open.has(id) ? "收起" : "展開傷兵、手感、近況、新聞";
      return;
    }
    const go = e.target.closest("[data-goto]");
    if (go) { gotoGame(go.dataset.goto); return; }
    const more = e.target.closest("[data-more]");
    if (more) { if (more.dataset.more === "inj") state.showAllInj = !state.showAllInj; else state.showAllHot = !state.showAllHot; renderInjuries(); renderHot(); return; }
    const nav = e.target.closest(".bottom-nav button");
    if (nav) { setView(nav.dataset.view); return; }
  });
  document.addEventListener("change", (e) => {
    if (e.target.id === "st-sel") { state.standingsLeague = e.target.value; localStorage.setItem("standingsLeague", e.target.value); renderStandings(); }
  });
  $("#btn-refresh").addEventListener("click", load);
  $("#btn-theme").addEventListener("click", () => {
    const cur = document.documentElement.dataset.theme === "light" ? "dark" : "light";
    document.documentElement.dataset.theme = cur;
    try { localStorage.setItem("theme", cur); } catch (_) { /* ignore */ }
    $("#btn-theme").innerHTML = icon(cur === "light" ? "i-moon" : "i-sun");
  });
  try {
    const saved = localStorage.getItem("theme");
    if (saved) { document.documentElement.dataset.theme = saved; $("#btn-theme").innerHTML = icon(saved === "light" ? "i-moon" : "i-sun"); }
  } catch (_) { /* ignore */ }

  load();
  setInterval(load, 5 * 60 * 1000);
})();
