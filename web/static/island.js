/* Dynamic Island: a liquid-glass pill at the top of the page that shows every update.
   Jobs: queued, started, each step with %, ready (with download), failed (with the reason),
   a new job from another typist (admin view).  System: welcome on sign in, API key missing,
   budget nearly / fully used, server offline and back online.  Pages can add their own:
   window.Island.flash(text, {kind: "ok"|"err"|"warn"|"info"|"run", sub, href, link});
   window.Island.refresh() polls now.  Scope comes from <body data-island-scope="mine|all">.
   Click the pill to open the panel: system status, running jobs and the recent updates feed. */
(function () {
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));
  const SHORT = {notes: "Notes", ppt: "PPT", test: "MCQ Test"};
  const ICON = {ok: "✓", err: "!", warn: "!", info: "i", run: "↻", queue: "⏳", off: "⚡"};
  const ss = {get(k) { try { return JSON.parse(sessionStorage.getItem("isl." + k)); } catch (e) { return null; } },
              set(k, v) { try { sessionStorage.setItem("isl." + k, JSON.stringify(v)); } catch (e) {} }};
  const el = document.createElement("div");
  el.id = "island"; el.className = "island idle k-ok"; el.setAttribute("role", "status"); el.setAttribute("aria-live", "polite"); el.tabIndex = 0;
  el.innerHTML = `
    <i class="isl-liquid"></i><i class="isl-shine"></i>
    <div class="isl-view isl-idle"><img class="isl-logo" src="/static/disha-logo.png" alt=""><i class="isl-dot"></i><span></span></div>
    <div class="isl-view isl-active"><svg class="isl-ring" viewBox="0 0 24 24"><circle class="bg" cx="12" cy="12" r="9.5"/><circle class="fg" cx="12" cy="12" r="9.5" stroke-dasharray="59.7" stroke-dashoffset="45"/></svg>
      <div class="tx"><b class="t"></b><small class="s"></small></div><span class="p"></span><span class="isl-count"></span></div>
    <div class="isl-view isl-open"></div>
    <div class="isl-view isl-flash"></div>`;
  document.body.appendChild(el);
  const V = n => el.querySelector(".isl-" + n);
  const scope = document.body.dataset.islandScope || "mine";
  const ADMIN = scope === "all", HDR = ADMIN ? {"X-Panel": "admin"} : {};  // admin panel uses its own login
  let state = "idle", kind = "ok", active = [], recent = [], known = null, steps = {};
  let health = {online: true, api_key: true, budget: null, user: null}, fails = 0;
  let flashTimer = null, openTimer = null, queue = [];
  let feed = ss.get("feed") || [];

  // progress 0..1 from the job's step text, or null when it can't be told
  function progress(j) {
    if (j.status === "queued") return null;
    const s = j.step || "";
    let m = s.match(/PDF (\d+)\/(\d+).*?slides \d+-(\d+) of (\d+)/);
    if (m) return ((+m[1] - 1) + (+m[3]) / (+m[4])) / (+m[2]) * 0.95;
    m = s.match(/PDF (\d+)\/(\d+)/);
    if (m) return ((+m[1] - 1) / (+m[2])) * 0.95 + 0.01;
    if (/Building the Word notes|Building the PPT/.test(s)) return 0.97;
    if (/Writing \d+ MCQs/.test(s)) return 0.5;
    if (/Reading files/.test(s)) return 0.08;
    if (/Starting/.test(s)) return 0.02;
    return null;
  }
  const name = j => j.title || j.chapter || (j.files && j.files[0]) || "Task";
  const tool = j => SHORT[j.task] || "Task";
  const ago = t => { const s = Math.max(0, (Date.now() - t) / 1000);
    return s < 50 ? "just now" : s < 3600 ? Math.round(s / 60) + " min ago" : Math.round(s / 3600) + " h ago"; };

  // what the resting pill says when no job is running
  function idleInfo() {
    if (!health.online) return ["off", "Offline · retrying"];
    if (!health.api_key) return ["warn", ADMIN ? "API key missing" : "Notes & MCQ paused"];
    const b = health.budget;
    if (b && b.limit_inr > 0) {
      const pc = b.spent_inr / b.limit_inr;
      if (pc >= 1) return ["err", ADMIN ? "Budget used up" : "Monthly limit reached"];
      if (pc >= 0.8 && ADMIN) return ["warn", `Budget ${Math.round(pc * 100)}% used`];
    }
    return ["ok", "All clear"];
  }

  function size() {
    const vw = window.innerWidth, max = Math.min(vw - 24, 400);
    const room = vw < 760 ? vw - 132 : vw - 24;  // phones: keep clear of the logo and avatar
    if (state === "idle") {
      const t = V("idle").querySelector("span"); t.style.position = "absolute";
      const w = 64 + t.scrollWidth; t.style.position = "";
      return [Math.min(Math.max(132, w), room), 36, 22];
    }
    if (state === "active") return [Math.min(room, 320), 46, 24];
    const view = V(state === "open" ? "open" : "flash");
    view.style.width = max + "px";
    return [max, Math.max(60, Math.min(view.scrollHeight, window.innerHeight - 40)), 28];
  }
  function set(s, k) {
    const morph = s !== state || (k || kind) !== kind;
    state = s; kind = k || kind;
    el.className = `island ${s} k-${kind}`;
    const [w, h, r] = size();
    el.style.width = w + "px"; el.style.height = h + "px"; el.style.borderRadius = r + "px";
    if (morph) { el.classList.remove("blob"); void el.offsetWidth; el.classList.add("blob"); }
  }
  function base() {
    if (queue.length) return show(queue.shift());
    if (active.length) { drawActive(); return set("active", active[0].status === "queued" ? "queue" : "run"); }
    const [k, t] = idleInfo();
    V("idle").querySelector("span").textContent = t;
    set("idle", k);
  }

  function drawActive() {
    const j = active[0];
    if (!j) return;
    const p = progress(j), ring = el.querySelector(".isl-ring"), fg = ring.querySelector(".fg"), A = V("active");
    ring.classList.toggle("spin", p === null);
    fg.style.strokeDashoffset = 59.7 * (1 - (p === null ? 0.25 : Math.max(0.04, p)));
    A.querySelector(".t").textContent = `${tool(j)} · ${name(j)}`;
    A.querySelector(".s").textContent = j.status === "queued" ? "Queued · waiting in line" : (j.step || "Working");
    A.querySelector(".p").textContent = p === null ? "" : Math.round(p * 100) + "%";
    const c = A.querySelector(".isl-count");
    c.textContent = "+" + (active.length - 1); c.style.display = active.length > 1 ? "" : "none";
  }

  const chip = (k, t) => `<span class="isl-chip c-${k}">${esc(t)}</span>`;
  function drawOpen() {
    const b = health.budget, pc = b && b.limit_inr > 0 ? Math.round(b.spent_inr / b.limit_inr * 100) : null;
    const sys = `<div class="isl-sys">
      ${chip(health.online ? "ok" : "off", health.online ? "Server online" : "Server offline")}
      ${ADMIN ? chip(health.api_key ? "ok" : "warn", health.api_key ? "API key set" : "API key missing")
        : chip(health.api_key && !(pc >= 100) ? "ok" : "warn", health.api_key && !(pc >= 100) ? "All tools ready" : "Notes & MCQ paused · contact Ashique Sir")}
      ${pc === null || !ADMIN ? "" : chip(pc >= 100 ? "err" : pc >= 80 ? "warn" : "ok", `Budget ${pc}% · ₹${b.spent_inr.toLocaleString("en-IN")} of ₹${Math.round(b.limit_inr).toLocaleString("en-IN")}`)}</div>`;
    const jobs = active.length ? active.slice(0, 4).map(j => { const p = progress(j), q = j.status === "queued";
        return `<div class="isl-job"><div class="row"><b>${esc(tool(j))} · ${esc(name(j))}</b>${chip(q ? "queue" : "run", q ? "Queued" : p === null ? "Running" : Math.round(p * 100) + "%")}</div>
          <small>${scope === "all" && j.by_name ? esc(j.by_name) + " · " : ""}${esc(q ? "Waiting in line" : j.step || "")}</small>
          <div class="isl-bar${p === null ? " ind" : ""}${q ? " q" : ""}"><i style="width:${p === null ? 0 : Math.round(p * 100)}%"></i></div></div>`; }).join("")
      : `<div class="isl-empty">Nothing is running right now.</div>`;
    const log = feed.length ? feed.slice(0, 6).map(f => `<div class="isl-log"><i class="ic c-${f.k}">${ICON[f.k] || "i"}</i>
        <div><b>${esc(f.t)}</b>${f.s ? `<small>${esc(f.s)}</small>` : ""}</div><time>${ago(f.at)}</time></div>`).join("")
      : `<div class="isl-empty">No updates yet. Finished files are in Work history.</div>`;
    V("open").innerHTML = sys +
      `<h4><span>${scope === "all" ? "Running now (everyone)" : "My tasks in progress"}</span><span>${active.length}</span></h4>${jobs}` +
      `<h4><span>Recent updates</span></h4>${log}`;
  }

  function show(o) {
    clearTimeout(flashTimer);
    V("flash").innerHTML = `<div class="ic">${ICON[o.kind] || "i"}</div><div class="msg"><b>${esc(o.text)}</b>${o.sub ? `<small>${esc(o.sub)}</small>` : ""}</div>` +
      (o.href ? `<a href="${esc(o.href)}">${esc(o.link || "Download")}</a>` : "");
    set("flash", o.kind);
    flashTimer = setTimeout(() => { if (state === "flash") base(); }, o.href ? 7000 : o.kind === "err" ? 6000 : (o.kind === "run" || o.kind === "queue") ? 2600 : 3800);
  }
  // show a message now (or after the one on screen when quiet) and keep it in the feed
  function flash(text, o = {}) {
    o = Object.assign({kind: "info"}, o, {text});
    if (!o.quietLog) { feed.unshift({k: o.kind, t: text, s: o.sub || "", at: Date.now()}); feed = feed.slice(0, 20); ss.set("feed", feed); }
    if (state === "open") { drawOpen(); return set("open"); }
    if (o.wait && state === "flash") { queue.push(o); return; }
    clearTimeout(openTimer); show(o);
  }

  async function checkHealth(first) {
    try {
      const r = await fetch("/api/me", {headers: HDR});
      if (!r.ok) return;
      const d = await r.json(), seen = ss.get("seen") || {};
      health.api_key = !!d.api_key; health.budget = d.budget; health.user = d.user;
      const b = d.budget, pc = b && b.limit_inr > 0 ? b.spent_inr / b.limit_inr : 0;
      if (first && !seen.welcome && d.user) {
        seen.welcome = 1;
        const w = String(d.user.name || d.user.username || "").trim().split(/\s+/),
          first = w.length > 1 && /^(md|mohd|mohammad|muhammad|mr|mrs|ms|dr|sri|shri)\.?$/i.test(w[0]) ? w[1] : w[0];
        flash(`Welcome, ${first}`, {kind: "ok", sub: "Signed in to Disha Online Solution", wait: true});
      }
      if (!d.api_key && !seen.key) { seen.key = 1;
        flash(ADMIN ? "API key missing" : "Notes & MCQ paused", {kind: "warn", sub: ADMIN ? "Settings → API integration" : "Please contact Ashique Sir to resolve this", wait: true}); }
      const lvl = pc >= 1 ? 2 : pc >= 0.8 ? 1 : 0;
      if (lvl > (seen.budget || 0) && (ADMIN || lvl === 2)) flash(!ADMIN ? "Monthly limit reached" : lvl === 2 ? "Monthly budget used up" : `Budget ${Math.round(pc * 100)}% used`,
        {kind: lvl === 2 ? "err" : "warn", sub: !ADMIN ? "Notes and MCQ are paused. Please contact Ashique Sir" : lvl === 2 ? "New AI tasks are paused until the limit is raised" : `₹${b.spent_inr} of ₹${Math.round(b.limit_inr)} this month`, wait: true});
      seen.budget = lvl;
      ss.set("seen", seen);
    } catch (e) {}
  }

  async function refresh() {
    if (document.hidden) return;
    let d;
    try {
      const r = await fetch(`/api/jobs?scope=${scope}&per=15`, {headers: HDR});
      if (r.status === 401) return;
      if (!r.ok) throw new Error(r.status);
      d = await r.json();
      if (!health.online) { health.online = true; flash("Back online", {kind: "ok", sub: "Live updates are on again"}); checkHealth(); }
      fails = 0;
    } catch (e) {
      if (++fails >= 2 && health.online) { health.online = false; flash("Server offline", {kind: "off", sub: "Check that start-windows.bat is running. Retrying…"}); }
      return;
    }
    const jobs = d.jobs || [];
    active = jobs.filter(j => j.status === "queued" || j.status === "running").reverse();
    if (known) {
      for (const j of jobs.slice().reverse()) {
        const was = known[j.id], who = scope === "all" && j.by_name ? j.by_name + " · " : "";
        if (!was && (j.status === "queued" || j.status === "running"))
          flash(`${tool(j)} task queued`, {kind: "queue", sub: who + name(j), wait: true});
        else if (was === "queued" && j.status === "running")
          flash(`${tool(j)} started`, {kind: "run", sub: who + name(j), wait: true});
        if ((was === "queued" || was === "running" || (!was && known)) && (j.status === "done" || j.status === "failed") && was !== j.status) {
          if (!was && !(Date.now() / 1000 - (j.finished || 0) < 60)) continue;
          if (j.status === "done") {
            const dl = j.downloads && j.downloads[0];
            flash(`${tool(j)} ready`, {kind: "ok", sub: who + name(j), href: dl && `/api/jobs/${j.id}/download?kind=${dl.kind}${scope === "all" ? "&panel=admin" : ""}`, link: dl && dl.label});
          } else flash(`${tool(j)} failed`, {kind: "err", sub: (j.error || name(j)) + (scope === "all" ? "" : " · Retry from Work history")});
          document.dispatchEvent(new CustomEvent("island:jobchange", {detail: j}));
        } else if (j.status === "running" && steps[j.id] && steps[j.id] !== j.step) {
          document.dispatchEvent(new CustomEvent("island:jobstep", {detail: j}));
        }
      }
    }
    known = Object.fromEntries(jobs.map(j => [j.id, j.status]));
    steps = Object.fromEntries(jobs.map(j => [j.id, j.step]));
    drawActive();
    if (state === "open") { drawOpen(); set("open"); }
    else if (state !== "flash") base();
  }

  el.addEventListener("click", e => {
    if (e.target.closest("a")) return;
    if (state === "open") return base();
    clearTimeout(flashTimer); queue = []; drawOpen(); set("open", active.length ? (active[0].status === "queued" ? "queue" : "run") : idleInfo()[0]);
    clearTimeout(openTimer); openTimer = setTimeout(() => state === "open" && base(), 12000);
  });
  el.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); el.click(); } });
  document.addEventListener("click", e => { if (state === "open" && !el.contains(e.target)) base(); });
  window.addEventListener("resize", () => set(state));
  window.addEventListener("offline", () => refresh());
  window.addEventListener("online", () => refresh());
  document.addEventListener("visibilitychange", refresh);

  window.Island = {flash, refresh, health: () => Object.assign({}, health)};
  base();
  checkHealth(true).then(refresh);
  setInterval(refresh, 4000);
  setInterval(() => !document.hidden && checkHealth(), 30000);
})();
