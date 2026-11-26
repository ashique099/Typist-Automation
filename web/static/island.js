/* Dynamic Island: shows live job progress and notifications in a pill at the top.
   window.Island.flash(text, {kind: "ok"|"err"|"info", sub, href, link}) shows a message;
   window.Island.refresh() polls jobs now.  Scope comes from <body data-island-scope="mine|all">. */
(function () {
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({"&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"}[c]));
  const SHORT = {notes: "Notes", ppt: "PPT", test: "MCQ Test"};
  const el = document.createElement("div");
  el.id = "island"; el.className = "island idle"; el.setAttribute("role", "status"); el.setAttribute("aria-live", "polite"); el.tabIndex = 0;
  el.innerHTML = `
    <div class="isl-view isl-idle"><img class="isl-logo" src="/static/disha-logo.png" alt=""><i class="isl-dot"></i><span>All clear</span></div>
    <div class="isl-view isl-active"><svg class="isl-ring" viewBox="0 0 24 24"><circle class="bg" cx="12" cy="12" r="9.5"/><circle class="fg" cx="12" cy="12" r="9.5" stroke-dasharray="59.7" stroke-dashoffset="45"/></svg>
      <span class="t"></span><span class="p"></span><span class="isl-count"></span></div>
    <div class="isl-view isl-open"></div>
    <div class="isl-view isl-flash"></div>`;
  document.body.appendChild(el);
  const V = n => el.querySelector(".isl-" + n);
  const scope = document.body.dataset.islandScope || "mine";
  let state = "idle", active = [], known = null, flashTimer = null, openTimer = null;

  // progress 0..1 from the job's step text, or null when it can't be told
  function progress(j) {
    const s = j.step || "";
    let m = s.match(/PDF (\d+)\/(\d+).*?slides \d+-(\d+) of (\d+)/);
    if (m) return ((+m[1] - 1) + (+m[3]) / (+m[4])) / (+m[2]) * 0.95;
    m = s.match(/PDF (\d+)\/(\d+)/);
    if (m) return ((+m[1] - 1) / (+m[2])) * 0.95 + 0.01;
    if (/Building the Word notes/.test(s)) return 0.97;
    return null;
  }
  const name = j => j.title || j.chapter || (j.files && j.files[0]) || "Job";

  function size() {
    const vw = window.innerWidth, max = Math.min(vw - 24, 400);
    const room = vw < 760 ? vw - 132 : vw - 24;  // phones: keep clear of the logo and avatar
    if (state === "idle") return [Math.min(132, room), 36, 22];
    if (state === "active") return [Math.min(room, 300), 40, 22];
    const view = V(state === "open" ? "open" : "flash");
    view.style.width = max + "px";
    return [max, Math.max(60, view.scrollHeight), 26];
  }
  function set(s, kind) {
    state = s;
    el.className = "island " + s + (kind ? " " + kind : "");
    const [w, h, r] = size();
    el.style.width = w + "px"; el.style.height = h + "px"; el.style.borderRadius = r + "px";
    el.classList.remove("blob"); void el.offsetWidth; el.classList.add("blob");
  }
  function base() { set(active.length ? "active" : "idle"); }

  function drawActive() {
    const j = active[0];
    if (!j) return;
    const p = progress(j), ring = el.querySelector(".isl-ring"), fg = ring.querySelector(".fg");
    ring.classList.toggle("spin", p === null);
    fg.style.strokeDashoffset = 59.7 * (1 - (p === null ? 0.25 : Math.max(0.04, p)));
    V("active").querySelector(".t").textContent = `${SHORT[j.task] || "Job"} · ${name(j)}`;
    V("active").querySelector(".p").textContent = p === null ? (j.status === "queued" ? "Queued" : "…") : Math.round(p * 100) + "%";
    const c = V("active").querySelector(".isl-count");
    c.textContent = "+" + (active.length - 1); c.style.display = active.length > 1 ? "" : "none";
  }
  function drawOpen() {
    V("open").innerHTML = `<h4><span>${scope === "all" ? "Running now (everyone)" : "My running jobs"}</span><span>${active.length}</span></h4>` +
      (active.length ? active.slice(0, 4).map(j => { const p = progress(j);
        return `<div class="isl-job"><b>${esc(SHORT[j.task] || "Job")} · ${esc(name(j))}</b>
          <small>${scope === "all" && j.by_name ? esc(j.by_name) + " · " : ""}${esc(j.status === "queued" ? "Waiting in line" : j.step || "")}</small>
          <div class="isl-bar${p === null ? " ind" : ""}"><i style="width:${p === null ? 0 : Math.round(p * 100)}%"></i></div></div>`; }).join("")
        : `<div class="isl-empty">Nothing is running. Finished files are in Work history.</div>`);
  }

  function flash(text, o = {}) {
    clearTimeout(flashTimer); clearTimeout(openTimer);
    const kind = o.kind || "info", icon = {ok: "✓", err: "!", info: "i"}[kind];
    V("flash").innerHTML = `<div class="ic">${icon}</div><div class="msg"><b>${esc(text)}</b>${o.sub ? `<small>${esc(o.sub)}</small>` : ""}</div>` +
      (o.href ? `<a href="${esc(o.href)}">${esc(o.link || "Download")}</a>` : "");
    set("flash", kind === "info" ? "" : kind);
    flashTimer = setTimeout(base, o.href ? 7000 : 3800);
  }

  async function refresh() {
    if (document.hidden) return;
    let d;
    try {
      const r = await fetch(`/api/jobs?scope=${scope}&per=15`);
      if (!r.ok) return;
      d = await r.json();
    } catch (e) { return; }
    const jobs = d.jobs || [];
    active = jobs.filter(j => j.status === "queued" || j.status === "running").reverse();
    if (known) {
      for (const j of jobs) {
        const was = known[j.id];
        if ((was === "queued" || was === "running") && (j.status === "done" || j.status === "failed")) {
          if (j.status === "done") {
            const dl = j.downloads && j.downloads[0];
            flash(`${SHORT[j.task] || "Job"} ready`, {kind: "ok", sub: name(j), href: dl && `/api/jobs/${j.id}/download?kind=${dl.kind}`, link: dl && dl.label});
          } else flash(`${SHORT[j.task] || "Job"} failed`, {kind: "err", sub: j.error || name(j)});
          document.dispatchEvent(new CustomEvent("island:jobchange", {detail: j}));
        }
      }
    }
    known = Object.fromEntries(jobs.map(j => [j.id, j.status]));
    drawActive();
    if (state === "open") { drawOpen(); set("open"); }
    else if (state !== "flash") base();
  }

  el.addEventListener("click", e => {
    if (e.target.closest("a")) return;
    if (state === "open") return base();
    clearTimeout(flashTimer); drawOpen(); set("open");
    clearTimeout(openTimer); openTimer = setTimeout(() => state === "open" && base(), 9000);
  });
  el.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); el.click(); } });
  document.addEventListener("click", e => { if (state === "open" && !el.contains(e.target)) base(); });
  window.addEventListener("resize", () => set(state));
  document.addEventListener("visibilitychange", refresh);

  window.Island = {flash, refresh};
  set("idle");
  refresh();
  setInterval(refresh, 4000);
})();
