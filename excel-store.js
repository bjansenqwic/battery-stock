/* Battery Stock – connection to the Excel file.
 *
 * The page itself is the same as the version inside Claude. This file gives it what it needs:
 *  - sign-in with the Microsoft work account (MSAL),
 *  - reading and writing the Excel workbook in OneDrive/SharePoint (Microsoft Graph),
 *  - a live camera scanner.
 *
 * The workbook has three tables:
 *  Movements – one row per check-in, check-out or type change. Rows are only added, never edited,
 *              so several people can work at the same time. Current stock is worked out from them.
 *  Types     – battery types and the start of their serial numbers.
 *  Locations – storage spots, their maximum and what they are for.
 *
 * Without clientId and workbookUrl in config.js the app runs in demo mode: data stays in this browser.
 */
(() => {
  "use strict";
  const CFG = Object.assign({ clientId: "", tenantId: "", workbookUrl: "" }, window.BATTERY_CONFIG || {});
  const DEMO = !CFG.clientId || !CFG.workbookUrl;
  const TABLE = { mov: "Movements", types: "Types", locs: "Locations" };
  const COLS = [["Date", "date"], ["Movement", "mv"], ["Serial number", "sn"], ["Type", "model"], ["Battery status", "condition"],
    ["Battery is for", "purpose"], ["In box", "inBox"], ["Location", "loc"], ["Reason", "reason"], ["Reference nr", "ref"],
    ["Note", "note"], ["By", "by"], ["Email", "email"], ["Timestamp", "iso"], ["ID", "id"]];
  const ONLY_TO_TEXT = { defect: "Defect", marketing: "Marketing", any: "Anything" };
  const TEXT_TO_ONLY = t => /^(defect|broken)/i.test(t) ? "defect" : /^marketing/i.test(t) ? "marketing" : /^(any|anything|temporary)/i.test(t) ? "any" : "";

  const p2 = n => String(n).padStart(2, "0");
  const localText = ms => { const d = new Date(ms); return `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}`; };
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
  const norm = s => String(s == null ? "" : s).trim().replace(/\s+/g, "").toUpperCase();
  const docId = sn => { let id = sn.replace(/[^A-Za-z0-9_\-.~:@+]/g, "_").slice(0, 190); if (id === "." || id === "..") id = "_" + id; return id; };
  const yn = v => { const s = String(v == null ? "" : v).trim().toLowerCase(); return ["yes", "y", "ja", "true", "1"].includes(s) ? "yes" : ["no", "n", "nee", "false", "0"].includes(s) ? "no" : ""; };
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const txt = v => (v == null ? "" : String(v)).replace(/^'/, "").trim();
  // Keep serial numbers as text in Excel and never let a cell start a formula.
  const cell = v => { if (typeof v === "number") return v; const s = v == null ? "" : String(v); return /^[=+\-@]/.test(s) || /^[\d.,\s]+(e\d+)?$/i.test(s) && s !== "" ? "'" + s : s; };

  let me = { name: "Demo user", email: "demo@example.com" };
  const names = new Map();

  // ---------- a message bar for connection problems ----------
  function problem(msg) {
    let bar = document.getElementById("xlProblem");
    if (!msg) { if (bar) bar.remove(); return; }
    if (!bar) {
      bar = document.createElement("div"); bar.id = "xlProblem"; bar.setAttribute("role", "alert");
      bar.style.cssText = "position:fixed;left:50%;transform:translateX(-50%);top:calc(10px + env(safe-area-inset-top,0px));z-index:70;max-width:calc(100% - 32px);background:var(--danger,#b42318);color:#fff;padding:10px 14px;border-radius:10px;font:500 14px system-ui,sans-serif;box-shadow:0 4px 16px rgba(0,0,0,.2)";
      document.body.append(bar);
    }
    bar.textContent = msg;
  }
  const explain = e => e && e.status === 404 ? "Couldn't find the Excel file or one of its tables (Movements, Types, Locations). Check the link in config.js."
    : e && e.status === 403 ? "You don't have access to the Excel file. Ask for edit access to the file."
    : e && e.status === 423 ? "The Excel file is locked. Close it in the desktop Excel app; Excel for the web is fine."
    : e && e.userMessage ? e.userMessage
    : "Couldn't reach the Excel file. Check your internet connection." + (e && e.message ? " (" + e.message + ")" : "");

  // ---------- demo backend: this browser only ----------
  const demo = {
    read(k, d) { try { return JSON.parse(localStorage.getItem(k) || "null") || d; } catch { return d; } },
    async movements() { return this.read("bs.mov", []); },
    async append(rows) { const a = this.read("bs.mov", []); a.push(...rows); localStorage.setItem("bs.mov", JSON.stringify(a)); },
    async remove(ids) { const s = new Set(ids); localStorage.setItem("bs.mov", JSON.stringify(this.read("bs.mov", []).filter(r => !s.has(r.id)))); },
    async settings() { return this.read("bs.set", window.BATTERY_DEFAULTS || { types: [], locations: [] }); },
    async saveSettings(v) { localStorage.setItem("bs.set", JSON.stringify({ types: v.types, locations: v.locations })); },
  };

  // ---------- Excel backend (Microsoft Graph) ----------
  const SCOPES = ["User.Read", "Files.ReadWrite.All"];
  let pca = null, account = null;
  async function token() {
    try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) { await pca.acquireTokenRedirect({ scopes: SCOPES, account }); throw e; }
  }
  async function graph(path, opts = {}, tries = 5) {
    const url = path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path;
    for (let i = 0; ; i++) {
      let res;
      try { res = await fetch(url, { method: opts.method || "GET", headers: { Authorization: "Bearer " + await token(), "Content-Type": "application/json" }, body: opts.body ? JSON.stringify(opts.body) : undefined }); }
      catch (e) { if (i < tries - 1) { await sleep(800 * 2 ** i); continue; } throw e; }
      if (res.ok) return res.status === 204 ? null : res.json().catch(() => null);
      if ([409, 423, 429, 500, 502, 503, 504].includes(res.status) && i < tries - 1) {
        const ra = Number(res.headers.get("Retry-After")); await sleep(ra > 0 ? ra * 1000 : 700 * 2 ** i + Math.random() * 400); continue;
      }
      let detail = ""; try { const j = await res.json(); detail = (j.error && (j.error.message || j.error.code)) || ""; } catch {}
      const err = new Error(detail || "HTTP " + res.status); err.status = res.status; throw err;
    }
  }
  const xl = {
    root: null, webUrl: null, head: null,
    async resolve() {
      if (this.root) return;
      const b64 = btoa(unescape(encodeURIComponent(CFG.workbookUrl.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
      const item = await graph(`/shares/u!${b64}/driveItem?$select=id,name,webUrl,parentReference`);
      this.root = `/drives/${item.parentReference.driveId}/items/${item.id}/workbook`;
      this.webUrl = item.webUrl;
      const a = document.getElementById("xlOpen"); if (a && this.webUrl) { a.href = this.webUrl; a.hidden = false; }
    },
    t(name) { return `${this.root}/tables/${encodeURIComponent(name)}`; },
    async values(name) { await this.resolve(); const r = await graph(this.t(name) + "/range?$select=values"); return (r && r.values) || []; },
    async movements() {
      const vals = await this.values(TABLE.mov);
      const head = (vals[0] || []).map(h => String(h).trim().toLowerCase()); this.head = head;
      const ix = {}; for (const [h, f] of COLS) ix[f] = head.indexOf(h.toLowerCase());
      if (ix.sn < 0 || ix.mv < 0 || ix.id < 0) throw Object.assign(new Error("columns"), { userMessage: 'The Movements table needs the columns "Serial number", "Movement" and "ID". Use the template workbook.' });
      const out = [];
      vals.slice(1).forEach((row, i) => {
        const g = f => ix[f] >= 0 ? txt(row[ix[f]]) : "";
        const sn = norm(g("sn")), mv = g("mv").toUpperCase(); if (!sn || !mv) return;
        let at = Date.parse(g("iso"));
        if (!Number.isFinite(at)) { const d = ix.date >= 0 ? row[ix.date] : ""; at = typeof d === "number" ? Math.round((d - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000 : Date.parse(String(d).replace(" ", "T")); }
        if (!Number.isFinite(at)) at = i;
        out.push({ id: g("id") || "row" + i, at, mv: mv.startsWith("O") ? "OUT" : mv.startsWith("T") ? "TYPE" : "IN", sn, model: g("model"), condition: g("condition"),
          purpose: g("purpose"), inBox: g("inBox"), loc: g("loc"), reason: g("reason"), ref: g("ref"), note: g("note"), by: g("by"), email: g("email").toLowerCase() });
      });
      return out;
    },
    async append(rows) {
      if (!this.head) await this.movements();
      const values = rows.map(r => this.head.map(h => { const c = COLS.find(x => x[0].toLowerCase() === h); return c ? cell(r[c[1]]) : ""; }));
      await graph(this.t(TABLE.mov) + "/rows", { method: "POST", body: { index: null, values } });
    },
    async remove(ids) {
      if (!ids.length) return;
      const r = await graph(this.t(TABLE.mov) + "/columns/ID/dataBodyRange?$select=values");
      const want = new Set(ids), idx = [];
      ((r && r.values) || []).forEach((v, i) => { if (want.has(txt(v[0]))) idx.push(i); });
      idx.sort((a, b) => b - a);   // delete from the bottom so the other positions stay right
      for (const i of idx) await graph(this.t(TABLE.mov) + `/rows/itemAt(index=${i})`, { method: "DELETE" });
    },
    async settings() {
      const tv = await this.values(TABLE.types), lv = await this.values(TABLE.locs);
      const types = tv.slice(1).map(r => ({ name: txt(r[0]), codes: txt(r[1]).split(/[,;]/).map(x => x.trim().toUpperCase()).filter(Boolean) })).filter(t => t.name);
      const locations = lv.slice(1).map(r => ({ name: txt(r[0]), cap: Math.max(0, parseInt(txt(r[1]), 10) || 0), only: TEXT_TO_ONLY(txt(r[2])) })).filter(l => l.name);
      return { types, locations };
    },
    async saveSettings(v) {
      await this.writeTable(TABLE.types, v.types.map(t => [t.name, (t.codes || []).join(", ")]));
      await this.writeTable(TABLE.locs, v.locations.map(l => [l.name, l.cap ? l.cap : "", ONLY_TO_TEXT[l.only] || ""]));
    },
    // Make a table hold exactly these rows: overwrite, then add or remove rows at the end.
    async writeTable(name, rows) {
      await this.resolve();
      const body = await graph(this.t(name) + "/dataBodyRange?$select=address,rowCount,columnCount");
      const count = body.rowCount || 0, cols = body.columnCount || 1;
      const pad = r => { const a = r.slice(0, cols).map(cell); while (a.length < cols) a.push(""); return a; };
      const [sheetRaw, addr] = body.address.includes("!") ? body.address.split("!") : ["", body.address];
      const sheet = sheetRaw.replace(/^'|'$/g, "").replace(/''/g, "'");
      const m = addr.replace(/\$/g, "").match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      const n = Math.min(count, Math.max(rows.length, 1));
      const first = rows.length ? rows.slice(0, n).map(pad) : [pad([])];
      await graph(`${this.root}/worksheets/${encodeURIComponent(sheet)}/range(address='${m[1]}${m[2]}:${m[3]}${+m[2] + n - 1}')`, { method: "PATCH", body: { values: first } });
      if (rows.length > count) await graph(this.t(name) + "/rows", { method: "POST", body: { index: null, values: rows.slice(count).map(pad) } });
      for (let i = count - 1; i >= Math.max(rows.length, 1); i--) await graph(this.t(name) + `/rows/itemAt(index=${i})`, { method: "DELETE" });
    },
  };
  const backend = DEMO ? demo : xl;

  // ---------- stock worked out from the movements ----------
  function buildDocs(movs) {
    const list = movs.slice().sort((a, b) => a.at - b.at), m = new Map();
    for (const r of list) {
      if (r.email && r.by) names.set(r.email, r.by);
      const id = docId(r.sn); let b = m.get(id);
      if (r.mv === "TYPE") { if (b) { b.model = r.model; b.typeRows = (b.typeRows || []).concat(r.id); } continue; }
      const e = { id: r.id, t: r.mv === "OUT" ? "out" : "in", at: r.at, loc: r.loc, model: r.model, reason: r.reason, note: r.note, ref: r.ref,
        condition: r.condition, inBox: yn(r.inBox), purpose: r.purpose, by: r.email || r.by || null };
      if (!b) { b = { sn: r.sn, model: "", firstSeen: r.at, lastIn: null, lastOut: null, location: "", lastLocation: "", history: [] }; m.set(id, b); }
      b.history.push(e); b.status = e.t; if (e.model) b.model = e.model;
      Object.assign(b, { reason: e.reason, note: e.note, ref: e.ref, condition: e.condition, inBox: e.inBox, purpose: e.purpose, updatedAt: e.at, updatedBy: e.by });
      b.firstSeen = Math.min(b.firstSeen, e.at);
      if (e.t === "in") { b.lastIn = e.at; b.location = e.loc; if (e.loc) b.lastLocation = e.loc; }
      else { b.lastOut = e.at; b.lastLocation = e.loc || b.location || b.lastLocation; b.location = ""; }
    }
    return m;
  }
  const toRow = (sn, e, mv) => ({ id: uid(), at: e.at || Date.now(), date: localText(e.at || Date.now()), iso: new Date(e.at || Date.now()).toISOString(),
    mv: mv || (e.t === "out" ? "OUT" : "IN"), sn, model: e.model || "", condition: e.condition || "", purpose: e.purpose || "",
    inBox: e.inBox === "yes" ? "Yes" : e.inBox === "no" ? "No" : "", loc: e.loc || "", reason: e.reason || "", ref: e.ref || "", note: e.note || "",
    by: me.name, email: me.email });

  // ---------- the store the page talks to ----------
  const st = { docs: new Map(), movs: [], settings: null, subs: [], ssubs: [], loaded: false };
  function fire() {
    const docs = [...st.docs].map(([id, v]) => ({ id, exists: true, data: () => v, metadata: { fromCache: false, hasPendingWrites: false } }));
    const snap = { docs, size: docs.length, empty: !docs.length, docChanges: () => [], metadata: { fromCache: false, hasPendingWrites: false } };
    for (const f of st.subs) { try { f(snap); } catch (e) { console.error(e); } }
  }
  function fireSettings() {
    const v = st.settings; const d = { id: "main", exists: !!v, data: () => v, metadata: { fromCache: false, hasPendingWrites: false } };
    for (const f of st.ssubs) { try { f(d); } catch (e) { console.error(e); } }
  }
  async function refresh() {
    st.movs = await backend.movements(); st.docs = buildDocs(st.movs); st.loaded = true; problem(""); fire();
  }
  async function refreshSettings() { st.settings = await backend.settings(); fireSettings(); }
  let busy = 0;
  async function write(fn) {
    busy++;
    try { await fn(); } catch (e) { problem(explain(e)); setTimeout(() => problem(""), 8000); throw Object.assign(new Error("save failed"), { code: "unavailable" }); }
    finally { busy--; }
    await refresh().catch(() => {});
  }
  async function setBattery(id, next) {
    await write(async () => {
      const movs = await backend.movements(), cur = buildDocs(movs).get(id) || null;
      const nextHist = Array.isArray(next.history) ? next.history : [];
      const kept = nextHist.filter(e => e.id), newest = kept.reduce((m, e) => Math.max(m, e.at || 0), -Infinity);
      const keptIds = new Set(kept.map(e => e.id));
      const add = nextHist.filter(e => !e.id).map(e => toRow(next.sn, e));
      // Undo: the newest movements that are no longer in the page's copy are removed.
      const remove = cur ? cur.history.filter(e => !keptIds.has(e.id) && e.at > newest).map(e => e.id) : [];
      if (!add.length && !remove.length && cur && (next.model || "") !== (cur.model || "")) add.push(toRow(next.sn, { at: Date.now(), model: next.model }, "TYPE"));
      if (add.length) await backend.append(add);
      if (remove.length) await backend.remove(remove);
    });
  }
  async function deleteBattery(id) {
    await write(async () => { const movs = await backend.movements(); await backend.remove(movs.filter(r => docId(r.sn) === id).map(r => r.id)); });
  }
  const db = {
    collection(name) {
      return {
        doc: id => ({ id, path: name + "/" + id, set: v => setBattery(id, v), delete: () => deleteBattery(id), update: v => setBattery(id, Object.assign({}, st.docs.get(id), v)) }),
        onSnapshot(next) { st.subs.push(next); if (st.loaded) setTimeout(() => next && fire(), 0); return () => { st.subs = st.subs.filter(f => f !== next); }; },
      };
    },
    doc(path) {
      return {
        id: path.split("/").pop(), path,
        set: v => write(async () => { await backend.saveSettings(v); await refreshSettings(); }),
        onSnapshot(next) { st.ssubs.push(next); if (st.settings) setTimeout(fireSettings, 0); return () => { st.ssubs = st.ssubs.filter(f => f !== next); }; },
      };
    },
  };
  const initials = n => (n || "?").split(/[\s.@]+/).filter(Boolean).slice(0, 2).map(x => x[0].toUpperCase()).join("");
  const avatar = n => "data:image/svg+xml;utf8," + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48"><rect width="48" height="48" rx="24" fill="#0f5bd8"/><text x="24" y="30" font-family="sans-serif" font-size="18" font-weight="700" fill="#fff" text-anchor="middle">${initials(n)}</text></svg>`);
  const user = {
    async me() { return { id: me.email, name: me.name, email: me.email, avatarUrl: avatar(me.name), color: "#0f5bd8", isOwner: false, canEdit: true }; },
    async id() { return me.email; }, async isOwner() { return false; }, async canEdit() { return true; }, async can() { return null; },
    async name() { return me.name; },
    async profiles(ids) { const o = {}; for (const id of [].concat(ids)) { const nm = id === me.email ? me.name : names.get(id) || ""; o[id] = { id, name: nm, avatarUrl: avatar(nm), color: "#0f5bd8", email: id, isMe: id === me.email, guest: false }; } return o; },
    async search() { return []; },
  };
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data], { type: "text/csv;charset=utf-8" });
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = filename; document.body.append(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 3000); return { status: "saved" };
    },
  };

  // ---------- sign-in ----------
  let readyResolve; const ready = new Promise(r => (readyResolve = r));
  window.claude = { use: async name => { await ready; return name === "db" ? db : name === "user" ? user : name === "downloads" ? downloads : null; } };

  function signInScreen(msg) {
    const d = document.createElement("div"); d.id = "xlSignin";
    d.style.cssText = "position:fixed;inset:0;z-index:80;display:flex;align-items:center;justify-content:center;padding:16px;background:var(--bg,#eef1f4)";
    d.innerHTML = '<div style="background:var(--surface,#fff);color:var(--ink,#14202c);border:1px solid var(--line,#d5dbe2);border-radius:14px;padding:28px;max-width:420px;width:100%;text-align:center;display:flex;flex-direction:column;gap:14px;font-family:system-ui,sans-serif">' +
      '<h1 style="margin:0;font-size:24px">Battery Stock</h1><p style="margin:0;color:var(--muted,#5b6876)">Sign in with your work Microsoft account to check batteries in and out.</p>' +
      '<button type="button" id="xlSignBtn" style="border:0;border-radius:9px;padding:13px 16px;font:600 16px system-ui,sans-serif;cursor:pointer;background:var(--accent,#0f5bd8);color:#fff">Sign in with Microsoft</button>' +
      '<div id="xlSignMsg" style="font-size:13px;color:var(--danger,#b42318)"></div></div>';
    document.body.append(d);
    d.querySelector("#xlSignMsg").textContent = msg || "";
    d.querySelector("#xlSignBtn").addEventListener("click", () => pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" }));
  }
  function addTopButtons() {
    const host = document.querySelector(".topbtns"); if (!host) return;
    const a = document.createElement("a"); a.id = "xlOpen"; a.className = "btn ghost"; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open Excel file"; a.hidden = true; a.style.textDecoration = "none";
    if (xl.webUrl) { a.href = xl.webUrl; a.hidden = false; }
    host.append(a);
    if (!DEMO) { const b = document.createElement("button"); b.type = "button"; b.className = "btn ghost"; b.textContent = "Sign out"; b.addEventListener("click", () => pca.logoutRedirect({ account })); host.append(b); }
    else { const s = document.createElement("span"); s.className = "pill out"; s.textContent = "Demo: data stays on this device"; s.title = "Fill in config.js to use the shared Excel file."; host.prepend(s); }
  }
  async function start() {
    if (!DEMO) {
      try {
        const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");
        pca = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + (CFG.tenantId || "organizations"), redirectUri }, cache: { cacheLocation: "localStorage" } });
        const r = await pca.handleRedirectPromise();
        account = (r && r.account) || pca.getAllAccounts()[0] || null;
      } catch (e) { signInScreen("Sign-in failed: " + (e.errorMessage || e.message || "unknown error")); return; }
      if (!account) { signInScreen(""); return; }
      pca.setActiveAccount(account);
      me = { name: account.name || account.username, email: (account.username || "").toLowerCase() };
    }
    addTopButtons();
    readyResolve();
    const first = async () => {
      try { await refreshSettings(); await refresh(); }
      catch (e) { problem(explain(e)); setTimeout(first, 20000); }
    };
    await first();
    setInterval(() => { if (!document.hidden && !busy && st.loaded) refresh().catch(e => problem(explain(e))); }, 20000);
    setInterval(() => { if (!document.hidden && !busy && st.loaded) refreshSettings().catch(() => {}); }, 120000);
    document.addEventListener("visibilitychange", () => { if (!document.hidden && !busy && st.loaded) refresh().catch(() => {}); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else start();

  // ---------- live camera scanner ----------
  // Outside Claude the browser may use the camera, so Scan opens a live picture that reads the code by itself.
  const clean = t => { const lines = String(t || "").replace(/[\u0000-\u001f]+/g, "\n").split("\n").map(x => x.trim()).filter(Boolean);
    const line = lines.find(x => /^(s\/?n|serial(\s*(no|nr|number))?)\s*[:#=]/i.test(x)) || lines[0] || "";
    return norm(line.replace(/^(s\/?n|serial(\s*(no|nr|number))?)\s*[:#=]\s*/i, "")); };
  let reader = null;
  function camModal() {
    let m = document.getElementById("liveCam"); if (m) return m;
    m = document.createElement("div"); m.id = "liveCam"; m.hidden = true;
    m.style.cssText = "position:fixed;inset:0;z-index:75;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(10,16,22,.6)";
    m.innerHTML = '<div style="background:var(--surface,#fff);color:var(--ink,#14202c);border-radius:12px;padding:16px;width:min(560px,100%);display:flex;flex-direction:column;gap:10px">' +
      '<strong style="font-size:17px">Scan the barcode</strong>' +
      '<video id="liveVideo" playsinline muted style="width:100%;max-height:60vh;background:#000;border-radius:8px"></video>' +
      '<div id="liveMsg" style="font-size:13px;color:var(--muted,#5b6876)">Starting camera…</div>' +
      '<div style="display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap">' +
      '<button type="button" class="btn ghost" id="livePhoto">Use a photo instead</button><button type="button" class="btn ghost" id="liveClose">Cancel</button></div></div>';
    document.body.append(m);
    m.querySelector("#liveClose").addEventListener("click", closeCam);
    m.querySelector("#livePhoto").addEventListener("click", () => { closeCam(); document.getElementById("camFile").click(); });
    return m;
  }
  function closeCam() { const m = document.getElementById("liveCam"); if (m) m.hidden = true; try { if (reader) reader.reset(); } catch {} reader = null; }
  const loadZX = () => window.ZXing ? Promise.resolve() : new Promise((res, rej) => { const s = document.createElement("script"); s.src = "lib/zxing.min.js"; s.onload = res; s.onerror = rej; document.head.append(s); });
  async function openCam() {
    const m = camModal(), msg = m.querySelector("#liveMsg"); m.hidden = false; msg.textContent = "Starting camera…";
    try {
      await loadZX();
      const hints = new Map(); hints.set(ZXing.DecodeHintType.TRY_HARDER, true);
      reader = new ZXing.BrowserMultiFormatReader(hints, 120);
      const types = new Set(((st.settings && st.settings.types) || []).map(t => norm(t.name)));
      let last = "", lastAt = 0;
      await reader.decodeFromConstraints({ audio: false, video: { facingMode: { ideal: "environment" }, width: { ideal: 1920 }, height: { ideal: 1080 } } }, m.querySelector("#liveVideo"), result => {
        if (!result || m.hidden) return;
        const code = clean(result.getText()); if (!code) return;
        if (code === last && Date.now() - lastAt < 2500) return; last = code; lastAt = Date.now();
        if (types.has(code)) {   // the type barcode: fill in the type and keep looking for the S/N
          const sel = document.getElementById("model"), opt = [...sel.options].find(o => norm(o.value) === code);
          if (opt) { sel.value = opt.value; sel.dispatchEvent(new Event("change", { bubbles: true })); }
          msg.textContent = `Type ${code} read. Now point the camera at the serial number barcode.`; return;
        }
        closeCam(); if (navigator.vibrate) navigator.vibrate(60);
        const sn = document.getElementById("sn"); sn.value = code; sn.focus();
        sn.dispatchEvent(new Event("input", { bubbles: true }));
        setTimeout(() => sn.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })), 50);
        const hint = document.getElementById("camHint"); if (hint) hint.textContent = `Barcode read: ${code}.`;
      });
      msg.textContent = "Hold the barcode steady in the picture. It is read by itself.";
    } catch (e) {
      msg.textContent = e && e.name === "NotAllowedError" ? "Camera access was blocked. Allow the camera for this site in the browser settings, or use a photo." : "No camera available here. Use a photo, type the S/N, or use a handheld barcode scanner.";
    }
  }
  document.addEventListener("click", e => {
    const b = e.target && e.target.closest && e.target.closest("#camBtn");
    if (!b || !(navigator.mediaDevices && navigator.mediaDevices.getUserMedia)) return;
    e.preventDefault(); e.stopImmediatePropagation(); openCam();
  }, true);
})();
