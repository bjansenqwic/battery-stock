/* Battery Stock – connection to the Excel file.
 *
 * The page itself is the same as the version inside Claude. This file gives it what it needs:
 *  - sign-in with the Microsoft work account (MSAL),
 *  - reading and writing the Excel workbook in OneDrive/SharePoint (Microsoft Graph),
 *  - a live camera scanner.
 *
 * The workbook has four tables:
 *  Movements – one row per check-in, check-out, charge or type change. Rows are only added, never edited,
 *              so several people can work at the same time. Current stock is worked out from them.
 *  Types     – battery types and the start of their serial numbers.
 *  Locations – storage spots, their maximum and what they are for ("Crate: <size>" marks a crate).
 *  CrateSizes – crate sizes and how many batteries each holds.
 *
 * Without clientId and workbookUrl in config.js the app runs in demo mode: data stays in this browser.
 */
(() => {
  "use strict";
  const CFG = Object.assign({ clientId: "", tenantId: "", workbookUrl: "" }, window.BATTERY_CONFIG || {});
  const DEMO = !CFG.clientId || !CFG.workbookUrl || /[?&]demo\b/.test(location.search);   // add ?demo to the address to try the app without signing in
  const TABLE = { mov: "Movements", types: "Types", locs: "Locations", crates: "CrateSizes" };
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

  let me = { name: "Demo user", email: "demo@example.com", isAdmin: true };   // demo: you may try the settings
  const names = new Map();

  // ---------- a message bar for connection problems ----------
  function problem(msg) {
    let bar = document.getElementById("xlProblem");
    if (!msg) { if (bar) bar.remove(); return; }
    if (!bar) {
      bar = document.createElement("div"); bar.id = "xlProblem"; bar.setAttribute("role", "alert");
      bar.style.cssText = "position:fixed;left:50%;transform:translateX(-50%);top:calc(10px + env(safe-area-inset-top,0px));z-index:70;max-width:calc(100% - 32px);background:var(--danger,#b42318);color:#fff;padding:10px 14px;border-radius:10px;font:500 14px Figtree,system-ui,sans-serif;box-shadow:0 4px 16px rgba(0,0,0,.2)";
      document.body.append(bar);
    }
    bar.textContent = msg;
  }
  const explain = e => e && e.userMessage ? e.userMessage
    : e && e.status === 404 ? "Couldn't find the Excel file or one of its tables (Movements, Types, Locations). Check the link in config.js."
    : e && e.status === 403 ? "You don't have access to the Excel file. Ask for edit access to the file."
    : e && e.status === 423 ? "The Excel file is locked. Close it in the desktop Excel app; Excel for the web is fine."
    : "Couldn't reach the Excel file. Check your internet connection." + (e && e.message ? " (" + e.message + ")" : "");

  // ---------- demo backend: this browser only ----------
  const demo = {
    read(k, d) { try { return JSON.parse(localStorage.getItem(k) || "null") || d; } catch { return d; } },
    async movements() { return this.read("bs.mov", []); },
    async append(rows) { const a = this.read("bs.mov", []); a.push(...rows); localStorage.setItem("bs.mov", JSON.stringify(a)); },
    async remove(ids) { const s = new Set(ids); localStorage.setItem("bs.mov", JSON.stringify(this.read("bs.mov", []).filter(r => !s.has(r.id)))); },
    async settings() { return this.read("bs.set", window.BATTERY_DEFAULTS || { types: [], locations: [] }); },
    async saveSettings(v) { localStorage.setItem("bs.set", JSON.stringify({ types: v.types, locations: v.locations, crateSizes: v.crateSizes || [] })); },
  };

  // ---------- Excel backend (Microsoft Graph) ----------
  const SCOPES = ["User.Read", "Files.ReadWrite.All"];
  let pca = null, account = null;
  // Installed app (Chrome "Install"): a full-page redirect to Microsoft can leave the app window and lose the sign-in state,
  // so there the sign-in opens in a small Microsoft window (popup) instead.
  const INSTALLED = matchMedia("(display-mode: standalone)").matches || matchMedia("(display-mode: window-controls-overlay)").matches || navigator.standalone === true;
  // Phones and tablets: a sign-in popup often never opens or never comes back there, so they always use the normal
  // redirect (it stays inside the installed app). Only an installed app on a computer uses the popup.
  const MOBILE = /Android|iPhone|iPad|iPod|Mobile/i.test(navigator.userAgent) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  const POPUP = INSTALLED && !MOBILE;
  // This page loaded inside that Microsoft sign-in window: do nothing, the app window reads the answer and closes it.
  const IN_POPUP = !!window.opener && window.opener !== window && /^msal\./.test(window.name || "");
  async function token() {
    try { return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) {
      if (POPUP) return (await pca.acquireTokenPopup({ scopes: SCOPES, account })).accessToken;
      await pca.acquireTokenRedirect({ scopes: SCOPES, account }); throw e;
    }
  }
  async function graph(path, opts = {}, tries = 5) {
    const url = path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path;
    for (let i = 0; ; i++) {
      let res;
      try { res = await fetch(url, { method: opts.method || "GET", headers: Object.assign({ Authorization: "Bearer " + await token(), "Content-Type": "application/json" }, opts.headers || {}), body: opts.body ? JSON.stringify(opts.body) : undefined }); }
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
      if (!this._p) this._p = (async () => {
        const b64 = btoa(unescape(encodeURIComponent(CFG.workbookUrl.trim()))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
        // redeemSharingLink: opening the QWIC-wide link gives this colleague lasting access to the file
        const item = await graph(`/shares/u!${b64}/driveItem?$select=id,name,webUrl,parentReference`, { headers: { Prefer: "redeemSharingLink" } });
        this.root = `/drives/${item.parentReference.driveId}/items/${item.id}/workbook`;
        this.webUrl = item.webUrl;
        const a = document.getElementById("xlOpen"); if (a && this.webUrl) { a.href = this.webUrl; a.hidden = false; }
        try { await this.ensureTables(); }
        catch (e) { if (!e.userMessage) e.userMessage = "The Excel file is missing the Movements, Types and Locations tables, and the app couldn't add them. Copy the sheets from the template Battery-Stock.xlsx into the file."; throw e; }
      })().catch(e => { this._p = null; throw e; });
      await this._p;
    },
    // An empty workbook gets the three tables it needs, with the battery types and locations filled in.
    async ensureTables() {
      const list = async path => (((await graph(path)) || {}).value || []).map(x => String(x.name).toLowerCase());
      const tables = new Set(await list(`${this.root}/tables?$select=name`));
      const sheets = new Set(await list(`${this.root}/worksheets?$select=name`));
      const d = window.BATTERY_DEFAULTS || { types: [], locations: [] };
      const plan = [
        [TABLE.mov, COLS.map(c => c[0]), []],
        [TABLE.types, ["Type", "S/N starts with"], d.types.map(t => [t.name, (t.codes || []).join(", ")])],
        [TABLE.locs, ["Location", "Max", "For"], d.locations.map(l => [l.name, l.cap || "", ONLY_TO_TEXT[l.only] || ""])],
        [TABLE.crates, ["Crate size", "Max"], (d.crateSizes || []).map(c => [c.name, c.cap || ""])],
      ];
      for (const [name, head, rows] of plan) {
        if (tables.has(name.toLowerCase())) continue;
        if (!sheets.has(name.toLowerCase())) await graph(`${this.root}/worksheets/add`, { method: "POST", body: { name } });
        const body = rows.length ? rows : [head.map(() => "")];
        const end = String.fromCharCode(64 + head.length) + (body.length + 1);
        const ws = `${this.root}/worksheets/${encodeURIComponent(name)}`;
        await graph(`${ws}/range(address='A1:${end}')`, { method: "PATCH", body: { values: [head, ...body.map(r => r.map(cell))] } });
        const t = await graph(`${ws}/tables/add`, { method: "POST", body: { address: `A1:${end}`, hasHeaders: true } });
        await graph(`${this.root}/tables/${encodeURIComponent(t.id)}`, { method: "PATCH", body: { name } });
      }
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
        out.push({ id: g("id") || "row" + i, at, mv: mv.startsWith("O") ? "OUT" : mv.startsWith("T") ? "TYPE" : mv.startsWith("C") ? "CHARGE" : "IN", sn, model: g("model"), condition: g("condition"),
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
      const locations = lv.slice(1).map(r => { const f = txt(r[2]), cr = f.match(/^crate\b\s*:?\s*(.*)$/i);
        return Object.assign({ name: txt(r[0]), cap: Math.max(0, parseInt(txt(r[1]), 10) || 0), only: cr ? "" : TEXT_TO_ONLY(f) }, cr ? { crate: cr[1].trim() || "Crate" } : {}); }).filter(l => l.name);
      const cv = await this.values(TABLE.crates);
      const crateSizes = cv.slice(1).map(r => ({ name: txt(r[0]), cap: Math.max(0, parseInt(txt(r[1]), 10) || 0) })).filter(c => c.name);
      return { types, locations, crateSizes };
    },
    async saveSettings(v) {
      await this.writeTable(TABLE.types, v.types.map(t => [t.name, (t.codes || []).join(", ")]));
      await this.writeTable(TABLE.locs, v.locations.map(l => [l.name, l.cap ? l.cap : "", l.crate ? "Crate: " + l.crate : ONLY_TO_TEXT[l.only] || ""]));
      await this.writeTable(TABLE.crates, (v.crateSizes || []).map(c => [c.name, c.cap || ""]));
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
      // a charge: the battery stays where it is; its 3 months start again
      if (r.mv === "CHARGE") { if (b) { b.history.push({ id: r.id, t: "charge", at: r.at, loc: r.loc, note: r.note, by: r.email || r.by || null }); b.lastCharge = r.at; } continue; }
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
    mv: mv || (e.t === "out" ? "OUT" : e.t === "charge" ? "CHARGE" : "IN"), sn, model: e.model || "", condition: e.condition || "", purpose: e.purpose || "",
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
        set: v => me.isAdmin ? write(async () => { await backend.saveSettings(v); await refreshSettings(); }) : Promise.reject(new Error("Only admins can change types and locations.")),
        onSnapshot(next) { st.ssubs.push(next); if (st.settings) setTimeout(fireSettings, 0); return () => { st.ssubs = st.ssubs.filter(f => f !== next); }; },
      };
    },
  };
  const initials = n => (n || "?").split(/[\s.@]+/).filter(Boolean).slice(0, 2).map(x => x[0].toUpperCase()).join("");
  const avatar = n => "data:image/svg+xml;utf8," + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="48" height="48"><rect width="48" height="48" rx="24" fill="#0f5bd8"/><text x="24" y="30" font-family="sans-serif" font-size="18" font-weight="700" fill="#fff" text-anchor="middle">${initials(n)}</text></svg>`);
  const user = {
    async me() { return { id: me.email, name: me.name, email: me.email, avatarUrl: avatar(me.name), color: "#0f5bd8", isOwner: !!me.isAdmin, canEdit: true }; },
    async id() { return me.email; }, async isOwner() { return !!me.isAdmin; }, async canEdit() { return true; }, async can() { return null; },
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
    d.innerHTML = '<div style="background:var(--surface,#fff);color:var(--ink,#14202c);border:1px solid var(--line,#d5dbe2);border-radius:8px;padding:28px;max-width:420px;width:100%;text-align:center;display:flex;flex-direction:column;gap:14px;font-family:Figtree,system-ui,sans-serif">' +
      '<h1 style="margin:0;font-size:24px">Battery Stock</h1><p style="margin:0;color:var(--muted,#5b6876)">Sign in with your work Microsoft account to check batteries in and out.</p>' +
      '<button type="button" id="xlSignBtn" style="border:0;border-radius:999px;padding:13px 16px;font:600 16px Figtree,system-ui,sans-serif;cursor:pointer;background:var(--accent,#0a0a0a);color:var(--accent-ink,#fff)">Sign in with Microsoft</button>' +
      '<div id="xlSignMsg" style="font-size:13px;color:var(--danger,#b42318)"></div><div style="font-size:11px;color:var(--muted,#5b6876)">App version 2026-11-27a</div></div>';
    document.body.append(d);
    d.querySelector("#xlSignMsg").textContent = msg || "";
    d.querySelector("#xlSignBtn").addEventListener("click", async () => {
      d.querySelector("#xlSignMsg").textContent = "";
      try {
        if (!POPUP) { await pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" }); return; }
        const r = await pca.loginPopup({ scopes: SCOPES, prompt: "select_account" }); pca.setActiveAccount(r.account); location.reload();
      } catch (e) {
        const code = (e && e.errorCode) || "";
        // a blocked popup: fall back to the normal redirect
        if (POPUP && /popup_window_error|empty_window|monitor_window_timeout/.test(code)) { try { await pca.loginRedirect({ scopes: SCOPES, prompt: "select_account" }); return; } catch {} }
        // an earlier sign-in that never finished blocks a new one: clear it and try again
        if (/interaction_in_progress/.test(code)) { for (const k of Object.keys(localStorage)) if (/interaction\.status/i.test(k)) localStorage.removeItem(k); for (const k of Object.keys(sessionStorage)) if (/interaction\.status/i.test(k)) sessionStorage.removeItem(k); d.querySelector("#xlSignMsg").textContent = "Please tap Sign in again."; return; }
        d.querySelector("#xlSignMsg").textContent = "Sign-in failed: " + ((e && (e.errorMessage || e.message)) || "unknown error") + (code ? " (" + code + ")" : "");
      }
    });
  }
  function addTopButtons() {
    const host = document.querySelector(".topbtns"); if (!host) return;
    const a = document.createElement("a"); a.id = "xlOpen"; a.className = "btn ghost"; a.target = "_blank"; a.rel = "noopener"; a.textContent = "Open Excel file"; a.hidden = true; a.style.textDecoration = "none";
    if (xl.webUrl) { a.href = xl.webUrl; a.hidden = false; }
    host.append(a);
    // Sign out sits in the account menu (the round icon top right); the page calls this
    if (!DEMO) window.bsSignOut = () => POPUP ? pca.logoutPopup({ account }).catch(() => {}).finally(() => location.reload()) : pca.logoutRedirect({ account });
    else { const s = document.createElement("span"); s.className = "pill out"; s.textContent = "Demo: data stays on this device"; s.title = "Fill in config.js to use the shared Excel file."; host.prepend(s); }
  }
  async function start() {
    if (IN_POPUP) { document.body.style.visibility = "hidden"; return; }
    if (!DEMO) {
      try {
        // A sign-in that never finished (for example a popup that never came back) blocks every new one; unless we are
        // just back from Microsoft, clear that leftover first.
        if (!/[#&?](code|error|state)=/.test(location.hash + location.search)) for (const st of [localStorage, sessionStorage]) for (const k of Object.keys(st)) if (/interaction\.status/i.test(k)) st.removeItem(k);
        const redirectUri = location.origin + location.pathname.replace(/index\.html$/, "");
        pca = new msal.PublicClientApplication({ auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + (CFG.tenantId || "organizations"), redirectUri }, cache: { cacheLocation: "localStorage", temporaryCacheLocation: "localStorage", storeAuthStateInCookie: true } });
        const r = await pca.handleRedirectPromise();
        account = (r && r.account) || pca.getAllAccounts()[0] || null;
      } catch (e) { signInScreen("Sign-in failed: " + (e.errorMessage || e.message || "unknown error")); return; }
      if (!account) { signInScreen(""); return; }
      pca.setActiveAccount(account);
      // Admins: users given the app role "Admin" (or config adminRole) in Entra ID > Enterprise applications > Users and groups
      me = { name: account.name || account.username, email: (account.username || "").toLowerCase(), isAdmin: (((account.idTokenClaims || {}).roles) || []).includes(CFG.adminRole || "Admin") };
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

  // ---------- reading the printed text (OCR) ----------
  // The letters and digits on the label are read first; barcodes are the backup. Tesseract (from the jsDelivr CDN)
  // loads on first use and is kept ready. bsOCR is also used by the page for "Use a photo".
  const OCR_CDN = "https://cdn.jsdelivr.net/npm/";
  let ocrP = null;
  function getOCR() {
    if (ocrP) return ocrP;
    ocrP = (async () => {
      if (!window.Tesseract) await new Promise((res, rej) => { const s = document.createElement("script"); s.src = OCR_CDN + "tesseract.js@5.1.1/dist/tesseract.min.js"; s.onload = res; s.onerror = rej; document.head.append(s); });
      const w = await window.Tesseract.createWorker("eng", 1, { workerPath: OCR_CDN + "tesseract.js@5.1.1/dist/worker.min.js", corePath: OCR_CDN + "tesseract.js-core@5.1.1", langPath: OCR_CDN + "@tesseract.js-data/eng@1.0.0/4.0.0_best_int" });
      return w;
    })().catch(e => { ocrP = null; throw e; });
    return ocrP;
  }
  const looksLikeSn = s => s.length >= 8 && s.length <= 32 && /[A-Z]/.test(s) && /\d{3}/.test(s);
  // how: "label" = printed after S/N or Serial, "family" = a known S/N start, "guess" = the likeliest code on the label
  function parseLabel(text) {
    const lines = String(text || "").toUpperCase().split(/\n+/).map(l => l.trim()).filter(Boolean);
    let sn = "", how = "", art = "";
    for (let i = 0; i < lines.length && !sn; i++) {
      const m = lines[i].match(/(?:S\/?N|SERIAL)(?:\s*(?:NO|NR)\.?)?\s*[:#.]?\s*([A-Z0-9][A-Z0-9 ]{5,})/);
      if (m) { const t = m[1].replace(/\s+/g, ""); if (looksLikeSn(t)) { sn = t; how = "label"; } }
      else if (/^(S\/?N|SERIAL)\b[^A-Z0-9]*$/.test(lines[i]) && lines[i + 1]) { const t = lines[i + 1].replace(/[^A-Z0-9]/g, ""); if (looksLikeSn(t)) { sn = t; how = "label"; } }
    }
    const tokens = lines.flatMap(l => l.split(/\s+/)).map(t => t.replace(/[^A-Z0-9]/g, "")).filter(Boolean);
    const am = lines.join(" ").match(/ART\s*(?:#|NR\.?|NO\.?)?\s*[:.]?\s*(BA0{2}\d{3})/); art = am ? am[1] : (tokens.find(t => /^BA0{2}\d{3}$/.test(t)) || "");
    const fam = t => { const m = t.match(/(BA\d{3}|WDHL|XWHL|BDLG|BPYT)/); return m && m.index <= 2 ? t.slice(m.index) : ""; };   // a known S/N start, also when a stray character was read in front of it
    if (sn && how === "label" && fam(sn)) sn = fam(sn);
    if (!sn) { const t = tokens.map(fam).find(t => t && looksLikeSn(t) && t !== art); if (t) { sn = t; how = "family"; } }
    if (!sn) { const t = tokens.filter(t => looksLikeSn(t) && t !== art).sort((a, b) => b.length - a.length)[0]; if (t) { sn = t; how = "guess"; } }
    return { sn: norm(sn), art: norm(art), how };
  }
  async function toCanvas(src, max, turn) {   // a picture, scaled down to max px on its long side, optionally turned a quarter
    const bmp = await createImageBitmap(src), k = Math.min(1, max / Math.max(bmp.width, bmp.height));
    const w = Math.round(bmp.width * k), h = Math.round(bmp.height * k), cv = document.createElement("canvas"), cx = cv.getContext("2d");
    if (turn) { cv.width = h; cv.height = w; cx.translate(turn > 0 ? h : 0, turn > 0 ? 0 : w); cx.rotate(turn * Math.PI / 2); } else { cv.width = w; cv.height = h; }
    cx.drawImage(bmp, 0, 0, w, h); return cv;
  }
  // read a picture as one block of text, then as scattered words; the first sure S/N wins
  async function readText(w, img, modes = ["6", "11"]) {
    let best = null;
    for (const psm of modes) {
      await w.setParameters({ tessedit_pageseg_mode: psm });
      const r = await w.recognize(img), p = parseLabel(r.data.text); p.conf = r.data.confidence || 0;
      if (p.sn && p.how !== "guess") return p;
      if (!best || (p.sn && !best.sn) || (p.art && !best.art)) best = Object.assign({}, best || {}, p);
    }
    return best;
  }
  window.bsOCR = async file => {   // { sn, art, unsure } from a photo of the label
    const w = await getOCR(); let guess = null, art = "";
    for (const turn of [0, 1, -1]) {   // as taken, then turned a quarter either way for upright labels
      const p = await readText(w, await toCanvas(file, 1800, turn));
      art = art || p.art;
      if (p.sn && p.how !== "guess") return { sn: p.sn, art: p.art || art, unsure: p.conf < 55 };
      if (p.sn && !guess) guess = p;
    }
    return guess ? { sn: guess.sn, art: guess.art || art, unsure: true } : { sn: "", art, unsure: false };
  };
  if (navigator.mediaDevices) setTimeout(() => getOCR().catch(() => {}), 5000);   // get the text reader ready before the first scan

  // ---------- live camera scanner ----------
  // Outside Claude the browser may use the camera, so Scan opens a live picture that reads the code by itself.
  // Fast path: the phone's own barcode reader (BarcodeDetector, built into Chrome on Android). Otherwise ZXing reads
  // the middle of the picture, scaled down; every third try is a thorough read that also finds upright codes.
  const clean = t => { const lines = String(t || "").replace(/[\u0000-\u001f]+/g, "\n").split("\n").map(x => x.trim()).filter(Boolean);
    const line = lines.find(x => /^(s\/?n|serial(\s*(no|nr|number))?)\s*[:#=]/i.test(x)) || lines[0] || "";
    return norm(line.replace(/^(s\/?n|serial(\s*(no|nr|number))?)\s*[:#=]\s*/i, "")); };
  let stream = null, scanning = false, detectorP = null;
  function camModal() {
    let m = document.getElementById("liveCam"); if (m) return m;
    m = document.createElement("div"); m.id = "liveCam"; m.hidden = true;
    m.style.cssText = "position:fixed;inset:0;z-index:75;display:flex;align-items:center;justify-content:center;padding:16px;background:rgba(10,16,22,.6)";
    m.innerHTML = '<div style="background:var(--surface,#fff);color:var(--ink,#14202c);border-radius:8px;padding:16px;width:min(560px,100%);display:flex;flex-direction:column;gap:10px">' +
      '<strong style="font-size:17px">Scan the barcode</strong>' +
      '<div style="position:relative;background:#000;border-radius:8px;overflow:hidden"><video id="liveVideo" playsinline muted style="display:block;width:100%;max-height:60vh;object-fit:cover"></video>' +
      '<div aria-hidden="true" style="position:absolute;left:10%;right:10%;top:25%;bottom:25%;border:2px solid #fa530b;border-radius:8px;box-shadow:0 0 0 9999px rgba(0,0,0,.25)"></div></div>' +
      '<div id="liveMsg" style="font-size:13px;color:var(--muted,#5b6876)">Starting camera…</div>' +
      '<div style="display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap">' +
      '<button type="button" class="btn ghost" id="liveTorch" hidden>Light on</button>' +
      '<button type="button" class="btn ghost" id="livePhoto">Use a photo instead</button><button type="button" class="btn ghost" id="liveClose">Cancel</button></div></div>';
    document.body.append(m);
    m.querySelector("#liveClose").addEventListener("click", closeCam);
    m.querySelector("#livePhoto").addEventListener("click", () => { closeCam(); document.getElementById("camFile").click(); });
    m.addEventListener("click", e => { if (e.target === m) closeCam(); });
    return m;
  }
  function closeCam() {
    const m = document.getElementById("liveCam"); if (m) m.hidden = true; scanning = false;
    if (stream) { for (const t of stream.getTracks()) t.stop(); stream = null; }
    const v = document.getElementById("liveVideo"); if (v) v.srcObject = null;
  }
  const loadZX = () => window.ZXing ? Promise.resolve() : new Promise((res, rej) => { const s = document.createElement("script"); s.src = "zxing.min.js"; s.onload = res; s.onerror = rej; document.head.append(s); });
  // Returns a function that reads codes from the video: the built-in reader when the phone has one, else ZXing.
  function makeDetector() {
    if (detectorP) return detectorP;
    detectorP = (async () => {
      if ("BarcodeDetector" in window) {
        try {
          const want = ["code_128","code_39","code_93","codabar","itf","ean_13","ean_8","upc_a","upc_e","qr_code","data_matrix","pdf417","aztec"];
          const have = await window.BarcodeDetector.getSupportedFormats(), formats = want.filter(f => have.includes(f));
          if (formats.length) { const d = new window.BarcodeDetector({ formats }); return { kind: "native", read: async v => (await d.detect(v)).map(x => x.rawValue) }; }
        } catch {}
      }
      await loadZX();
      const F = ZXing.BarcodeFormat, hints = new Map();
      hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [F.CODE_128, F.CODE_39, F.CODE_93, F.ITF, F.EAN_13, F.EAN_8, F.QR_CODE, F.DATA_MATRIX]);
      const fast = new ZXing.MultiFormatReader(); fast.setHints(hints);
      const hard = new ZXing.MultiFormatReader(); hard.setHints(new Map([...hints, [ZXing.DecodeHintType.TRY_HARDER, true]]));   // also tries the code turned a quarter
      const cv = document.createElement("canvas"), cx = cv.getContext("2d", { willReadFrequently: true });
      let n = 0;
      return { kind: "zxing", read: async v => {
        const vw = v.videoWidth, vh = v.videoHeight; if (!vw || !vh) return [];
        const sw = vw * .8, sh = vh * .5, sx = (vw - sw) / 2, sy = (vh - sh) / 2;   // the middle of the picture, where the frame is
        const k = Math.min(1, 1000 / sw); cv.width = Math.round(sw * k); cv.height = Math.round(sh * k);
        cx.drawImage(v, sx, sy, sw, sh, 0, 0, cv.width, cv.height);
        const r = ++n % 3 ? fast : hard;   // mostly the quick read, every third try the thorough one
        try { return [r.decodeWithState(new ZXing.BinaryBitmap(new ZXing.HybridBinarizer(new ZXing.HTMLCanvasElementLuminanceSource(cv)))).getText()]; }
        catch { return []; } finally { try { r.reset(); } catch {} }
      } };
    })().catch(e => { detectorP = null; throw e; });
    return detectorP;
  }
  async function openCam() {
    const m = camModal(), msg = m.querySelector("#liveMsg"), video = m.querySelector("#liveVideo"), torch = m.querySelector("#liveTorch");
    m.hidden = false; msg.textContent = "Starting camera…"; torch.hidden = true;
    try {
      const [s, det] = await Promise.all([
        navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: "environment" }, width: { ideal: 1280 }, height: { ideal: 720 } } }),
        makeDetector()]);
      if (m.hidden) { for (const t of s.getTracks()) t.stop(); return; }   // closed while starting
      stream = s; video.srcObject = s; await video.play();
      const track = s.getVideoTracks()[0], caps = track.getCapabilities ? track.getCapabilities() : {};
      if (caps.focusMode && caps.focusMode.includes("continuous")) track.applyConstraints({ advanced: [{ focusMode: "continuous" }] }).catch(() => {});
      if (caps.torch) { let on = false; torch.hidden = false; torch.textContent = "Light on";
        torch.onclick = () => { on = !on; track.applyConstraints({ advanced: [{ torch: on }] }).catch(() => {}); torch.textContent = on ? "Light off" : "Light on"; }; }
      msg.textContent = "Hold the label inside the orange frame. The printed S/N is read first; the barcode is the backup.";
      const types = new Set(((st.settings && st.settings.types) || []).map(t => norm(t.name)));
      const setType = code => { const sel = document.getElementById("model"), opt = [...sel.options].find(o => norm(o.value) === code); if (opt) { sel.value = opt.value; sel.dispatchEvent(new Event("change", { bubbles: true })); } };
      let last = "", lastAt = 0, done = false, ocrOk = true, ocrSn = "", bar = ""; const startAt = Date.now(); scanning = true;
      const finish = (code, how) => {   // fill in the S/N and go on with the form
        if (done) return; done = true; closeCam(); if (navigator.vibrate) navigator.vibrate(60);
        const sn = document.getElementById("sn"); sn.value = code; sn.focus();
        sn.dispatchEvent(new Event("input", { bubbles: true }));
        setTimeout(() => sn.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true })), 50);
        const hint = document.getElementById("camHint"); if (hint) hint.textContent = how === "text" ? `Read from the printed label: ${code}. Check it against the label.` : `Barcode read: ${code}.`;
      };
      // 1. the printed letters and digits, read from the middle of the picture
      (async () => {
        let w; try { w = await getOCR(); } catch { ocrOk = false; return; }
        const cv = document.createElement("canvas"), cx = cv.getContext("2d", { willReadFrequently: true }); let tries = 0;
        while (scanning && !m.hidden && !done) {
          const vw = video.videoWidth, vh = video.videoHeight;
          if (vw) {
            const sw = vw * .8, sh = vh * .5, k = Math.min(1, 1400 / sw); cv.width = Math.round(sw * k); cv.height = Math.round(sh * k);
            cx.drawImage(video, (vw - sw) / 2, (vh - sh) / 2, sw, sh, 0, 0, cv.width, cv.height);
            try { const p = await readText(w, cv, [++tries % 2 ? "6" : "11"]);   // block and scattered reading take turns
              if (p.art && types.has(p.art)) setType(p.art);
              if (p.sn && p.how !== "guess" && p.conf >= 50) { if (p.sn === ocrSn || p.how === "label" || p.sn === bar) { finish(p.sn, "text"); return; } ocrSn = p.sn; }   // a code read twice, or printed after S/N, counts
            } catch {}
          }
          await new Promise(r => setTimeout(r, 120));
        }
      })();
      // 2. the barcode: used when the text gives nothing within 3 seconds, or when it says the same
      const tick = async () => {
        if (!scanning || m.hidden || done) return;
        let codes = []; try { codes = await det.read(video); } catch {}
        for (const raw of codes) {
          const code = clean(raw); if (!code) continue;
          if (types.has(code)) { if (code !== last) { setType(code); msg.textContent = `Type ${code} read. Now hold the serial number in the frame.`; } last = code; lastAt = Date.now(); continue; }
          bar = code;
        }
        if (bar && (!ocrOk || bar === ocrSn || Date.now() - startAt > 3000)) { finish(bar, "barcode"); return; }
        setTimeout(tick, det.kind === "native" ? 60 : 30);
      };
      tick();
    } catch (e) {
      closeCam(); m.hidden = false;
      msg.textContent = e && e.name === "NotAllowedError" ? "Camera access was blocked. Allow the camera for this site in the browser settings, or use a photo." : "No camera available here. Use a photo, type the S/N, or use a handheld barcode scanner.";
    }
  }
  document.addEventListener("click", e => {
    const b = e.target && e.target.closest && e.target.closest("#camBtn");
    if (!b || !(navigator.mediaDevices && navigator.mediaDevices.getUserMedia)) return;
    e.preventDefault(); e.stopImmediatePropagation(); openCam();
  }, true);
  if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) setTimeout(() => makeDetector().catch(() => {}), 3000);   // get the reader ready before the first scan
})();
