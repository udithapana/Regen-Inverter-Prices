/**
 * REGEN INVERTER PRICE LIST — Google Apps Script backend
 * ---------------------------------------------------------------
 * Procurement updates inverter prices; every other team (Sales, Design,
 * Management, ...) sees the same list live. Every change is written to
 * the "History" tab.
 *
 * Setup (once):
 *   1. Create a new Google Sheet, e.g. "Regen Inverter Prices".
 *   2. Extensions > Apps Script, delete the sample code, paste this file, Save.
 *   3. Choose the function "setup" and press Run (allow the permissions).
 *      It creates the tabs, loads the 202 inverters from the SLD tool and
 *      writes the two team keys into the "Settings" tab.
 *   4. Deploy > New deployment > type "Web app"
 *        Execute as: Me      Who has access: Anyone
 *      Copy the web app link that ends with /exec into the app.
 *   After editing this code later: Deploy > Manage deployments > Edit (pencil)
 *   > Version: New version > Deploy, so the same /exec link keeps working.
 *
 * Team keys (in the "Settings" tab — change them there at any time):
 *   Procurement key : can edit prices, add brands/models, hide models
 *   View key        : everyone else, read only
 *
 * © Regen Renewables (Pvt) Ltd. Confidential — internal use only.
 */

// ===================== SETTINGS =====================
const CFG = {
  TZ: 'Asia/Colombo',
  RECENT_LIMIT: 300          // max rows returned for "Recent changes"
};
const SH_ITEMS = 'Inverters', SH_HIST = 'History', SH_SET = 'Settings';
const COLS = ['ID', 'Brand', 'Model', 'Category', 'Phase', 'kW', 'PriceLKR', 'Supplier', 'Stock', 'ValidUntil',
              'Remarks', 'Active', 'Source', 'UpdatedBy', 'UpdatedAt', 'CreatedBy', 'CreatedAt'];
const HIST_COLS = ['Time', 'ID', 'Brand', 'Model', 'Field', 'Old', 'New', 'By'];
const EDITABLE = ['Brand', 'Model', 'Category', 'Phase', 'kW', 'PriceLKR', 'Supplier', 'Stock', 'ValidUntil', 'Remarks', 'Active'];
const CATEGORIES = ['On-grid', 'Hybrid', 'BESS', 'Other'];
const STOCK = ['', 'In stock', 'Limited stock', 'Order basis', 'Out of stock'];
const TEAMS = ['Procurement', 'Sales', 'Design', 'Management', 'Finance', 'Operations', 'Other'];
const DEFAULT_NOTE = 'Unit price in LKR, excluding VAT. Confirm with Procurement before quoting large orders.';


// ===================== SETUP =====================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const items = sheet_(ss, SH_ITEMS, COLS);
  sheet_(ss, SH_HIST, HIST_COLS);
  const set = sheet_(ss, SH_SET, ['Setting', 'Value']);
  textFormat_(items);
  // settings rows (kept if they already exist)
  const have = settings_();
  const want = [['Procurement key', have['Procurement key'] || 'PROC-' + rand_()],
                ['View key', have['View key'] || 'VIEW-' + rand_()],
                ['Price note', have['Price note'] || DEFAULT_NOTE]];
  set.getRange(2, 1, want.length, 2).setValues(want);
  set.getRange(2, 2, want.length, 1).setNumberFormat('@');
  // initial list from the SLD tool (only when the sheet is empty)
  if (items.getLastRow() < 2) {
    const now = iso_(), rows = SEED.map((s, i) => [id_(i + 1), s[0], s[1], s[2], s[3], s[4], '', '', '', '', '', true, 'SLD tool', '', '', 'Setup', now]);
    writeRows_(items, 2, rows);
  }
  Logger.log('Setup done. Inverters: ' + (items.getLastRow() - 1) +
             '\nProcurement key: ' + want[0][1] + '\nView key: ' + want[1][1] +
             '\n(Change the keys in the "Settings" tab whenever you like.)');
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
// text columns stay text (stops Sheets turning "2026-12-31" or model names into dates/numbers)
function textFormat_(sh) {
  TEXT_COLS.forEach(c =>
    sh.getRange(2, COLS.indexOf(c) + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@'));
}
// writes rows at startRow, growing the sheet first and keeping text columns as text
const TEXT_COLS = ['ID', 'Brand', 'Model', 'Supplier', 'ValidUntil', 'Remarks', 'UpdatedAt', 'CreatedAt'];
function writeRows_(sh, startRow, rows, textAll) {
  if (!rows.length) return;
  const need = startRow + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (textAll) sh.getRange(startRow, 1, rows.length, rows[0].length).setNumberFormat('@');
  else TEXT_COLS.forEach(c => sh.getRange(startRow, COLS.indexOf(c) + 1, rows.length, 1).setNumberFormat('@'));
  sh.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}
function rand_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase(); }
function id_(n) { return 'INV-' + String(n).padStart(4, '0'); }
function iso_() { return new Date().toISOString(); }

// ===================== WEB APP =====================
function doGet() { return out_({ ok: true, app: 'Regen Inverter Prices' }); }

function doPost(e) {
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fn = ACTIONS[b.action];
    if (!fn) throw new Error('Unknown action');
    const user = auth_(b);
    return out_(Object.assign({ ok: true }, fn(b, user)));
  } catch (err) {
    return out_({ ok: false, error: String((err && err.message) || err) });
  }
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function settings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SET), o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
}

function auth_(b) {
  const name = String(b.name || '').trim().slice(0, 60);
  const team = TEAMS.indexOf(b.team) >= 0 ? b.team : 'Other';
  const key = String(b.key || '').trim().toUpperCase();
  if (!name) throw new Error('Enter your name');
  if (!key) throw new Error('Enter the team key');
  const s = settings_();
  let role = '';
  if (key === String(s['Procurement key'] || '').toUpperCase()) role = 'procurement';
  else if (key === String(s['View key'] || '').toUpperCase()) role = 'viewer';
  if (!role) throw new Error('Wrong team key');
  return { name: name, team: role === 'procurement' ? 'Procurement' : team, role: role, note: s['Price note'] || '' };
}
function requireProc_(u) { if (u.role !== 'procurement') throw new Error('Only Procurement can change prices'); }

const ACTIONS = {
  login: (b, u) => ({ user: { name: u.name, team: u.team, role: u.role } }),
  list: (b, u) => ({ items: readItems_().map(x => x.o), note: u.note, lists: { categories: CATEGORIES, stock: STOCK, teams: TEAMS }, time: iso_() }),
  save: (b, u) => { requireProc_(u); return withLock_(() => ({ item: saveOne_(b.item || {}, u) })); },
  bulk: (b, u) => { requireProc_(u); return withLock_(() => bulk_(b.changes || [], u)); },
  add: (b, u) => { requireProc_(u); return withLock_(() => ({ item: add_(b.item || {}, u, 'Manual') })); },
  sync: (b, u) => { requireProc_(u); return withLock_(() => sync_(b.models || [], u)); },
  history: (b) => ({ rows: history_(String(b.id || '')) }),
  recent: () => ({ rows: recent_() }),
  note: (b, u) => {
    requireProc_(u);
    const text = String(b.text || '').trim().slice(0, 300);
    return withLock_(() => {
      const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_SET);
      const v = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 2).getValues();
      let r = v.findIndex(x => String(x[0]).trim() === 'Price note');
      if (r < 0) { sh.appendRow(['Price note', text]); } else sh.getRange(r + 2, 2).setValue(text);
      logHist_([[iso_(), '', '', '', 'Price note', '', text, u.name]]);
      return { note: text };
    });
  }
};

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Server busy, please try again');
  try { return fn(); } finally { lock.releaseLock(); }
}

// ===================== DATA =====================
function itemsSheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_ITEMS); }

function readItems_() {
  const sh = itemsSheet_(), n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, COLS.length).getValues().map((r, i) => ({ row: i + 2, o: toObj_(r) })).filter(x => x.o.ID);
}
function toObj_(r) {
  const o = {};
  COLS.forEach((c, i) => {
    let v = r[i];
    if (v instanceof Date) v = c === 'ValidUntil' ? Utilities.formatDate(v, CFG.TZ, 'yyyy-MM-dd') : v.toISOString();
    o[c] = v;
  });
  o.Active = !(o.Active === false || String(o.Active).toUpperCase() === 'FALSE');
  o.PriceLKR = o.PriceLKR === '' || o.PriceLKR === null ? '' : Number(o.PriceLKR);
  o.Phase = Number(o.Phase) === 1 ? 1 : 3;
  o.kW = o.kW === '' ? '' : Number(o.kW);
  return o;
}
function toRow_(o) { return COLS.map(c => o[c] === undefined ? '' : o[c]); }
const keyOf_ = (brand, model) => String(brand).trim().toLowerCase() + '|' + String(model).trim().toLowerCase();

// cleans one field coming from the app
function clean_(f, v) {
  switch (f) {
    case 'Brand': case 'Model': return String(v || '').replace(/\s+/g, ' ').trim().slice(0, 80);
    case 'Category': return CATEGORIES.indexOf(v) >= 0 ? v : 'Other';
    case 'Phase': return Number(v) === 1 ? 1 : 3;
    case 'kW': { if (v === '' || v === null || v === undefined) return ''; const n = Number(v); if (!isFinite(n) || n < 0) throw new Error('kW must be a positive number'); return n; }
    case 'PriceLKR': {
      if (v === '' || v === null || v === undefined) return '';
      const n = Number(String(v).replace(/[, ]/g, ''));
      if (!isFinite(n) || n < 0) throw new Error('Price must be a positive number');
      return Math.round(n * 100) / 100;
    }
    case 'Stock': return STOCK.indexOf(v) >= 0 ? v : '';
    case 'ValidUntil': { const s = String(v || '').trim(); if (s && !/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('Valid until must be a date'); return s; }
    case 'Supplier': return String(v || '').trim().slice(0, 80);
    case 'Remarks': return String(v || '').trim().slice(0, 300);
    case 'Active': return !(v === false || String(v).toUpperCase() === 'FALSE');
  }
  return v;
}

// applies the changed fields of `item` to an existing row; returns the saved object
function apply_(cur, row, item, u, all, histOut) {
  const next = Object.assign({}, cur);
  EDITABLE.forEach(f => { if (Object.prototype.hasOwnProperty.call(item, f)) next[f] = clean_(f, item[f]); });
  if (!next.Brand || !next.Model) throw new Error('Brand and model are required');
  const k = keyOf_(next.Brand, next.Model);
  if (k !== keyOf_(cur.Brand, cur.Model) && all.some(x => x.o.ID !== cur.ID && keyOf_(x.o.Brand, x.o.Model) === k))
    throw new Error(next.Brand + ' ' + next.Model + ' is already in the list');
  const now = iso_(), hist = [];
  EDITABLE.forEach(f => { if (String(cur[f]) !== String(next[f])) hist.push([now, cur.ID, next.Brand, next.Model, f, String(cur[f]), String(next[f]), u.name]); });
  if (!hist.length) return { obj: cur, changed: false };
  next.UpdatedBy = u.name; next.UpdatedAt = now;
  itemsSheet_().getRange(row, 1, 1, COLS.length).setValues([toRow_(next)]);
  histOut.push.apply(histOut, hist);
  return { obj: next, changed: true };
}

function saveOne_(item, u) {
  const all = readItems_(), cur = all.find(x => x.o.ID === item.ID);
  if (!cur) throw new Error('This inverter was not found — reload the list');
  if (item.base !== undefined && String(item.base || '') !== String(cur.o.UpdatedAt || ''))
    throw new Error((cur.o.UpdatedBy || 'Someone') + ' changed this model just now. Reload and try again.');
  const hist = [], r = apply_(cur.o, cur.row, item, u, all, hist);
  logHist_(hist);
  return r.obj;
}

function bulk_(changes, u) {
  const all = readItems_(), byId = {}, hist = [], saved = [], skipped = [];
  all.forEach(x => byId[x.o.ID] = x);
  changes.slice(0, 500).forEach(ch => {
    const cur = byId[ch.ID];
    if (!cur) { skipped.push({ ID: ch.ID, why: 'not found' }); return; }
    if (ch.base !== undefined && String(ch.base || '') !== String(cur.o.UpdatedAt || '')) { skipped.push({ ID: ch.ID, model: cur.o.Model, why: 'changed by ' + (cur.o.UpdatedBy || 'someone') + ' meanwhile' }); return; }
    try {
      const r = apply_(cur.o, cur.row, ch, u, all, hist);
      if (r.changed) { cur.o = r.obj; saved.push(r.obj); }
    } catch (e) { skipped.push({ ID: ch.ID, model: cur.o.Model, why: e.message }); }
  });
  logHist_(hist);
  return { saved: saved, skipped: skipped };
}

function add_(item, u, source) {
  const all = readItems_();
  const o = { Brand: clean_('Brand', item.Brand), Model: clean_('Model', item.Model) };
  if (!o.Brand || !o.Model) throw new Error('Brand and model are required');
  if (all.some(x => keyOf_(x.o.Brand, x.o.Model) === keyOf_(o.Brand, o.Model))) throw new Error(o.Brand + ' ' + o.Model + ' is already in the list');
  EDITABLE.forEach(f => { if (f !== 'Brand' && f !== 'Model') o[f] = clean_(f, item[f] === undefined ? (f === 'Active' ? true : '') : item[f]); });
  const max = all.reduce((m, x) => Math.max(m, parseInt(String(x.o.ID).replace(/\D/g, ''), 10) || 0), 0);
  const now = iso_();
  Object.assign(o, { ID: id_(max + 1), Source: source, UpdatedBy: u.name, UpdatedAt: now, CreatedBy: u.name, CreatedAt: now });
  const sh = itemsSheet_();
  writeRows_(sh, sh.getLastRow() + 1, [toRow_(o)]);
  logHist_([[now, o.ID, o.Brand, o.Model, 'Added', '', o.PriceLKR === '' ? '' : String(o.PriceLKR), u.name]]);
  return toObj_(toRow_(o));
}

// adds models from the SLD tool that are not yet in the list (existing ones are never changed)
function sync_(models, u) {
  const all = readItems_(), have = {}, rows = [], hist = [], added = [];
  all.forEach(x => have[keyOf_(x.o.Brand, x.o.Model)] = 1);
  let max = all.reduce((m, x) => Math.max(m, parseInt(String(x.o.ID).replace(/\D/g, ''), 10) || 0), 0);
  const now = iso_();
  models.slice(0, 2000).forEach(m => {
    const brand = clean_('Brand', m.brand), model = clean_('Model', m.model);
    if (!brand || !model || have[keyOf_(brand, model)]) return;
    have[keyOf_(brand, model)] = 1;
    const o = { ID: id_(++max), Brand: brand, Model: model, Category: clean_('Category', m.category), Phase: clean_('Phase', m.phase),
                kW: m.kw === undefined ? '' : clean_('kW', m.kw), PriceLKR: '', Supplier: '', Stock: '', ValidUntil: '', Remarks: '',
                Active: true, Source: 'SLD tool', UpdatedBy: u.name, UpdatedAt: now, CreatedBy: u.name, CreatedAt: now };
    rows.push(toRow_(o)); added.push(toObj_(toRow_(o)));
    hist.push([now, o.ID, brand, model, 'Added', '', '', u.name + ' (SLD sync)']);
  });
  if (rows.length) { const sh = itemsSheet_(); writeRows_(sh, sh.getLastRow() + 1, rows); }
  logHist_(hist);
  return { added: added, checked: models.length };
}

// ===================== HISTORY =====================
function logHist_(rows) {
  if (!rows.length) return;
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_HIST);
  writeRows_(sh, sh.getLastRow() + 1, rows, true);
}
function histRows_(fromRow) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_HIST), last = sh.getLastRow();
  const start = Math.max(2, fromRow || 2);
  if (last < start) return [];
  return sh.getRange(start, 1, last - start + 1, HIST_COLS.length).getValues().map(r => {
    const o = {}; HIST_COLS.forEach((c, i) => o[c] = r[i] instanceof Date ? r[i].toISOString() : String(r[i])); return o;
  });
}
function history_(id) { return histRows_(2).filter(h => h.ID === id).reverse(); }
function recent_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_HIST);
  return histRows_(sh.getLastRow() - 2000).reverse().slice(0, CFG.RECENT_LIMIT);
}

// ===================== INITIAL LIST =====================
// From the SLD tool (udithapana/SLD-Genarator, inverter-data.js): [brand, model, category, phase, kW]
const SEED = [
 ["KACO (Ongrid)","blueplanet 125 NX3 M10","On-grid",3,125],
 ["KACO (Ongrid)","blueplanet 100 NX3 M8","On-grid",3,100],
 ["KACO (Ongrid)","blueplanet 60 NX3 M5","On-grid",3,60],
 ["KACO (Ongrid)","blueplanet 50 NX3 M5","On-grid",3,50],
 ["KACO (Ongrid)","blueplanet 20.0 NX3 M2","On-grid",3,20],
 ["KACO (Ongrid)","blueplanet 15.0 NX3 M2","On-grid",3,15],
 ["KACO (Ongrid)","blueplanet 10.0 NX3 M2","On-grid",3,10],
 ["Sunways (Ongrid)","STT-150KTL-SE","On-grid",3,150],
 ["Sunways (Ongrid)","STT-136KTL-SE","On-grid",3,136],
 ["Sunways (Ongrid)","STT-125KTL-SE","On-grid",3,125],
 ["Sunways (Ongrid)","STT-110KTL-SE","On-grid",3,110],
 ["Sunways (Ongrid)","STT-100KTL-SE","On-grid",3,100],
 ["Sunways (Ongrid)","STT-80KTL","On-grid",3,80],
 ["Sunways (Ongrid)","STT-60KTL-P","On-grid",3,60],
 ["Sunways (Ongrid)","STT-50KTL-P","On-grid",3,50],
 ["Sunways (Ongrid)","STT-60KTL-M","On-grid",3,60],
 ["Sunways (Ongrid)","STT-50KTL-M","On-grid",3,50],
 ["Sunways (Ongrid)","STT-45KTL","On-grid",3,45],
 ["Sunways (Ongrid)","STT-40KTL","On-grid",3,40],
 ["Sunways (Ongrid)","STT-36KTL","On-grid",3,36],
 ["Sunways (Ongrid)","STT-33KTL","On-grid",3,33],
 ["Sunways (Ongrid)","STT-30KTL","On-grid",3,30],
 ["Sunways (Ongrid)","STT-29.9KTL","On-grid",3,29.9],
 ["Sunways (Ongrid)","STT-25KTL-P","On-grid",3,25],
 ["Sunways (Ongrid)","STT-20KTL-P","On-grid",3,20],
 ["Sunways (Ongrid)","STT-17KTL-P","On-grid",3,17],
 ["Sunways (Ongrid)","STT-15KTL-P","On-grid",3,15],
 ["Sunways (Ongrid)","STT-12KTL-P","On-grid",3,12],
 ["Sunways (Ongrid)","STT-10KTL-P","On-grid",3,10],
 ["Huawei (Ongrid)","Sun2000-150K-MG0","On-grid",3,150],
 ["Huawei (Ongrid)","Sun2000-115KTL-M2","On-grid",3,115],
 ["Huawei (Ongrid)","Sun2000-100KTL-M2","On-grid",3,100],
 ["Huawei (Ongrid)","Sun2000-50KTL-M3","On-grid",3,50],
 ["Huawei (Ongrid)","Sun2000-40KTL-M3","On-grid",3,40],
 ["Huawei (Ongrid)","Sun2000-36KTL-M3","On-grid",3,36],
 ["Huawei (Ongrid)","Sun2000-30KTL-M3","On-grid",3,30],
 ["Huawei (Ongrid)","Sun2000-25KTL-M5","On-grid",3,25],
 ["Huawei (Ongrid)","Sun2000-20KTL-M5","On-grid",3,20],
 ["Huawei (Ongrid)","Sun2000-17KTL-M5","On-grid",3,17],
 ["Huawei (Ongrid)","Sun2000-15KTL-M5","On-grid",3,15],
 ["Huawei (Ongrid)","Sun2000-12KTL-M5","On-grid",3,12],
 ["Huawei (Ongrid)","Sun2000-10KTL-M1","On-grid",3,10],
 ["Sungrow (Ongrid)","Sungrow-SG150CX","On-grid",3,150],
 ["Sungrow (Ongrid)","Sungrow-SG125CX-P2","On-grid",3,125],
 ["Sungrow (Ongrid)","Sungrow-SG125CX-P3","On-grid",3,125],
 ["Sungrow (Ongrid)","Sungrow-SG110CX-P2","On-grid",3,110],
 ["Sungrow (Ongrid)","Sungrow-SG100CX-P2","On-grid",3,100],
 ["Sungrow (Ongrid)","Sungrow-SG5.0RT-P2","On-grid",3,5],
 ["Sungrow (Ongrid)","Sungrow-SG6.0RT-P2","On-grid",3,6],
 ["Sungrow (Ongrid)","Sungrow-SG7.0RT-P2","On-grid",3,7],
 ["Sungrow (Ongrid)","Sungrow-SG8.0RT-P2","On-grid",3,8],
 ["Sungrow (Ongrid)","Sungrow-SG10.0RT-P2","On-grid",3,10],
 ["Sungrow (Ongrid)","Sungrow-SG12.0RT-P2","On-grid",3,12],
 ["Sungrow (Ongrid)","Sungrow-SG15RT-P2","On-grid",3,15],
 ["Sungrow (Ongrid)","Sungrow-SG17RT-P2","On-grid",3,17],
 ["Sungrow (Ongrid)","Sungrow-SG20RT-P2","On-grid",3,20],
 ["Sungrow (Ongrid)","Sungrow-SG36CX-P2","On-grid",3,36],
 ["Sungrow (Ongrid)","Sungrow-SG40CX-P2","On-grid",3,40],
 ["Sungrow (Ongrid)","Sungrow-SG50CX-P2","On-grid",3,50],
 ["Sunways BESS","STA-HT-125","BESS",3,125],
 ["Sunways Hybrid","STH-80KTL-HT","Hybrid",3,80],
 ["Sunways Hybrid","STH-70KTL-HT","Hybrid",3,70],
 ["Sunways Hybrid","STH-60KTL-HT","Hybrid",3,60],
 ["Sunways Hybrid","STH-50KTL-HT","Hybrid",3,50],
 ["Sunways Hybrid","STH-40KTL-HT","Hybrid",3,40],
 ["Sunways Hybrid","STH-33KTL-HT","Hybrid",3,33],
 ["Sunways Hybrid","STH-30KTL-HT","Hybrid",3,30],
 ["Sunways Hybrid","STH-25KTL-HT","Hybrid",3,25],
 ["Sunways Hybrid","STH-20KTL-HT","Hybrid",3,20],
 ["Sunways Hybrid","STH-17KTL-HT","Hybrid",3,17],
 ["Sunways Hybrid","STH-15KTL-HT","Hybrid",3,15],
 ["Sunways Hybrid","STH-12KTL-HTP","Hybrid",3,12],
 ["Sunways Hybrid","STH-10KTL-HTP","Hybrid",3,10],
 ["Sunways Hybrid","STH-8KTL-HTP","Hybrid",3,8],
 ["Sunways Hybrid","STH-6KTL-HTP","Hybrid",3,6],
 ["Sunways Hybrid","STH-5KTL-HTP","Hybrid",3,5],
 ["Sunways Hybrid","STH-4KTL-HTP","Hybrid",3,4],
 ["Sunways Hybrid","STH-20KTL-LT","Hybrid",3,20],
 ["Sunways Hybrid","STH-15KTL-LT","Hybrid",3,15],
 ["Sunways Hybrid","STH-12KTL-LT","Hybrid",3,12],
 ["Sunways Hybrid","STH-10KTL-LT","Hybrid",3,10],
 ["Sunways Hybrid","STH-8KTL-LT","Hybrid",3,8],
 ["Sunways Hybrid","STH-6KTL-LT","Hybrid",3,6],
 ["Sunways Hybrid","STH-5KTL-LT","Hybrid",3,5],
 ["Sunways Hybrid","STH-3KTL-LSS","Hybrid",1,3],
 ["Sunways Hybrid","STH-3.6KTL-LSS","Hybrid",1,3.6],
 ["Sunways Hybrid","STH-4KTL-LS","Hybrid",1,4],
 ["Sunways Hybrid","STH-4.6KTL-LS","Hybrid",1,4.6],
 ["Sunways Hybrid","STH-5KTL-LS","Hybrid",1,5],
 ["Sunways Hybrid","STH-6KTL-LS","Hybrid",1,6],
 ["Sunways Hybrid","STH-7.5KTL-LS","Hybrid",1,7.5],
 ["Sunways Hybrid","STH-8KTL-LS","Hybrid",1,8],
 ["Sunways Hybrid","STH-3KTL-HSS","Hybrid",1,3],
 ["Sunways Hybrid","STH-3.6KTL-HSS","Hybrid",1,3.6],
 ["Sunways Hybrid","STH-4.2KTL-HSS","Hybrid",1,4.2],
 ["Sunways Hybrid","STH-4.6KTL-HSS","Hybrid",1,4.6],
 ["Sunways Hybrid","STH-5KTL-HSS","Hybrid",1,5],
 ["Sunways Hybrid","STH-6KTL-HSS","Hybrid",1,6],
 ["Sunways Hybrid","STH-7KTL-HSS","Hybrid",1,7],
 ["Sunways Hybrid","STH-8KTL-HSS","Hybrid",1,8],
 ["ATESS Hybrid","HPS-125KTL","Hybrid",3,125],
 ["ATESS Hybrid","HPS-100KTL","Hybrid",3,100],
 ["ATESS Hybrid","HPS-150HV","Hybrid",3,150],
 ["ATESS Hybrid","HPS-100HV","Hybrid",3,100],
 ["ATESS Hybrid","HPS-150","Hybrid",3,150],
 ["ATESS Hybrid","HPS-120","Hybrid",3,120],
 ["ATESS Hybrid","HPS-100","Hybrid",3,100],
 ["ATESS Hybrid","HPS-50","Hybrid",3,50],
 ["ATESS Hybrid","HPS-30","Hybrid",3,30],
 ["ATESS Hybrid","HPS-50KTL","Hybrid",3,50],
 ["ATESS Hybrid","HPS-40KTL","Hybrid",3,40],
 ["ATESS Hybrid","HPS-30KTL","Hybrid",3,30],
 ["ATESS Hybrid","HPS-20KTL","Hybrid",3,20],
 ["ATESS Hybrid","HPS-15KTL","Hybrid",3,15],
 ["Sungrow Hybrid","Sungrow-SH50CX","Hybrid",3,50],
 ["Sungrow Hybrid","Sungrow-SH80CX","Hybrid",3,80],
 ["Sungrow Hybrid","Sungrow-SH125CX","Hybrid",3,125],
 ["Suntech C&I Hybrid","STE-2ML-1MP","Hybrid",3,1125],
 ["Suntech C&I Hybrid","STE-1ML-500P","Hybrid",3,500],
 ["Suntech C&I Hybrid","STE-261L-125P","Hybrid",3,125],
 ["Suntech C&I Hybrid","STE-261L-75HT-300","Hybrid",3,75],
 ["Suntech C&I Hybrid","STE-261L-80HT-300","Hybrid",3,80],
 ["Suntech C&I Hybrid","STE-261L-99HT-300","Hybrid",3,99],
 ["Suntech C&I Hybrid","STE-261L-100HT-300","Hybrid",3,100],
 ["Suntech C&I Hybrid","STE-261L-125HT-300","Hybrid",3,125],
 ["Suntech C&I Hybrid","STE-25HT-150","Hybrid",3,25],
 ["Suntech C&I Hybrid","STE-29.9HT-150","Hybrid",3,29.9],
 ["Suntech C&I Hybrid","STE-30HT-150","Hybrid",3,30],
 ["Suntech C&I Hybrid","STE-40HT-150","Hybrid",3,40],
 ["Suntech C&I Hybrid","STE-50HT-150","Hybrid",3,50],
 ["Suntech C&I Hybrid","STE-75HT-300","Hybrid",3,75],
 ["Suntech C&I Hybrid","STE-80HT-300","Hybrid",3,80],
 ["Suntech C&I Hybrid","STE-99HT-300","Hybrid",3,99],
 ["Suntech C&I Hybrid","STE-100HT-300","Hybrid",3,100],
 ["Suntech C&I Hybrid","STE-125HT-300","Hybrid",3,125],
 ["Suntech Domestic Hybrid","STE-3HS-30","Hybrid",1,3],
 ["Suntech Domestic Hybrid","STE-3.6HS-30","Hybrid",1,3.6],
 ["Suntech Domestic Hybrid","STE-4.2HS-30","Hybrid",1,4.2],
 ["Suntech Domestic Hybrid","STE-4.6HS-30","Hybrid",1,4.6],
 ["Suntech Domestic Hybrid","STE-5HS-30","Hybrid",1,5],
 ["Suntech Domestic Hybrid","STE-6HS-30","Hybrid",1,6],
 ["Suntech Domestic Hybrid","STE-4HT-25","Hybrid",3,4],
 ["Suntech Domestic Hybrid","STE-5HT-25","Hybrid",3,5],
 ["Suntech Domestic Hybrid","STE-6HT-25","Hybrid",3,6],
 ["Suntech Domestic Hybrid","STE-8HT-25","Hybrid",3,8],
 ["Suntech Domestic Hybrid","STE-10HT-25","Hybrid",3,10],
 ["Suntech Domestic Hybrid","STE-12HT-25","Hybrid",3,12],
 ["Suntech Domestic Hybrid","STE-10HT-40","Hybrid",3,10],
 ["Suntech Domestic Hybrid","STE-12HT-40","Hybrid",3,12],
 ["Suntech Domestic Hybrid","STE-15HT-40","Hybrid",3,15],
 ["Suntech Domestic Hybrid","STE-20HT-40","Hybrid",3,20],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -3.0 SP","Hybrid",1,3],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -4.0 SP","Hybrid",1,4],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -5.0 SP","Hybrid",1,5],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -6.0 SP","Hybrid",1,6],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -8.0 SP","Hybrid",1,8],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -10.0 SP","Hybrid",1,10],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -12.0 SP","Hybrid",1,12],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 5.0 TP","Hybrid",3,5],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 6.0 TP","Hybrid",3,6],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 8.0 TP","Hybrid",3,8],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 10.0 TP","Hybrid",3,10],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC -12.0 TP","Hybrid",3,12],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 15.0 TP","Hybrid",3,15],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 17.0 TP","Hybrid",3,17],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 20.0 TP","Hybrid",3,20],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 25.0 TP","Hybrid",3,25],
 ["Sigenergy (ALL In One Hybrid)","Sigenstor EC - 30.0 TP","Hybrid",3,30],
 ["Sigenergy (Ongrid)","Sigen - 50M1","On-grid",3,50],
 ["Sigenergy (Ongrid)","Sigens - 60M1","On-grid",3,60],
 ["Sigenergy (Ongrid)","Sigen - 80M1","On-grid",3,80],
 ["Sigenergy (Ongrid)","Sigen - 100M1","On-grid",3,100],
 ["Sigenergy (Ongrid)","Sigen - 110M1","On-grid",3,110],
 ["Sigenergy (Ongrid)","Sigen - 125M1","On-grid",3,125],
 ["Sigenergy (Ongrid)","Sigen- 150M1","On-grid",3,150],
 ["Sigenergy (Ongrid)","Sigen - 166M1","On-grid",3,166],
 ["Sigenergy (Hybrid)","Sigenstor -3.0 SP2","Hybrid",1,3],
 ["Sigenergy (Hybrid)","Sigenstor -5.0 SP2","Hybrid",1,5],
 ["Sigenergy (Hybrid)","Sigenstor -6.0 SP2","Hybrid",1,6],
 ["Sigenergy (Hybrid)","Sigenstor -3.0 TP2","Hybrid",3,3],
 ["Sigenergy (Hybrid)","Sigenstor -4.0 TP2","Hybrid",3,4],
 ["Sigenergy (Hybrid)","Sigenstor -5.0 TP2","Hybrid",3,5],
 ["Sigenergy (Hybrid)","Sigenstor -6.0 TP2","Hybrid",3,6],
 ["Sigenergy (Hybrid)","Sigenstor -8.0 TP2","Hybrid",3,8],
 ["Sigenergy (Hybrid)","Sigenstor -10.0 TP2","Hybrid",3,10],
 ["Sigenergy (Hybrid)","Sigenstor 12.0 TP2","Hybrid",3,12],
 ["Sigenergy (Hybrid)","Sigenstor 12.0 TP3","Hybrid",3,12],
 ["Sigenergy (Hybrid)","Sigenstor 12.0 TP4","Hybrid",3,12],
 ["Sigenergy (Hybrid)","Sigenstor 12.0 TP5","Hybrid",3,12],
 ["Sigenergy (Hybrid)","Sigenstor 12.0 TP6","Hybrid",3,12],
 ["Sigenergy (Hybrid)","Sigen 50M1-HYA","Hybrid",3,50],
 ["Sigenergy (Hybrid)","Sigen 60M1-HYA","Hybrid",3,60],
 ["Sigenergy (Hybrid)","Sigen 80M1-HYA","Hybrid",3,80],
 ["Sigenergy (Hybrid)","Sigen 100M1-HYA","Hybrid",3,100],
 ["Sigenergy (Hybrid)","Sigen 110M1-HYA","Hybrid",3,110],
 ["Sigenergy (Hybrid)","Sigen 125M1-HYA","Hybrid",3,125],
 ["Sigenergy (Hybrid)","Sigen 50M1-HYB","Hybrid",3,50],
 ["Sigenergy (Hybrid)","Sigen 60M1-HYB","Hybrid",3,60],
 ["Sigenergy (Hybrid)","Sigen 80M1-HYB","Hybrid",3,80],
 ["Sigenergy (Hybrid)","Sigen 100M1-HYB","Hybrid",3,100],
 ["Sigenergy (Hybrid)","Sigen 110M1-HYB","Hybrid",3,110],
 ["Sigenergy (Hybrid)","Sigen 125M1-HYB","Hybrid",3,125]
];
