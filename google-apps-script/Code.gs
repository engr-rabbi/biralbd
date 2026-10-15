/**
 * ============================================================================
 *  BiralBond API  —  Google Sheets as the database (Google Apps Script)
 * ============================================================================
 *  This script is the "backend" of the BiralBond website. The website (on
 *  GitHub Pages) talks to this script; the script reads and writes the tabs
 *  of the Google Sheet it lives in.
 *
 *  Setup (details in DEPLOY.md):
 *   1. Paste this whole file into  Extensions > Apps Script  of your Sheet.
 *   2. Project Settings > Script properties: add ADMIN_EMAIL, ADMIN_PASSWORD
 *      (and optionally GEMINI_API_KEY for the Miau AI chat).
 *   3. Run the function  setup  once (creates the tabs and the starter data).
 *   4. Deploy > New deployment > Web app  (Execute as: Me, Access: Anyone).
 *
 *  Helper functions you can run from the editor:  setup, checkSetup, testAi
 *
 *  SECURITY MODEL
 *   - The Sheet itself stays private. Visitors only reach it through this
 *     script, and only for the actions listed in handle() below.
 *   - Visitors can read public content and submit forms. Anything that edits
 *     or deletes data needs an admin token (given after a successful login).
 *   - The admin password and the Gemini key live in Script Properties, which
 *     only you (the script owner) can see. Never put them in this file.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------
var CACHE_SECONDS = 21600;            // server-side cache for sheet reads (max 6h)
var TOKEN_HOURS = 12;                 // how long an admin login stays valid
var MAX_FIELD = 5000;                 // max characters per visitor-submitted field
var UPLOAD_MAX_BYTES = 6 * 1024 * 1024;
var UPLOAD_FOLDER_NAME = 'BiralBond Uploads';

// Tabs visitors may read (Member, Contact, Newsletter, Comment are private).
var PUBLIC_TABLES = ['Article', 'Breed', 'CareTip', 'CatFoodShop', 'CatShop', 'CommunityPost',
  'Event', 'FeatureToggle', 'Gallery', 'LostFound', 'ServiceProvider', 'SiteContent',
  'Testimonial', 'VeterinaryClinic'];

// Forms visitors may submit: which fields they may set, and which are required.
var SUBMIT_RULES = {
  Newsletter:    { fields: ['email'], required: ['email'] },
  Member:        { fields: ['name', 'email', 'phone', 'city', 'plan', 'cats'], required: ['name', 'email'] },
  Contact:       { fields: ['name', 'email', 'subject', 'message'], required: ['name', 'email', 'message'] },
  CommunityPost: { fields: ['author', 'avatar', 'title', 'body', 'category', 'city', 'image'], required: ['author', 'title', 'body'] },
  LostFound:     { fields: ['type', 'catName', 'breed', 'color', 'location', 'city', 'date', 'contact', 'phone', 'description', 'reward', 'image'],
                   required: ['type', 'location', 'city', 'date', 'contact', 'phone', 'description'] },
  Comment:       { fields: ['postId', 'author', 'avatar', 'body'], required: ['postId', 'author', 'body'] }
};
// Values filled in when a visitor leaves an optional-looking field empty.
var SUBMIT_DEFAULTS = {
  Contact: { subject: 'General enquiry' },
  Member: { city: 'Dhaka' }
};
// Max submissions per minute (all visitors together) - a basic spam brake.
var SUBMIT_LIMIT_PER_MIN = 30;
var AI_LIMIT_PER_MIN = 20;

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------
function doGet(e) {
  var params = (e && e.parameter) || {};
  if (!params.action) params = { action: 'ping' };
  // Only harmless read actions work through a plain link. Everything else
  // (login, forms, admin) must be a POST from the website.
  if (['ping', 'bootstrap', 'comments'].indexOf(params.action) < 0) {
    return respond_({ ok: false, error: 'Use POST', code: 'BAD_REQUEST' });
  }
  return respond_(run_(params));
}

function doPost(e) {
  var params = {};
  try {
    params = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return respond_({ ok: false, error: 'Bad request', code: 'BAD_REQUEST' });
  }
  return respond_(run_(params));
}

function respond_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function run_(p) {
  try {
    return { ok: true, data: handle_(p || {}) };
  } catch (err) {
    var code = (err && err.code) || 'ERROR';
    var msg = (err && err.message) || String(err);
    if (code === 'ERROR') console.error(msg + (err && err.stack ? '\n' + err.stack : ''));
    return { ok: false, error: msg, code: code };
  }
}

function fail_(message, code) {
  var e = new Error(message);
  e.code = code || 'ERROR';
  throw e;
}

// ---------------------------------------------------------------------------
// Router: every request from the website is one of these actions
// ---------------------------------------------------------------------------
function handle_(p) {
  switch (p.action) {
    // ---- public ----
    case 'ping':       return ping_();
    case 'bootstrap':  return bootstrap_();
    case 'comments':   return listComments_(p.postId);
    case 'submit':     return submit_(p.table, p.data || {});
    case 'chat':       return aiChat_(p.messages);
    case 'names':      return aiNames_(p.vibe, p.gender, p.breed);
    case 'login':      return login_(p.email, p.password);

    // ---- admin (need a valid token) ----
    case 'session':    requireAdmin_(p.token); return { ok: true };
    case 'adminList':  requireAdmin_(p.token); return readTable_(requireTable_(p.table, true));
    case 'adminStats': requireAdmin_(p.token); return adminStats_();
    case 'adminCreate': requireAdmin_(p.token); return adminCreate_(p.table, p.data || {});
    case 'adminUpdate': requireAdmin_(p.token); return adminUpdate_(p.table, p.id, p.data || {});
    case 'adminDelete': requireAdmin_(p.token); return adminDelete_(p.table, p.id);
    case 'siteContentSave': requireAdmin_(p.token); return siteContentSave_(p.section, p.data || {});
    case 'upload':     requireAdmin_(p.token); return upload_(p.name, p.mime, p.base64);
    default:
      fail_('Unknown action', 'BAD_ACTION');
  }
}

function ping_() {
  var props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    time: new Date().toISOString(),
    sheetReady: !!props.getProperty('SHEET_ID'),
    adminConfigured: !!(props.getProperty('ADMIN_EMAIL') && props.getProperty('ADMIN_PASSWORD')),
    aiConfigured: !!props.getProperty('GEMINI_API_KEY')
  };
}

// ---------------------------------------------------------------------------
// Sheet access + cache
// ---------------------------------------------------------------------------
function ss_() {
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) fail_('Sheet is not set up yet. Run the setup function (see DEPLOY.md).', 'NOT_SETUP');
  return active;
}

function requireTable_(name, allowPrivate) {
  if (!SCHEMA[name]) fail_('Unknown table', 'BAD_TABLE');
  if (!allowPrivate && PUBLIC_TABLES.indexOf(name) < 0) fail_('Not allowed', 'FORBIDDEN');
  return name;
}

function pack_(obj) {
  var blob = Utilities.newBlob(JSON.stringify(obj), 'application/octet-stream', 'x');
  return Utilities.base64Encode(Utilities.gzip(blob).getBytes());
}

function unpack_(str) {
  var blob = Utilities.newBlob(Utilities.base64Decode(str), 'application/x-gzip', 'x.gz');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString());
}

function cacheKey_(table) { return 'tbl_' + table; }

function cacheGetMany_(tables) {
  var out = {};
  try {
    var keys = tables.map(cacheKey_);
    var hit = CacheService.getScriptCache().getAll(keys);
    tables.forEach(function (t) {
      var raw = hit[cacheKey_(t)];
      if (raw) { try { out[t] = unpack_(raw); } catch (e) { /* ignore bad entry */ } }
    });
  } catch (e) { /* cache is optional */ }
  return out;
}

function cachePut_(table, rows) {
  try {
    var s = pack_(rows);
    if (s.length < 95000) CacheService.getScriptCache().put(cacheKey_(table), s, CACHE_SECONDS);
  } catch (e) { /* too big or cache unavailable: just skip */ }
}

function cacheDrop_(table) {
  try { CacheService.getScriptCache().remove(cacheKey_(table)); } catch (e) { /* ignore */ }
}

// Convert one sheet cell into the type the website expects.
function fromCell_(v, type) {
  if (type === 'b') return v === true || String(v).toLowerCase() === 'true' || v === 1;
  if (type === 'i' || type === 'n') {
    if (v === '' || v === null || v === undefined) return 0;
    var n = Number(v);
    if (isNaN(n)) return 0;
    return type === 'i' ? Math.round(n) : n;
  }
  if (v instanceof Date) return v.toISOString();
  if (v === '' || v === null || v === undefined) return null;
  return String(v);
}

// Convert a value into what we write into a cell.
function toCell_(v, type) {
  if (type === 'b') return v === true || String(v).toLowerCase() === 'true';
  if (type === 'i' || type === 'n') {
    var n = Number(v);
    return isNaN(n) ? 0 : (type === 'i' ? Math.round(n) : n);
  }
  if (v === null || v === undefined) return '';
  var s = String(v);
  // never let user text turn into a spreadsheet formula
  if (s.charAt(0) === '=') s = ' ' + s;
  return s;
}

function sheetFor_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) fail_('Tab "' + name + '" is missing. Run the setup function again.', 'NOT_SETUP');
  return sh;
}

// Read a whole tab as an array of objects (cached).
function readTable_(name, skipCache) {
  if (!skipCache) {
    var cached = cacheGetMany_([name])[name];
    if (cached) return cached;
  }
  var rows = readSheetRows_(name);
  cachePut_(name, rows);
  return rows;
}

function readSheetRows_(name) {
  var cols = SCHEMA[name];
  var sh = sheetFor_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var values = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var rows = [];
  var needIds = false;
  for (var r = 0; r < values.length; r++) {
    var line = values[r];
    var blank = true;
    for (var c = 0; c < line.length; c++) { if (line[c] !== '' && line[c] !== null) { blank = false; break; } }
    if (blank) continue;
    var obj = {};
    for (var k = 0; k < cols.length; k++) obj[cols[k].n] = fromCell_(line[k], cols[k].t);
    if (!obj.id) needIds = true;
    obj.__row = r + 2;
    rows.push(obj);
  }
  if (needIds) {
    // Someone typed a row by hand in the Sheet: give it an id and timestamps.
    withLock_(function () { fillMissing_(name, sh, cols, values); });
    return readSheetRows_(name);
  }
  rows.forEach(function (o) { delete o.__row; });
  return rows;
}

function fillMissing_(name, sh, cols, values) {
  var idx = {}; cols.forEach(function (c, i) { idx[c.n] = i; });
  var now = new Date().toISOString();
  var changed = false;
  for (var r = 0; r < values.length; r++) {
    var line = values[r];
    var blank = line.every(function (v) { return v === '' || v === null; });
    if (blank) continue;
    ['id', 'createdAt', 'updatedAt', 'joinedAt', 'lastUpdated'].forEach(function (cn) {
      if (idx[cn] !== undefined && (line[idx[cn]] === '' || line[idx[cn]] === null)) {
        line[idx[cn]] = cn === 'id' ? Utilities.getUuid() : now;
        changed = true;
      }
    });
  }
  if (changed) sh.getRange(2, 1, values.length, cols.length).setValues(values);
}

var lockHeld_ = false;
function withLock_(fn) {
  if (lockHeld_) return fn();            // already inside a locked section
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  lockHeld_ = true;
  try { return fn(); } finally { lockHeld_ = false; lock.releaseLock(); }
}

function rowFromObject_(name, obj) {
  return SCHEMA[name].map(function (c) { return toCell_(obj[c.n], c.t); });
}

function findRowIndex_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

// Build a complete new record: defaults first, then the allowed values on top.
function buildRecord_(name, input, allowedFields, enforceRequired) {
  var now = new Date().toISOString();
  var rec = {};
  SCHEMA[name].forEach(function (c) {
    var d = c.d;
    if (d && d.uuid) rec[c.n] = Utilities.getUuid();
    else if (d && d.now) rec[c.n] = now;
    else if (d !== null && d !== undefined) rec[c.n] = d;
    else rec[c.n] = c.t === 'b' ? false : (c.t === 'i' || c.t === 'n') ? 0 : null;
    var allowed = !allowedFields || allowedFields.indexOf(c.n) >= 0;
    if (allowed && input[c.n] !== undefined && input[c.n] !== null && input[c.n] !== '') {
      var v = input[c.n];
      if (c.t === 's' || c.t === 't') v = String(v).slice(0, MAX_FIELD);
      rec[c.n] = v;
    }
  });
  if (enforceRequired) {
    SCHEMA[name].forEach(function (c) {
      if (c.req && (rec[c.n] === null || rec[c.n] === '' || rec[c.n] === undefined)) {
        fail_('"' + c.n + '" is required', 'VALIDATION');
      }
    });
  }
  return rec;
}

function appendRecord_(name, rec) {
  var sh = sheetFor_(name);
  var row = sh.getLastRow() + 1;
  if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 200);
  sh.getRange(row, 1, 1, SCHEMA[name].length).setValues([rowFromObject_(name, rec)]);
  cacheDrop_(name);
}

function cleanRecord_(rec) {
  var out = {};
  Object.keys(rec).forEach(function (k) { if (k !== '__row') out[k] = rec[k]; });
  return out;
}

// ---------------------------------------------------------------------------
// Public actions
// ---------------------------------------------------------------------------
function bootstrap_() {
  var hit = cacheGetMany_(PUBLIC_TABLES);
  var tables = {};
  PUBLIC_TABLES.forEach(function (t) {
    if (hit[t]) { tables[t] = hit[t]; return; }
    tables[t] = readSheetRows_(t);
    cachePut_(t, tables[t]);
  });
  var members = readTable_('Member');
  return { tables: tables, counts: { members: members.length } };
}

function listComments_(postId) {
  if (!postId) return [];
  return readTable_('Comment')
    .filter(function (c) { return String(c.postId) === String(postId); })
    .sort(function (a, b) { return String(a.createdAt) < String(b.createdAt) ? -1 : 1; });
}

function throttle_(key, limit) {
  try {
    var cache = CacheService.getScriptCache();
    var k = 'rl_' + key + '_' + Math.floor(Date.now() / 60000);
    var n = Number(cache.get(k) || 0) + 1;
    cache.put(k, String(n), 120);
    if (n > limit) fail_('Too many requests. Please wait a minute and try again.', 'RATE_LIMIT');
  } catch (e) {
    if (e && e.code === 'RATE_LIMIT') throw e;
  }
}

function submit_(table, data) {
  var rule = SUBMIT_RULES[table];
  if (!rule) fail_('Not allowed', 'FORBIDDEN');
  if (data.hp) return { ok: true };                 // honeypot: bots fill hidden fields
  throttle_('submit', SUBMIT_LIMIT_PER_MIN);

  var defaults = SUBMIT_DEFAULTS[table] || {};
  Object.keys(defaults).forEach(function (k) {
    if (data[k] === undefined || data[k] === null || String(data[k]).trim() === '') data[k] = defaults[k];
  });
  rule.required.forEach(function (f) {
    if (data[f] === undefined || data[f] === null || String(data[f]).trim() === '') fail_('Missing fields', 'VALIDATION');
  });
  if (data.email !== undefined) {
    data.email = String(data.email).trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) fail_('Invalid email', 'VALIDATION');
  }
  if (data.image && !/^(https?:\/\/|\/)/i.test(String(data.image))) data.image = null;

  return withLock_(function () {
    if (table === 'Newsletter') {
      var exists = readSheetRows_('Newsletter').some(function (r) { return r.email === data.email; });
      if (exists) return { ok: true, already: true };
    }
    if (table === 'Comment') {
      var posts = readSheetRows_('CommunityPost');
      var post = posts.filter(function (p) { return String(p.id) === String(data.postId); })[0];
      if (!post) fail_('Post not found', 'VALIDATION');
      if (!data.avatar) data.avatar = String(data.author).slice(0, 2).toUpperCase();
    }
    var rec = buildRecord_(table, data, rule.fields, true);
    appendRecord_(table, rec);
    if (table === 'Comment') bumpCommentCount_(data.postId);
    return (table === 'Newsletter') ? { ok: true } : rec;
  });
}

function bumpCommentCount_(postId) {
  var sh = sheetFor_('CommunityPost');
  var row = findRowIndex_(sh, postId);
  if (row < 0) return;
  var cols = SCHEMA.CommunityPost;
  var ci = cols.map(function (c) { return c.n; }).indexOf('comments');
  var cell = sh.getRange(row, ci + 1);
  cell.setValue((Number(cell.getValue()) || 0) + 1);
  cacheDrop_('CommunityPost');
}

// ---------------------------------------------------------------------------
// Admin login (stateless signed token)
// ---------------------------------------------------------------------------
function safeEqual_(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  var n = Math.max(a.length, b.length);
  for (var i = 0; i < n; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function secret_() {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('TOKEN_SECRET');
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('TOKEN_SECRET', s);
  }
  return s;
}

function sign_(payload) {
  var sig = Utilities.computeHmacSha256Signature(payload, secret_());
  return Utilities.base64EncodeWebSafe(sig);
}

function login_(email, password) {
  var props = PropertiesService.getScriptProperties();
  var adminEmail = (props.getProperty('ADMIN_EMAIL') || '').trim().toLowerCase();
  var adminPass = props.getProperty('ADMIN_PASSWORD') || '';
  if (!adminEmail || !adminPass) {
    fail_('ADMIN_EMAIL / ADMIN_PASSWORD সেট করা নেই (DEPLOY.md-এর ধাপ দেখুন)', 'NOT_CONFIGURED');
  }
  if (!email || !password) fail_('ইমেইল ও পাসওয়ার্ড প্রয়োজন', 'VALIDATION');

  var cache = CacheService.getScriptCache();
  var failKey = 'loginfail';
  var fails = Number(cache.get(failKey) || 0);
  if (fails >= 8) fail_('অনেকবার ভুল চেষ্টা হয়েছে। ১৫ মিনিট পরে আবার চেষ্টা করুন।', 'RATE_LIMIT');

  var okEmail = safeEqual_(String(email).trim().toLowerCase(), adminEmail);
  var okPass = safeEqual_(password, adminPass);
  if (!(okEmail && okPass)) {
    cache.put(failKey, String(fails + 1), 900);
    fail_('ভুল ইমেইল বা পাসওয়ার্ড', 'BAD_LOGIN');
  }
  cache.remove(failKey);

  var exp = Date.now() + TOKEN_HOURS * 3600 * 1000;
  var payload = Utilities.base64EncodeWebSafe(JSON.stringify({ e: adminEmail, x: exp }));
  return {
    token: payload + '.' + sign_(payload),
    exp: exp,
    user: { email: adminEmail, name: 'Admin', role: 'admin' }
  };
}

function requireAdmin_(token) {
  var parts = String(token || '').split('.');
  if (parts.length !== 2) fail_('Login required', 'AUTH');
  if (!safeEqual_(parts[1], sign_(parts[0]))) fail_('Login required', 'AUTH');
  var data;
  try {
    data = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  } catch (e) { fail_('Login required', 'AUTH'); }
  if (!data || !data.x || Date.now() > data.x) fail_('Session expired. Please log in again.', 'AUTH');
  return data;
}

// ---------------------------------------------------------------------------
// Admin actions
// ---------------------------------------------------------------------------
function adminCreate_(table, data) {
  requireTable_(table, true);
  return withLock_(function () {
    var clean = {};
    Object.keys(data).forEach(function (k) {
      if (['id', 'createdAt', 'updatedAt'].indexOf(k) < 0) clean[k] = data[k];
    });
    var rec = buildRecord_(table, clean, null, true);
    appendRecord_(table, rec);
    return rec;
  });
}

function adminUpdate_(table, id, data) {
  requireTable_(table, true);
  if (!id) fail_('id required', 'VALIDATION');
  return withLock_(function () {
    var sh = sheetFor_(table);
    var row = findRowIndex_(sh, id);
    if (row < 0) fail_('Record not found', 'NOT_FOUND');
    var cols = SCHEMA[table];
    var current = sh.getRange(row, 1, 1, cols.length).getValues()[0];
    var rec = {};
    cols.forEach(function (c, i) { rec[c.n] = fromCell_(current[i], c.t); });
    var protectedCols = ['id', 'createdAt'];
    cols.forEach(function (c) {
      if (protectedCols.indexOf(c.n) >= 0) return;
      if (data[c.n] === undefined) return;
      var v = data[c.n];
      if (c.t === 's' || c.t === 't') v = (v === '' || v === null) ? null : String(v).slice(0, MAX_FIELD);
      rec[c.n] = v;
    });
    if (rec.updatedAt !== undefined) rec.updatedAt = new Date().toISOString();
    sh.getRange(row, 1, 1, cols.length).setValues([rowFromObject_(table, rec)]);
    cacheDrop_(table);
    return rec;
  });
}

function adminDelete_(table, id) {
  requireTable_(table, true);
  if (!id) fail_('id required', 'VALIDATION');
  return withLock_(function () {
    var sh = sheetFor_(table);
    var row = findRowIndex_(sh, id);
    if (row < 0) fail_('Record not found', 'NOT_FOUND');
    sh.deleteRow(row);
    cacheDrop_(table);
    return { ok: true };
  });
}

function siteContentSave_(section, data) {
  if (!section) fail_('Section required', 'VALIDATION');
  var allowed = ['eyebrow', 'title', 'titleAccent', 'titleEnd', 'description', 'buttonText', 'buttonLink', 'buttonText2', 'buttonLink2', 'image'];
  return withLock_(function () {
    var sh = sheetFor_('SiteContent');
    var rows = readSheetRows_('SiteContent');
    var existing = rows.filter(function (r) { return r.section === section; })[0];
    var patch = {};
    allowed.forEach(function (k) {
      if (data[k] !== undefined) patch[k] = data[k] === '' ? null : data[k];
    });
    if (!existing) {
      var base = { section: section };
      allowed.forEach(function (k) { if (patch[k] !== undefined) base[k] = patch[k]; });
      var rec = buildRecord_('SiteContent', base, null, true);
      appendRecord_('SiteContent', rec);
      return rec;
    }
    var row = findRowIndex_(sh, existing.id);
    var merged = cleanRecord_(existing);
    allowed.forEach(function (k) { if (patch[k] !== undefined) merged[k] = patch[k]; });
    merged.updatedAt = new Date().toISOString();
    sh.getRange(row, 1, 1, SCHEMA.SiteContent.length).setValues([rowFromObject_('SiteContent', merged)]);
    cacheDrop_('SiteContent');
    return merged;
  });
}

function adminStats_() {
  var names = ['Breed', 'CatShop', 'CatFoodShop', 'VeterinaryClinic', 'ServiceProvider', 'Article', 'Event',
    'Gallery', 'Testimonial', 'CommunityPost', 'LostFound', 'Member', 'Contact', 'Newsletter'];
  var t = {};
  names.forEach(function (n) { t[n] = readTable_(n); });
  var byNewest = function (a, b) { return String(a.createdAt) < String(b.createdAt) ? 1 : -1; };
  return {
    counts: {
      breeds: t.Breed.length, catShops: t.CatShop.length, catFoodShops: t.CatFoodShop.length,
      vetClinics: t.VeterinaryClinic.length, serviceProviders: t.ServiceProvider.length,
      articles: t.Article.length, events: t.Event.length, gallery: t.Gallery.length,
      testimonials: t.Testimonial.length, communityPosts: t.CommunityPost.length,
      lostFound: t.LostFound.length, members: t.Member.length, contacts: t.Contact.length,
      newsletter: t.Newsletter.length
    },
    alerts: {
      activeLost: t.LostFound.filter(function (r) { return r.status === 'Active'; }).length,
      verifiedShops: t.CatShop.filter(function (r) { return r.verified === 'Verified'; }).length,
      verifiedVets: t.VeterinaryClinic.filter(function (r) { return r.verified === 'Verified'; }).length
    },
    recent: {
      posts: t.CommunityPost.slice().sort(byNewest).slice(0, 5),
      contacts: t.Contact.slice().sort(byNewest).slice(0, 5)
    }
  };
}

// ---------------------------------------------------------------------------
// Image upload (saved in your Google Drive, shared as "anyone with the link")
// ---------------------------------------------------------------------------
function uploadFolder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('UPLOAD_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recreate below */ } }
  var folder = DriveApp.createFolder(UPLOAD_FOLDER_NAME);
  props.setProperty('UPLOAD_FOLDER_ID', folder.getId());
  return folder;
}

function upload_(name, mime, base64) {
  var okMime = ['image/jpeg', 'image/png', 'image/webp', 'image/gif'];
  if (okMime.indexOf(mime) < 0) fail_('Only image files allowed (JPEG, PNG, WebP, GIF)', 'VALIDATION');
  if (!base64) fail_('No file provided', 'VALIDATION');
  var bytes = Utilities.base64Decode(base64);
  if (bytes.length > UPLOAD_MAX_BYTES) fail_('File too large (max 6MB)', 'VALIDATION');
  var ext = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp', 'image/gif': 'gif' }[mime];
  var filename = 'upload-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8) + '.' + ext;
  var file = uploadFolder_().createFile(Utilities.newBlob(bytes, mime, filename));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: 'https://lh3.googleusercontent.com/d/' + file.getId(), filename: filename };
}

// ---------------------------------------------------------------------------
// Miau AI (Google Gemini). The key is read from Script Properties.
// ---------------------------------------------------------------------------
var GEMINI_API = 'https://generativelanguage.googleapis.com/v1beta';

var CHAT_SYSTEM = 'You are "Miau" — the AI Cat Care Assistant for BiralBond, Bangladesh\'s premium cat lovers platform.\n\n' +
  'Your role: give warm, practical, vet-informed cat care advice tailored for Bangladeshi cat parents.\n\n' +
  'Guidelines:\n' +
  '- Be friendly, concise and genuinely helpful. Use a warm tone. A little cat-themed warmth is welcome but don\'t overdo emojis.\n' +
  '- You can answer in English, Bangla (Bengali), or Banglish — mirror the user\'s language. Many users write in Banglish (Bengali in roman script).\n' +
  '- Focus on everyday cat care: nutrition, grooming, litter training, behaviour, health basics, kitten care, senior cats, seasonal care (monsoon, summer heat), breed-specific advice.\n' +
  '- Reference Bangladesh context when relevant: local vet availability, products available locally (Whiskas, Royal Canin, Friskies), local climate, deshi vs pedigree cats.\n' +
  '- Keep answers short and scannable — use short paragraphs or bullet points. Max ~180 words unless the user asks for detail.\n' +
  '- For anything urgent or medical (suspected poisoning, serious injury, difficulty breathing, not eating for 24h+), ALWAYS advise seeing a vet immediately and mention BiralBond\'s Vet Directory has 24/7 emergency clinics. Never diagnose or prescribe medication.\n' +
  '- You are not a substitute for a veterinarian. Add a gentle "see a vet for certainty" note for health questions.\n' +
  '- If asked about non-cat topics, gently redirect back to cat care.\n\n' +
  'You are the heart of BiralBond\'s community — caring, knowledgeable and always rooting for every cat.';

var NAMES_SYSTEM = 'You are a playful cat name generator for BiralBond, a Bangladeshi cat lovers platform.\n\n' +
  'Generate exactly 8 unique, creative cat names based on the user\'s preferences. Mix Bengali and English names appropriate for Bangladesh.\n\n' +
  'Rules:\n' +
  '- Return ONLY a JSON array of 8 strings, no other text.\n' +
  '- Each name should be short (1-3 words max), memorable and cute.\n' +
  '- Vary the style: some cute, some royal, some food-inspired, some descriptive.\n' +
  '- Include at least 2 Bengali/Banglish names (romanized) when a Bengali vibe is requested.';

// Google retires Gemini models now and then, so this list is only the first guess.
// If all of them fail with "not found", the script asks Google which Flash models
// exist today and uses those. You can also force one by adding a script
// property named GEMINI_MODEL.
var MODEL_CHAIN = ['gemini-2.5-flash-lite', 'gemini-2.5-flash', 'gemini-3.5-flash-lite', 'gemini-3.5-flash'];

function geminiKey_() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) fail_('GEMINI_API_KEY is not set', 'AI_NOT_CONFIGURED');
  return key;
}

function gemini_(system, messages, maxTokens) {
  var key = geminiKey_();
  var contents = (messages || [])
    .filter(function (m) { return m && typeof m.content === 'string' && m.content.trim(); })
    .map(function (m) { return { role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content.slice(0, 2000) }] }; });
  while (contents.length && contents[0].role !== 'user') contents.shift();
  if (!contents.length) fail_('No user message', 'VALIDATION');

  var custom = (PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL') || '').trim();
  var chain = custom ? [custom].concat(MODEL_CHAIN.filter(function (m) { return m !== custom; })) : MODEL_CHAIN.slice();
  var attempts = [];
  var tried = {};
  for (var i = 0; i < chain.length; i++) {
    tried[chain[i]] = true;
    var text = callModel_(key, chain[i], system, contents, maxTokens, attempts);
    if (text) return text;
  }
  var found = discoverModels_(key, tried);
  for (var j = 0; j < found.length; j++) {
    var t2 = callModel_(key, found[j], system, contents, maxTokens, attempts);
    if (t2) return t2;
  }
  console.error('Gemini attempts: ' + JSON.stringify(attempts));
  fail_('AI unavailable', 'AI_FAILED');
}

function callModel_(key, model, system, contents, maxTokens, attempts) {
  var res = UrlFetchApp.fetch(GEMINI_API + '/models/' + model + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      systemInstruction: { parts: [{ text: system }] },
      contents: contents,
      generationConfig: { maxOutputTokens: maxTokens }
    })
  });
  var code = res.getResponseCode();
  var body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) { /* not json */ }
  if (code === 200) {
    var parts = (body.candidates && body.candidates[0] && body.candidates[0].content && body.candidates[0].content.parts) || [];
    var text = parts.map(function (p) { return p.thought ? '' : (p.text || ''); }).join('').trim();
    if (text) return text;
    attempts.push({ model: model, status: 200, message: 'empty response' });
    return null;
  }
  attempts.push({ model: model, status: code, message: String((body.error && body.error.message) || '').slice(0, 200) });
  if (code === 401 || code === 403) fail_('Gemini rejected the API key', 'AI_KEY');
  return null;
}

function discoverModels_(key, skip) {
  var res = UrlFetchApp.fetch(GEMINI_API + '/models?pageSize=200', { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return [];
  var data = {};
  try { data = JSON.parse(res.getContentText()); } catch (e) { return []; }
  var version = function (id) { var m = id.match(/gemini-(\d+(?:\.\d+)?)/); return m ? parseFloat(m[1]) : 0; };
  return (data.models || [])
    .filter(function (m) {
      return typeof m.name === 'string' &&
        (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0 &&
        /gemini-[\d.]+-flash/.test(m.name) &&
        !/(image|tts|live|audio|embed|robotics|computer|thinking|exp|-\d{2}-\d{4}|preview)/.test(m.name);
    })
    .map(function (m) { return m.name.replace(/^models\//, ''); })
    .filter(function (id) { return !skip[id]; })
    .sort(function (a, b) { return version(b) - version(a); })
    .slice(0, 3);
}

function aiChat_(messages) {
  if (!Array.isArray(messages) || messages.length === 0) fail_('Messages required', 'VALIDATION');
  throttle_('ai', AI_LIMIT_PER_MIN);
  return { reply: gemini_(CHAT_SYSTEM, messages.slice(-12), 1500) };
}

function aiNames_(vibe, gender, breed) {
  throttle_('ai', AI_LIMIT_PER_MIN);
  var prompt = 'Generate 8 cat names. Preferences:\n- Vibe: ' + String(vibe || 'any').slice(0, 80) +
    '\n- Gender: ' + String(gender || 'any').slice(0, 40) +
    '\n- Breed: ' + String(breed || 'any').slice(0, 80) + '\n\nReturn only the JSON array of 8 name strings.';
  var raw = gemini_(NAMES_SYSTEM, [{ role: 'user', content: prompt }], 600);
  var match = raw.match(/\[[\s\S]*\]/);
  var names = [];
  try { names = JSON.parse(match ? match[0] : raw); } catch (e) { /* fall through */ }
  if (!Array.isArray(names) || names.length === 0) fail_('Bad model output', 'AI_FAILED');
  return { names: names.map(String).slice(0, 8) };
}

// ---------------------------------------------------------------------------
// Sheet edited by hand? Drop the cache so the website sees the change quickly.
// (Runs automatically; no trigger setup needed.)
// ---------------------------------------------------------------------------
function onEdit(e) {
  try {
    if (e && e.range) cacheDrop_(e.range.getSheet().getName());
  } catch (err) { /* ignore */ }
}

// ---------------------------------------------------------------------------
// Editor helpers - run these from the Apps Script editor
// ---------------------------------------------------------------------------

/** Run ONCE (safe to run again): creates the tabs, headers and starter data. */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the script from your Sheet: Extensions > Apps Script');
  var props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  secret_();

  var created = [];
  var seeded = [];
  Object.keys(SCHEMA).forEach(function (name) {
    var cols = SCHEMA[name];
    var sh = ss.getSheetByName(name);
    var isNew = false;
    if (!sh) { sh = ss.insertSheet(name); isNew = true; created.push(name); }
    // header row
    var headers = cols.map(function (c) { return c.n; });
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3e8dc');
    sh.setFrozenRows(1);
    // text columns stay text (keeps phone numbers like 017... and +880... intact)
    cols.forEach(function (c, i) {
      if (c.t === 's' || c.t === 't') sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
    });
    if (cols[0] && cols[0].n === 'id') sh.setColumnWidth(1, 90);
    // starter data only into empty tabs
    var rows = SEED[name];
    if (rows && rows.length && sh.getLastRow() < 2) {
      var values = rows.map(function (r) { return rowFromObject_(name, r); });
      sh.getRange(2, 1, values.length, cols.length).setValues(values);
      seeded.push(name + ' (' + rows.length + ')');
    }
    cacheDrop_(name);
  });

  // remove the empty default tab ("Sheet1") if it is still there
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  uploadFolder_();   // creates the Drive folder for admin image uploads (asks for Drive permission)
  console.log('✅ Setup complete.');
  console.log('New tabs: ' + (created.join(', ') || 'none (already existed)'));
  console.log('Starter data added to: ' + (seeded.join(', ') || 'none (tabs already had data)'));
  checkSetup();
}

/** Prints a health report of your setup to the execution log. */
function checkSetup() {
  var props = PropertiesService.getScriptProperties();
  var problems = [];
  if (!props.getProperty('SHEET_ID')) problems.push('❌ Run setup() first.');
  if (!props.getProperty('ADMIN_EMAIL')) problems.push('❌ Script property ADMIN_EMAIL is missing.');
  if (!props.getProperty('ADMIN_PASSWORD')) problems.push('❌ Script property ADMIN_PASSWORD is missing.');
  else if (props.getProperty('ADMIN_PASSWORD').length < 8) problems.push('⚠️ ADMIN_PASSWORD is shorter than 8 characters - use a stronger one.');
  if (!props.getProperty('GEMINI_API_KEY')) console.log('ℹ️ GEMINI_API_KEY not set: Miau AI chat will be off (the rest of the site works).');
  if (props.getProperty('SHEET_ID')) {
    try {
      var missing = Object.keys(SCHEMA).filter(function (n) { return !ss_().getSheetByName(n); });
      if (missing.length) problems.push('❌ Missing tabs: ' + missing.join(', ') + ' - run setup() again.');
      else {
        var b = bootstrap_();
        var n = 0;
        Object.keys(b.tables).forEach(function (t) { n += b.tables[t].length; });
        console.log('✅ Sheet OK - public tabs hold ' + n + ' records in total.');
      }
    } catch (e) { problems.push('❌ ' + e.message); }
  }
  if (problems.length) console.log(problems.join('\n')); else console.log('✅ Everything looks good. Now deploy as a Web app (see DEPLOY.md).');
}

/** Sends one test question to Gemini and prints the answer (or the reason it failed). */
function testAi() {
  try {
    console.log('✅ Gemini replied: ' + gemini_(CHAT_SYSTEM, [{ role: 'user', content: 'Say hi in one short sentence.' }], 200));
  } catch (e) {
    console.log('❌ ' + e.message + ' (' + e.code + '). Open View > Logs / Executions for details.');
  }
}

// ===========================================================================
//  DATA MODEL + STARTER DATA  (generated - do not edit below this line)
//  SCHEMA: n = column name, t = type (s text, i integer, n number, b boolean,
//  t timestamp text), req = required, d = default value.
// ===========================================================================

var SCHEMA = {
  Article: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"title","t":"s","req":true},{"n":"category","t":"s","req":true},{"n":"excerpt","t":"s","req":true},{"n":"content","t":"s","req":true},{"n":"author","t":"s","req":true},{"n":"readTime","t":"i","d":5},{"n":"image","t":"s"},{"n":"tags","t":"s"},{"n":"featured","t":"b","d":false},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  Breed: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"name","t":"s","req":true},{"n":"bnName","t":"s"},{"n":"origin","t":"s"},{"n":"temperament","t":"s"},{"n":"lifespan","t":"s"},{"n":"weight","t":"s"},{"n":"rarity","t":"s","d":"Common"},{"n":"category","t":"s","d":"Pedigree"},{"n":"description","t":"s","req":true},{"n":"careLevel","t":"s","d":"Moderate"},{"n":"goodWithKids","t":"b","d":true},{"n":"hypoallergenic","t":"b","d":false},{"n":"image","t":"s"},{"n":"accent","t":"s"},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}},{"n":"price","t":"s"}],
  CareTip: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"title","t":"s","req":true},{"n":"body","t":"s","req":true},{"n":"icon","t":"s","d":"🐾"},{"n":"category","t":"s","d":"Daily"},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  CatFoodShop: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"shopName","t":"s","req":true},{"n":"ownerName","t":"s"},{"n":"shopImage","t":"s"},{"n":"address","t":"s","req":true},{"n":"division","t":"s","req":true},{"n":"district","t":"s","req":true},{"n":"area","t":"s","req":true},{"n":"phone","t":"s","req":true},{"n":"whatsapp","t":"s"},{"n":"facebook","t":"s"},{"n":"website","t":"s"},{"n":"mapUrl","t":"s"},{"n":"openingHours","t":"s"},{"n":"homeDelivery","t":"b","d":false},{"n":"availableBrands","t":"s"},{"n":"foodTypes","t":"s"},{"n":"packSizes","t":"s"},{"n":"priceInfo","t":"s"},{"n":"description","t":"s"},{"n":"verified","t":"s","d":"Not Verified"},{"n":"rating","t":"n","d":0},{"n":"reviewCount","t":"i","d":0},{"n":"lastUpdated","t":"t","d":{"now":1}},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  CatShop: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"shopName","t":"s","req":true},{"n":"ownerName","t":"s"},{"n":"shopImage","t":"s"},{"n":"catImage","t":"s"},{"n":"address","t":"s","req":true},{"n":"division","t":"s","req":true},{"n":"district","t":"s","req":true},{"n":"upazila","t":"s"},{"n":"area","t":"s","req":true},{"n":"phone","t":"s","req":true},{"n":"whatsapp","t":"s"},{"n":"facebook","t":"s"},{"n":"website","t":"s"},{"n":"mapUrl","t":"s"},{"n":"openingHours","t":"s"},{"n":"availableBreeds","t":"s"},{"n":"description","t":"s"},{"n":"verified","t":"s","d":"Not Verified"},{"n":"rating","t":"n","d":0},{"n":"reviewCount","t":"i","d":0},{"n":"lastUpdated","t":"t","d":{"now":1}},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  Comment: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"postId","t":"s","req":true},{"n":"author","t":"s","req":true},{"n":"avatar","t":"s"},{"n":"body","t":"s","req":true},{"n":"createdAt","t":"t","d":{"now":1}}],
  CommunityPost: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"author","t":"s","req":true},{"n":"avatar","t":"s"},{"n":"title","t":"s","req":true},{"n":"body","t":"s","req":true},{"n":"category","t":"s","d":"Discussion"},{"n":"city","t":"s"},{"n":"likes","t":"i","d":0},{"n":"comments","t":"i","d":0},{"n":"image","t":"s"},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  Contact: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"name","t":"s","req":true},{"n":"email","t":"s","req":true},{"n":"subject","t":"s","req":true},{"n":"message","t":"s","req":true},{"n":"createdAt","t":"t","d":{"now":1}}],
  Event: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"title","t":"s","req":true},{"n":"type","t":"s","req":true},{"n":"date","t":"s","req":true},{"n":"time","t":"s","req":true},{"n":"venue","t":"s","req":true},{"n":"city","t":"s","req":true},{"n":"description","t":"s","req":true},{"n":"capacity","t":"i","d":50},{"n":"registered","t":"i","d":0},{"n":"image","t":"s"},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  FeatureToggle: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"feature","t":"s","req":true},{"n":"enabled","t":"b","d":false},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  Gallery: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"title","t":"s","req":true},{"n":"cat","t":"s"},{"n":"owner","t":"s","req":true},{"n":"city","t":"s","req":true},{"n":"image","t":"s","req":true},{"n":"likes","t":"i","d":0},{"n":"caption","t":"s"},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  LostFound: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"type","t":"s","req":true},{"n":"catName","t":"s"},{"n":"breed","t":"s"},{"n":"color","t":"s"},{"n":"location","t":"s","req":true},{"n":"city","t":"s","req":true},{"n":"date","t":"s","req":true},{"n":"contact","t":"s","req":true},{"n":"phone","t":"s","req":true},{"n":"description","t":"s","req":true},{"n":"reward","t":"s"},{"n":"image","t":"s"},{"n":"status","t":"s","d":"Active"},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  Member: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"name","t":"s","req":true},{"n":"email","t":"s","req":true},{"n":"phone","t":"s"},{"n":"city","t":"s","req":true},{"n":"plan","t":"s","d":"Free"},{"n":"cats","t":"i","d":0},{"n":"joinedAt","t":"t","d":{"now":1}},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  Newsletter: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"email","t":"s","req":true},{"n":"createdAt","t":"t","d":{"now":1}}],
  ServiceProvider: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"providerName","t":"s","req":true},{"n":"businessName","t":"s"},{"n":"type","t":"s","req":true},{"n":"image","t":"s"},{"n":"address","t":"s","req":true},{"n":"division","t":"s","req":true},{"n":"district","t":"s","req":true},{"n":"area","t":"s","req":true},{"n":"phone","t":"s","req":true},{"n":"whatsapp","t":"s"},{"n":"facebook","t":"s"},{"n":"website","t":"s"},{"n":"mapUrl","t":"s"},{"n":"openingHours","t":"s"},{"n":"homeService","t":"b","d":false},{"n":"description","t":"s"},{"n":"verified","t":"s","d":"Not Verified"},{"n":"rating","t":"n","d":0},{"n":"reviewCount","t":"i","d":0},{"n":"lastUpdated","t":"t","d":{"now":1}},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  SiteContent: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"section","t":"s","req":true},{"n":"eyebrow","t":"s"},{"n":"title","t":"s"},{"n":"titleAccent","t":"s"},{"n":"titleEnd","t":"s"},{"n":"description","t":"s"},{"n":"buttonText","t":"s"},{"n":"buttonLink","t":"s"},{"n":"buttonText2","t":"s"},{"n":"buttonLink2","t":"s"},{"n":"image","t":"s"},{"n":"updatedAt","t":"t","d":{"now":1}},{"n":"createdAt","t":"t","d":{"now":1}}],
  Testimonial: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"name","t":"s","req":true},{"n":"role","t":"s","req":true},{"n":"city","t":"s","req":true},{"n":"avatar","t":"s"},{"n":"rating","t":"i","d":5},{"n":"quote","t":"s","req":true},{"n":"catName","t":"s"},{"n":"featured","t":"b","d":false},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}}],
  VeterinaryClinic: [{"n":"id","t":"s","d":{"uuid":1}},{"n":"clinicName","t":"s","req":true},{"n":"doctorName","t":"s"},{"n":"specialization","t":"s"},{"n":"address","t":"s","req":true},{"n":"division","t":"s","req":true},{"n":"district","t":"s","req":true},{"n":"area","t":"s","req":true},{"n":"phone","t":"s","req":true},{"n":"emergencyPhone","t":"s"},{"n":"whatsapp","t":"s"},{"n":"openingHours","t":"s"},{"n":"emergencyService","t":"b","d":false},{"n":"mapUrl","t":"s"},{"n":"website","t":"s"},{"n":"facebook","t":"s"},{"n":"services","t":"s"},{"n":"description","t":"s"},{"n":"verified","t":"s","d":"Not Verified"},{"n":"rating","t":"n","d":0},{"n":"reviewCount","t":"i","d":0},{"n":"lastUpdated","t":"t","d":{"now":1}},{"n":"createdAt","t":"t","d":{"now":1}},{"n":"updatedAt","t":"t","d":{"now":1}},{"n":"clinicImage","t":"s"}]
};

var SEED = {
  SiteContent: [{"id":"cmunl3d7f0000ne808q6glt08","section":"hero","eyebrow":"বাংলাদেশের #১ বিড়াল কমিউনিটি","title":"বাংলাদেশে বিড়াল, খাবার ও","titleAccent":"ভেট কোথায় পাবেন?","titleEnd":"এখানেই খুঁজে নিন।","description":"বিরালবন্ড হলো বাংলাদেশের বিড়াল তথ্য ও অনুসন্ধান প্ল্যাটফর্ম। কোথায় বিড়াল পাওয়া যায়, কোথায় বিড়ালের খাবার পাওয়া যায়, কোথায় পশু হাসপাতাল আছে — সব তথ্য এক জায়গায়।","buttonText":"বিড়াল কোথায় পাব","buttonLink":"/breeds","buttonText2":"ভেট খুঁজুন","buttonLink2":"/vets","image":null,"updatedAt":"2026-09-30T04:14:11.144Z","createdAt":"2026-09-30T04:07:11.260Z"},{"id":"cmunl3d7h0001ne80eolkgdis","section":"breeds","eyebrow":"বিড়ালের জাত","title":"আপনার জন্য সেরা","titleAccent":"জাত","titleEnd":"খুঁজে নিন","description":"অভিজাত পারসিয়ান থেকে আমাদের দৃঢ় দেশি বিরাল — বাংলাদেশের সব প্রিয় বিড়াল দেখুন এবং আপনার মনের মতো সঙ্গী খুঁজে নিন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.262Z","createdAt":"2026-09-30T04:07:11.262Z"},{"id":"cmunl3d7i0002ne80b51k0k2x","section":"food","eyebrow":"বিড়ালের খাবার কোথায় পাওয়া যায়","title":"ক্যাট","titleAccent":"ফুড ফাইন্ডার","titleEnd":null,"description":"কোথায় বিড়ালের খাবার পাওয়া যায় তা খুঁজুন। ব্র্যান্ড ও এলাকা অনুযায়ী খাবারের দোকানের তথ্য পান।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.263Z","createdAt":"2026-09-30T04:07:11.263Z"},{"id":"cmunl3d7j0003ne80mrgcd5z1","section":"vets","eyebrow":"ভেট ও হাসপাতাল","title":"পশু","titleAccent":"চিকিৎসা কেন্দ্র","titleEnd":null,"description":"বাংলাদেশের বিভিন্ন এলাকার পশু হাসপাতাল, ভেট ক্লিনিক ও পশু চিকিৎসকদের তথ্য খুঁজুন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.264Z","createdAt":"2026-09-30T04:07:11.264Z"},{"id":"cmunl3d7k0004ne80rnisb8uq","section":"services","eyebrow":"সেবা প্রদানকারী","title":"বিড়াল সেবা","titleAccent":"ফাইন্ডার","titleEnd":null,"description":"গ্রুমিং, বোর্ডিং, ট্রেনিং, ফটোগ্রাফি সহ বিড়াল সংক্রান্ত সব সেবা খুঁজুন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.264Z","createdAt":"2026-09-30T04:07:11.264Z"},{"id":"cmunl3d7l0005ne80389m2uck","section":"lostfound","eyebrow":"হারানো ও পাওয়া","title":"হারানো বিড়ালদের","titleAccent":"একসাথে ঘরে ফিরিয়ে আনা","titleEnd":null,"description":"আপনার বিড়াল হারিয়েছে? কোনো বেওয়ারিশ পেয়েছেন? এখানে পোস্ট করুন এবং কমিউনিটির সাহায্য নিন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.265Z","createdAt":"2026-09-30T04:07:11.265Z"},{"id":"cmunl3d7l0006ne80vuk4lvh4","section":"care","eyebrow":"বিড়াল যত্ন লাইব্রেরি","title":"বিশেষজ্ঞ গাইড, লেখা হয়েছে","titleAccent":"বিড়ালপ্রেমীদের দ্বারা","titleEnd":null,"description":"পুষ্টি, স্বাস্থ্য, আচরণ, গ্রুমিং সহ আরও অনেক কিছুর উপর ভেট-পর্যালোচিত প্রবন্ধ।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.266Z","createdAt":"2026-09-30T04:07:11.266Z"},{"id":"cmunl3d7m0007ne80drlxafus","section":"community","eyebrow":"কমিউনিটি ফোরাম","title":"যেখানে বিড়াল অভিভাবকরা","titleAccent":"সংযুক্ত হন","titleEnd":null,"description":"প্রশ্ন করুন, গল্প শেয়ার করুন, ছবি পোস্ট করুন এবং সারাবাংলার সহৃদয় বিড়ালপ্রেমীদের সাথে পরিচিত হন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.267Z","createdAt":"2026-09-30T04:07:11.267Z"},{"id":"cmunl3d7n0008ne80dwss7nzs","section":"events","eyebrow":"ইভেন্ট ও মিটআপ","title":"ক্যাট শো, মিটআপ ও","titleAccent":"আরও","titleEnd":null,"description":"বাংলাদেশের বিড়াল অভিভাবকদের উষ্ণতম সমাবেশে যোগ দিন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.268Z","createdAt":"2026-09-30T04:07:11.268Z"},{"id":"cmunl3d7o0009ne80wp4vov32","section":"faq","eyebrow":"প্রশ্নোত্তর","title":"প্রশ্ন,","titleAccent":"উত্তর দেওয়া হলো","titleEnd":null,"description":"বাংলাদেশের বিড়াল অভিভাবকরা আমাদের সবচেয়ে সাধারণ প্রশ্ন জিজ্ঞাসা করেন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.269Z","createdAt":"2026-09-30T04:07:11.269Z"},{"id":"cmunl3d7q000ane809ruwsfaw","section":"membership","eyebrow":"সদস্যপদ","title":"হয়ে যান","titleAccent":"বিরালবন্ড সদস্য","titleEnd":null,"description":"ছাড়, অগ্রাধিকার সেবা এবং বাংলাদেশের উষ্ণতম বিড়াল অভিভাবক কমিউনিটির সদস্য হন।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.270Z","createdAt":"2026-09-30T04:07:11.270Z"},{"id":"cmunl3d7r000bne80w4q457f8","section":"contact","eyebrow":"বিরালবন্ড সম্পর্কে","title":"বিড়ালপ্রেমীদের দ্বারা তৈরি,","titleAccent":"বিড়ালপ্রেমীদের জন্য","titleEnd":null,"description":"বিরালবন্ড ঢাকার একটি ফ্ল্যাটে একটি উদ্ধারকৃত কিটেন এবং একটি সরল বিশ্বাস নিয়ে শুরু হয়েছিল।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.271Z","createdAt":"2026-09-30T04:07:11.271Z"},{"id":"cmunl3d7s000cne80y8uhk5jv","section":"footer","eyebrow":null,"title":null,"titleAccent":null,"titleEnd":null,"description":"বাংলাদেশের বিড়ালপ্রেমীদের জন্য প্রিমিয়াম প্ল্যাটফর্ম।","buttonText":null,"buttonLink":null,"buttonText2":null,"buttonLink2":null,"image":null,"updatedAt":"2026-09-30T04:07:11.272Z","createdAt":"2026-09-30T04:07:11.272Z"},{"id":"cmunl3d7t000dne80tb7l0dom","section":"aiBanner","eyebrow":"এআই-চালিত · সবার জন্য ফ্রি","title":"মিয়াও-এর সাথে পরিচিত হোন — আপনার ব্যক্তিগত","titleAccent":"বিড়াল যত্ন বিশেষজ্ঞ","titleEnd":null,"description":"রাত ২টায় বিড়াল সম্পর্কে প্রশ্ন আছে? মিয়াও আমাদের এআই সহকারী সবসময় জেগে আছে।","buttonText":"মিয়াও-এর সাথে চ্যাট করুন","buttonLink":null,"buttonText2":"যত্ন গাইড পড়ুন","buttonLink2":"/care","image":null,"updatedAt":"2026-09-30T04:07:11.273Z","createdAt":"2026-09-30T04:07:11.273Z"}],
  Breed: [{"id":"persian-seed","name":"Persian","bnName":"পারসিয়ান বিরাল","origin":"Iran (Persia)","temperament":"Calm, gentle, affectionate","lifespan":"12–17 years","weight":"3–5.5 kg","rarity":"Premium","category":"Pedigree","description":"The Persian is the aristocrat of the cat world — beloved in Dhaka households for its luxurious long coat, sweet round face and serene temperament. Perfect for apartment living and families with children.","careLevel":"High","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1518791841217-8f162f1e1131?w=800&q=80","accent":"amber","createdAt":"2026-09-27T03:02:51.892Z","updatedAt":"2026-09-27T03:02:51.892Z"},{"id":"maine-coon-seed","name":"Maine Coon","bnName":"মেইন কুন","origin":"Maine, USA","temperament":"Friendly, playful, dog-like","lifespan":"12–15 years","weight":"5–11 kg","rarity":"Exotic","category":"Pedigree","description":"Known as the \"gentle giants\" of the cat world, Maine Coons are large, sociable cats with tufted ears and bushy tails. They love water and follow their humans from room to room.","careLevel":"Moderate","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1592194996308-7b43878e84a6?w=800&q=80","accent":"orange","createdAt":"2026-09-27T03:02:51.895Z","updatedAt":"2026-09-27T03:02:51.895Z"},{"id":"british-shorthair-seed","name":"British Shorthair","bnName":"ব্রিটিশ শর্টহেয়ার","origin":"United Kingdom","temperament":"Quiet, loyal, easygoing","lifespan":"14–20 years","weight":"4–8 kg","rarity":"Premium","category":"Pedigree","description":"With their plush teddy-bear coats and round copper eyes, British Shorthairs are calm, dignified companions. They are undemanding yet deeply loyal to their families.","careLevel":"Easy","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1574158622682-e40e69881006?w=800&q=80","accent":"rose","createdAt":"2026-09-27T03:02:51.896Z","updatedAt":"2026-09-27T03:02:51.896Z"},{"id":"siamese-seed","name":"Siamese","bnName":"সিয়ামিজ","origin":"Thailand (Siam)","temperament":"Vocal, social, intelligent","lifespan":"12–20 years","weight":"3–5 kg","rarity":"Premium","category":"Pedigree","description":"Siamese cats are famous for their striking blue almond eyes and chatty personalities. They form intense bonds with their humans and love to \"converse\" throughout the day.","careLevel":"Moderate","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1561948955-570b270e7c36?w=800&q=80","accent":"cyan","createdAt":"2026-09-27T03:02:51.897Z","updatedAt":"2026-09-27T03:02:51.897Z"},{"id":"bengal-seed","name":"Bengal","bnName":"বেঙ্গল","origin":"United States (Asian Leopard Cat cross)","temperament":"Energetic, athletic, curious","lifespan":"12–16 years","weight":"4–8 kg","rarity":"Exotic","category":"Pedigree","description":"With their wild leopard-like rosettes, Bengals look like miniature jungle cats. They are highly active, intelligent and thrive with lots of play and vertical space.","careLevel":"High","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1606214174585-fe31582dc6ee?w=800&q=80","accent":"amber","createdAt":"2026-09-27T03:02:51.898Z","updatedAt":"2026-09-27T03:02:51.898Z"},{"id":"ragdoll-seed","name":"Ragdoll","bnName":"র‍্যাগডল","origin":"California, USA","temperament":"Docile, affectionate, floppy","lifespan":"13–18 years","weight":"4.5–9 kg","rarity":"Exotic","category":"Pedigree","description":"Ragdolls literally go limp with relaxation when picked up — hence the name. They are devoted, dog-like companions who greet their humans at the door.","careLevel":"Moderate","goodWithKids":true,"hypoallergenic":false,"image":"/cats/ragdoll.png","accent":"violet","createdAt":"2026-09-27T03:02:51.899Z","updatedAt":"2026-09-27T03:18:39.414Z"},{"id":"bangladeshi-local--biral--seed","name":"Bangladeshi Local (Biral)","bnName":"দেশি বিরাল","origin":"Bangladesh","temperament":"Hardy, independent, affectionate","lifespan":"12–18 years","weight":"3–5 kg","rarity":"Common","category":"Local","description":"Our very own deshi biral — resilient, smart and wonderfully adapted to the Bangladeshi climate. These short-haired beauties are low-maintenance, healthy and full of personality. Adopt, never shop!","careLevel":"Easy","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1495360010541-f48722b34f7d?w=800&q=80","accent":"orange","createdAt":"2026-09-27T03:02:51.901Z","updatedAt":"2026-09-27T03:02:51.901Z"},{"id":"scottish-fold-seed","name":"Scottish Fold","bnName":"স্কটিশ ফোল্ড","origin":"Scotland","temperament":"Sweet, calm, attached","lifespan":"12–15 years","weight":"3–6 kg","rarity":"Exotic","category":"Pedigree","description":"Recognisable by their adorable folded ears, Scottish Folds are sweet-tempered and love being involved in whatever their humans are doing — often sitting up on their hind legs like little meerkats.","careLevel":"Moderate","goodWithKids":true,"hypoallergenic":false,"image":"https://images.unsplash.com/photo-1573865526739-10659fec78a5?w=800&q=80","accent":"rose","createdAt":"2026-09-27T03:02:51.902Z","updatedAt":"2026-09-27T03:02:51.902Z"}],
  CatShop: [{"id":"cmujok19u0000oecihw4rdhci","shopName":"Dhaka Cat House","ownerName":"Rahim Uddin","shopImage":null,"catImage":null,"address":"House 12, Road 5, Dhanmondi","division":"Dhaka","district":"Dhaka","upazila":null,"area":"Dhanmondi","phone":"+8801711111111","whatsapp":"+8801711111111","facebook":"dhakacathouse","website":null,"mapUrl":null,"openingHours":"10am–8pm","availableBreeds":"Persian, Maine Coon, British Shorthair","description":"প্রিমিয়াম বিড়ালের বিশেষায়িত দোকান। সব বিড়াল টিকাকৃত ও স্বাস্থ্যকর্মী।","verified":"Verified","rating":4.8,"reviewCount":124,"lastUpdated":"2026-09-27T10:33:03.091Z","createdAt":"2026-09-27T10:33:03.091Z","updatedAt":"2026-09-27T10:33:03.091Z"},{"id":"cmujok19x0001oeci576rrh5x","shopName":"Chattogram Pet Center","ownerName":"Karim Mia","shopImage":null,"catImage":null,"address":"GEC Circle, Khulshi","division":"Chattogram","district":"Chattogram","upazila":null,"area":"Khulshi","phone":"+8801822222222","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"9am–9pm","availableBreeds":"Persian, Siamese, Local","description":"চট্টগ্রামের নির্ভরযোগ্য বিড়ালের দোকান। দেশি ও বিদেশি জাতের বিড়াল পাওয়া যায়।","verified":"Verified","rating":4.6,"reviewCount":87,"lastUpdated":"2026-09-27T10:33:03.093Z","createdAt":"2026-09-27T10:33:03.093Z","updatedAt":"2026-09-27T10:33:03.093Z"},{"id":"cmujok19y0002oecidy937sbd","shopName":"Sylhet Cat World","ownerName":"Jamal Ahmed","shopImage":null,"catImage":null,"address":"Zindabazar, Sylhet","division":"Sylhet","district":"Sylhet","upazila":null,"area":"Zindabazar","phone":"+8801933333333","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"10am–8pm","availableBreeds":"Persian, Ragdoll, Bengal","description":"সিলেটের সেরা বিড়ালের দোকান। বাচ্চা বিড়াল ও প্রাপ্তবয়স্ক বিড়াল উপলব্ধ।","verified":"Pending","rating":4.4,"reviewCount":56,"lastUpdated":"2026-09-27T10:33:03.094Z","createdAt":"2026-09-27T10:33:03.094Z","updatedAt":"2026-09-27T10:33:03.094Z"},{"id":"cmujok19z0003oeci9vz6odtf","shopName":"Rajshahi Pet Bari","ownerName":"Salam Sheikh","shopImage":null,"catImage":null,"address":"Kazla, Rajshahi","division":"Rajshahi","district":"Rajshahi","upazila":null,"area":"Kazla","phone":"+8801544444444","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"9am–7pm","availableBreeds":"Local, Persian Mix","description":"দেশি বিরাল ও মিক্স জাতের বিড়াল পাওয়া যায়। দত্তক নেওয়ার সুবিধাও আছে।","verified":"Not Verified","rating":4.2,"reviewCount":34,"lastUpdated":"2026-09-27T10:33:03.095Z","createdAt":"2026-09-27T10:33:03.095Z","updatedAt":"2026-09-27T10:33:03.095Z"},{"id":"cmujok1a00004oeci6uzaar7h","shopName":"Khulna Cat Corner","ownerName":"Babul Hossain","shopImage":null,"catImage":null,"address":"Boyra, Khulna","division":"Khulna","district":"Khulna","upazila":null,"area":"Boyra","phone":"+8801655555555","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"10am–8pm","availableBreeds":"Local, British Shorthair","description":"খুলনায় বিড়াল পাওয়ার নির্ভরযোগ্য ঠিকানা।","verified":"Verified","rating":4.5,"reviewCount":67,"lastUpdated":"2026-09-27T10:33:03.096Z","createdAt":"2026-09-27T10:33:03.096Z","updatedAt":"2026-09-27T10:33:03.096Z"},{"id":"cmujok1a10005oeci8mtfazsk","shopName":"Mirpur Cat Garden","ownerName":"Shahidul Islam","shopImage":null,"catImage":null,"address":"Road 2, Mirpur 10, Dhaka","division":"Dhaka","district":"Dhaka","upazila":null,"area":"Mirpur","phone":"+8801766666666","whatsapp":"+8801766666666","facebook":"mirpurcatgarden","website":null,"mapUrl":null,"openingHours":"10am–9pm","availableBreeds":"Persian, Maine Coon, Scottish Fold, Bengal","description":"মিরপুরে বিভিন্ন জাতের বিড়াল পাওয়া যায়। সব বিড়াল টিকাকৃত ও নিবন্ধিত।","verified":"Verified","rating":4.7,"reviewCount":145,"lastUpdated":"2026-09-27T10:33:03.097Z","createdAt":"2026-09-27T10:33:03.097Z","updatedAt":"2026-09-27T10:33:03.097Z"}],
  CatFoodShop: [{"id":"cmujok1a60006oeci28uo6wae","shopName":"Pet Food BD","ownerName":"Nasir Uddin","shopImage":null,"address":"Gulshan 2, Dhaka","division":"Dhaka","district":"Dhaka","area":"Gulshan","phone":"+8801777777777","whatsapp":"+8801777777777","facebook":"petfoodbd","website":null,"mapUrl":null,"openingHours":"9am–9pm","homeDelivery":true,"availableBrands":"Royal Canin, Whiskas, Me-O, Purina, SmartHeart","foodTypes":"Dry, Wet, Kitten, Adult, Senior","packSizes":"400g, 1kg, 2kg, 7kg","priceInfo":"Royal Canin 2kg ৳2400, Whiskas 1.2kg ৳720","description":"ঢাকায় সব ব্র্যান্ডের বিড়ালের খাবার পাওয়া যায়। হোম ডেলিভারি সুবিধা।","verified":"Verified","rating":4.8,"reviewCount":234,"lastUpdated":"2026-09-27T10:33:03.103Z","createdAt":"2026-09-27T10:33:03.103Z","updatedAt":"2026-09-27T10:33:03.103Z"},{"id":"cmujok1a80007oeciheogn0d2","shopName":"Mirpur Pet Supply","ownerName":"Faruk Hossain","shopImage":null,"address":"Mirpur 1, Dhaka","division":"Dhaka","district":"Dhaka","area":"Mirpur","phone":"+8801888888888","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"10am–8pm","homeDelivery":true,"availableBrands":"Whiskas, Me-O, SmartHeart, Friskies","foodTypes":"Dry, Wet, Kitten","packSizes":"500g, 1kg, 3kg","priceInfo":null,"description":"মিরপুরে বিড়ালের খাবারের বিশেষায়িত দোকান।","verified":"Verified","rating":4.5,"reviewCount":156,"lastUpdated":"2026-09-27T10:33:03.104Z","createdAt":"2026-09-27T10:33:03.104Z","updatedAt":"2026-09-27T10:33:03.104Z"},{"id":"cmujok1a90008oeciecou4gd0","shopName":"Chattogram Pet Food","ownerName":"Rashed Khan","shopImage":null,"address":"Agrabad, Chattogram","division":"Chattogram","district":"Chattogram","area":"Agrabad","phone":"+8801999999999","whatsapp":"+8801999999999","facebook":null,"website":null,"mapUrl":null,"openingHours":"9am–8pm","homeDelivery":false,"availableBrands":"Royal Canin, Whiskas, Me-O","foodTypes":"Dry, Wet, Adult, Kitten","packSizes":null,"priceInfo":null,"description":"চট্টগ্রামে বিড়ালের খাবার পাওয়ার নির্ভরযোগ্য দোকান।","verified":"Pending","rating":4.3,"reviewCount":78,"lastUpdated":"2026-09-27T10:33:03.106Z","createdAt":"2026-09-27T10:33:03.106Z","updatedAt":"2026-09-27T10:33:03.106Z"},{"id":"cmujok1aa0009oeciqqosc8wr","shopName":"Sylhet Pet Store","ownerName":"Habibur Rahman","shopImage":null,"address":"Tilagarh, Sylhet","division":"Sylhet","district":"Sylhet","area":"Tilagarh","phone":"+8801500000001","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"10am–7pm","homeDelivery":true,"availableBrands":"Whiskas, Me-O, Purina","foodTypes":"Dry, Kitten, Adult","packSizes":null,"priceInfo":null,"description":"সিলেটে বিড়ালের খাবার ও সামগ্রী পাওয়া যায়।","verified":"Not Verified","rating":4.0,"reviewCount":45,"lastUpdated":"2026-09-27T10:33:03.107Z","createdAt":"2026-09-27T10:33:03.107Z","updatedAt":"2026-09-27T10:33:03.107Z"},{"id":"cmujok1ab000aoeci7z1i9c2a","shopName":"Khulna Pet Food Corner","ownerName":"Mizanur Rahman","shopImage":null,"address":"New Market, Khulna","division":"Khulna","district":"Khulna","area":"New Market","phone":"+8801500000002","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"9am–8pm","homeDelivery":true,"availableBrands":"Royal Canin, Whiskas, SmartHeart","foodTypes":"Dry, Wet, Senior","packSizes":null,"priceInfo":null,"description":"খুলনায় বিড়ালের খাবারের সেরা দোকান।","verified":"Verified","rating":4.6,"reviewCount":89,"lastUpdated":"2026-09-27T10:33:03.108Z","createdAt":"2026-09-27T10:33:03.108Z","updatedAt":"2026-09-27T10:33:03.108Z"},{"id":"cmujok1ac000boeciry3xum11","shopName":"Uttara Pet World","ownerName":"Tanvir Ahmed","shopImage":null,"address":"Sector 7, Uttara, Dhaka","division":"Dhaka","district":"Dhaka","area":"Uttara","phone":"+8801500000003","whatsapp":"+8801500000003","facebook":"uttarapetworld","website":null,"mapUrl":null,"openingHours":"10am–9pm","homeDelivery":true,"availableBrands":"Royal Canin, Whiskas, Me-O, Purina, SmartHeart, Friskies","foodTypes":"Dry, Wet, Kitten, Adult, Senior, Special Diet","packSizes":"400g, 1kg, 2kg, 5kg, 7kg","priceInfo":"Whiskas 1.2kg ৳720, Royal Canin Persian 2kg ৳2400, Me-O 1kg ৳650","description":"উত্তরায় সব ব্র্যান্ডের বিড়ালের খাবার সহ সামগ্রী পাওয়া যায়। হোম ডেলিভারি সুবিধা।","verified":"Verified","rating":4.9,"reviewCount":312,"lastUpdated":"2026-09-27T10:33:03.109Z","createdAt":"2026-09-27T10:33:03.109Z","updatedAt":"2026-09-27T10:33:03.109Z"}],
  VeterinaryClinic: [{"id":"cmujok1af000coeciynye09d9","clinicName":"Banalata Animal Hospital","doctorName":"Dr. Anisur Rahman","specialization":"General Surgery","address":"Road 2, Dhanmondi, Dhaka","division":"Dhaka","district":"Dhaka","area":"Dhanmondi","phone":"+8809613000001","emergencyPhone":"+8809613000001","whatsapp":"+8809613000001","openingHours":"9am–9pm","emergencyService":true,"mapUrl":null,"website":null,"facebook":null,"services":"Vaccination, Consultation, Surgery, Diagnostic, Pharmacy","description":"ঢাকার অন্যতম পুরোনো ও বিশ্বস্ত পশু হাসপাতাল। ২৪/৭ ইমারজেন্সি সেবা।","verified":"Verified","rating":4.8,"reviewCount":234,"lastUpdated":"2026-09-27T10:33:03.111Z","createdAt":"2026-09-27T10:33:03.111Z","updatedAt":"2026-09-27T10:33:03.111Z","clinicImage":null},{"id":"cmujok1af000doeci9eawoapp","clinicName":"Paws & Claws Vet Clinic","doctorName":"Dr. Farhana Yasmin","specialization":"Dermatology & Nutrition","address":"Gulshan 1, Dhaka","division":"Dhaka","district":"Dhaka","area":"Gulshan","phone":"+8809613000002","emergencyPhone":null,"whatsapp":"+8809613000002","openingHours":"10am–8pm","emergencyService":true,"mapUrl":null,"website":null,"facebook":null,"services":"Consultation, Dermatology, Nutrition, Vaccination","description":"গুলশানে আধুনিক ভেট ক্লিনিক। বিশেষজ্ঞ চিকিৎসক।","verified":"Verified","rating":4.7,"reviewCount":187,"lastUpdated":"2026-09-27T10:33:03.112Z","createdAt":"2026-09-27T10:33:03.112Z","updatedAt":"2026-09-27T10:33:03.112Z","clinicImage":null},{"id":"cmujok1ah000eoeci4ys8rnw0","clinicName":"City Pet Care","doctorName":"Dr. Mahmudul Hasan","specialization":"Emergency & ICU","address":"Banani, Dhaka","division":"Dhaka","district":"Dhaka","area":"Banani","phone":"+8809613000003","emergencyPhone":"+8809613000003","whatsapp":null,"openingHours":"24/7 Emergency","emergencyService":true,"mapUrl":null,"website":null,"facebook":null,"services":"Emergency, Surgery, ICU, Imaging, Vaccination","description":"২৪/৭ ইমারজেন্সি পশু হাসপাতাল। আইসিইউ ও ইমেজিং সুবিধা।","verified":"Verified","rating":4.9,"reviewCount":312,"lastUpdated":"2026-09-27T10:33:03.113Z","createdAt":"2026-09-27T10:33:03.113Z","updatedAt":"2026-09-27T10:33:03.113Z","clinicImage":null},{"id":"cmujok1aj000foeciag810jge","clinicName":"Bay Vet Hospital","doctorName":"Dr. Sharmin Sultana","specialization":"Spay/Neuter","address":"Khulshi, Chattogram","division":"Chattogram","district":"Chattogram","area":"Khulshi","phone":"+8809613000004","emergencyPhone":null,"whatsapp":null,"openingHours":"9am–9pm","emergencyService":true,"mapUrl":null,"website":null,"facebook":null,"services":"General, Surgery, Spay/Neuter, Vaccination","description":"চট্টগ্রামের নির্ভরযোগ্য ভেট হাসপাতাল।","verified":"Verified","rating":4.6,"reviewCount":145,"lastUpdated":"2026-09-27T10:33:03.115Z","createdAt":"2026-09-27T10:33:03.115Z","updatedAt":"2026-09-27T10:33:03.115Z","clinicImage":null},{"id":"cmujok1ak000goeciitps72l5","clinicName":"Sylhet Pet Clinic","doctorName":"Dr. Kamrul Islam","specialization":"General Practice","address":"Zindabazar, Sylhet","division":"Sylhet","district":"Sylhet","area":"Zindabazar","phone":"+8809613000005","emergencyPhone":null,"whatsapp":null,"openingHours":"10am–8pm","emergencyService":false,"mapUrl":null,"website":null,"facebook":null,"services":"General, Vaccination, Wellness","description":"সিলেটে পশু চিকিৎসা সেবা।","verified":"Pending","rating":4.4,"reviewCount":67,"lastUpdated":"2026-09-27T10:33:03.116Z","createdAt":"2026-09-27T10:33:03.116Z","updatedAt":"2026-09-27T10:33:03.116Z","clinicImage":null},{"id":"cmujok1al000hoeciiew1ljpt","clinicName":"Rajshahi Animal Care","doctorName":"Dr. Nazmun Nahar","specialization":"Dental & Dermatology","address":"Kazla, Rajshahi","division":"Rajshahi","district":"Rajshahi","area":"Kazla","phone":"+8809613000006","emergencyPhone":null,"whatsapp":null,"openingHours":"9am–7pm","emergencyService":false,"mapUrl":null,"website":null,"facebook":null,"services":"General, Dental, Dermatology","description":"রাজশাহীতে পশু চিকিৎসা ও পরামর্শ।","verified":"Not Verified","rating":4.3,"reviewCount":45,"lastUpdated":"2026-09-27T10:33:03.117Z","createdAt":"2026-09-27T10:33:03.117Z","updatedAt":"2026-09-27T10:33:03.117Z","clinicImage":null},{"id":"cmujok1am000ioecixghjb4pn","clinicName":"Khulna Pet Hospital","doctorName":"Dr. Imran Hossain","specialization":"Surgery","address":"Boyra, Khulna","division":"Khulna","district":"Khulna","area":"Boyra","phone":"+8809613000007","emergencyPhone":null,"whatsapp":null,"openingHours":"9am–9pm","emergencyService":true,"mapUrl":null,"website":null,"facebook":null,"services":"General, Surgery, Emergency, Vaccination","description":"খুলনায় ২৪ ঘন্টা পশু চিকিৎসা সেবা।","verified":"Verified","rating":4.6,"reviewCount":102,"lastUpdated":"2026-09-27T10:33:03.118Z","createdAt":"2026-09-27T10:33:03.118Z","updatedAt":"2026-09-27T10:33:03.118Z","clinicImage":null},{"id":"cmujok1an000joeciy498mbgh","clinicName":"Mirpur Vet Center","doctorName":"Dr. Rezaul Karim","specialization":"General & Vaccination","address":"Mirpur 10, Dhaka","division":"Dhaka","district":"Dhaka","area":"Mirpur","phone":"+8809613000008","emergencyPhone":null,"whatsapp":"+8809613000008","openingHours":"9am–9pm","emergencyService":true,"mapUrl":null,"website":null,"facebook":null,"services":"Vaccination, Consultation, Surgery, Diagnostic","description":"মিরপুরে সাশ্রয়ী পশু চিকিৎসা সেবা। টিকাকরণ ও সার্জারি।","verified":"Verified","rating":4.5,"reviewCount":178,"lastUpdated":"2026-09-27T10:33:03.119Z","createdAt":"2026-09-27T10:33:03.119Z","updatedAt":"2026-09-27T10:33:03.119Z","clinicImage":null}],
  ServiceProvider: [{"id":"cmujok1ao000koeci1eqs45cm","providerName":"Purrfect Paws Spa","businessName":null,"type":"Grooming","image":null,"address":"Dhanmondi, Dhaka","division":"Dhaka","district":"Dhaka","area":"Dhanmondi","phone":"+8801700000010","whatsapp":"+8801700000010","facebook":null,"website":null,"mapUrl":null,"openingHours":"10am–7pm","homeService":true,"description":"বিড়ালের গ্রুমিং সেবা — বাথ, ব্রাশ, নেইল ট্রিম। হোম সার্ভিস উপলব্ধ।","verified":"Verified","rating":4.8,"reviewCount":156,"lastUpdated":"2026-09-27T10:33:03.120Z","createdAt":"2026-09-27T10:33:03.120Z","updatedAt":"2026-09-27T10:33:03.120Z"},{"id":"cmujok1ap000loecii8xcrv9p","providerName":"Cat Cottage Dhaka","businessName":null,"type":"Boarding","image":null,"address":"Banani, Dhaka","division":"Dhaka","district":"Dhaka","area":"Banani","phone":"+8801700000011","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"24/7","homeService":false,"description":"বিড়ালের বোর্ডিং ও ডে কেয়ার। নিরাপদ ও আরামদায়ক পরিবেশ।","verified":"Verified","rating":4.7,"reviewCount":89,"lastUpdated":"2026-09-27T10:33:03.122Z","createdAt":"2026-09-27T10:33:03.122Z","updatedAt":"2026-09-27T10:33:03.122Z"},{"id":"cmujok1aq000moecigdhgbmgp","providerName":"Whisker Lens","businessName":null,"type":"Photography","image":null,"address":"Gulshan, Dhaka","division":"Dhaka","district":"Dhaka","area":"Gulshan","phone":"+8801700000012","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"By appointment","homeService":true,"description":"বিড়ালের প্রফেশনাল ফটোগ্রাফি। হোম ও স্টুডিও সেবা।","verified":"Verified","rating":4.9,"reviewCount":45,"lastUpdated":"2026-09-27T10:33:03.123Z","createdAt":"2026-09-27T10:33:03.123Z","updatedAt":"2026-09-27T10:33:03.123Z"},{"id":"cmujok1ar000noecizi03c4p3","providerName":"Care n Cure Pharmacy","businessName":null,"type":"Pharmacy","image":null,"address":"Dhanmondi, Dhaka","division":"Dhaka","district":"Dhaka","area":"Dhanmondi","phone":"+8801700000013","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"9am–10pm","homeService":false,"description":"বিড়ালের ওষুধ ও স্বাস্থ্য সামগ্রীর বিশেষায়িত ফার্মেসি।","verified":"Verified","rating":4.6,"reviewCount":134,"lastUpdated":"2026-09-27T10:33:03.124Z","createdAt":"2026-09-27T10:33:03.124Z","updatedAt":"2026-09-27T10:33:03.124Z"},{"id":"cmujok1as000ooecifd9r980h","providerName":"PAWSS Bangladesh","businessName":null,"type":"Rescue","image":null,"address":"Mirpur, Dhaka","division":"Dhaka","district":"Dhaka","area":"Mirpur","phone":"+8801700000014","whatsapp":"+8801700000014","facebook":"pawssbd","website":null,"mapUrl":null,"openingHours":"24/7 Helpline","homeService":true,"description":"বিড়াল উদ্ধার সংস্থা। আহত ও বেওয়ারিশ বিড়াল উদ্ধার, চিকিৎসা ও পুনর্বাসন।","verified":"Verified","rating":4.9,"reviewCount":267,"lastUpdated":"2026-09-27T10:33:03.125Z","createdAt":"2026-09-27T10:33:03.125Z","updatedAt":"2026-09-27T10:33:03.125Z"},{"id":"cmujok1au000poeciyyowx1oo","providerName":"PawPort BD","businessName":null,"type":"Transport","image":null,"address":"Dhaka","division":"Dhaka","district":"Dhaka","area":"All areas","phone":"+8801700000015","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"7am–10pm","homeService":true,"description":"বিড়ালের নিরাপদ পরিবহন সেবা। ভেট, গ্রুমিং ও যেকোনো স্থানে।","verified":"Pending","rating":4.4,"reviewCount":78,"lastUpdated":"2026-09-27T10:33:03.127Z","createdAt":"2026-09-27T10:33:03.127Z","updatedAt":"2026-09-27T10:33:03.127Z"},{"id":"cmujok1av000qoecikwcgbb2k","providerName":"Chattogram Cat Grooming","businessName":null,"type":"Grooming","image":null,"address":"Khulshi, Chattogram","division":"Chattogram","district":"Chattogram","area":"Khulshi","phone":"+8801700000016","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"10am–7pm","homeService":true,"description":"চট্টগ্রামে বিড়ালের গ্রুমিং সেবা।","verified":"Not Verified","rating":4.3,"reviewCount":56,"lastUpdated":"2026-09-27T10:33:03.128Z","createdAt":"2026-09-27T10:33:03.128Z","updatedAt":"2026-09-27T10:33:03.128Z"},{"id":"cmujok1aw000roecijovedt44","providerName":"Sylhet Boarding Home","businessName":null,"type":"Boarding","image":null,"address":"Zindabazar, Sylhet","division":"Sylhet","district":"Sylhet","area":"Zindabazar","phone":"+8801700000017","whatsapp":null,"facebook":null,"website":null,"mapUrl":null,"openingHours":"24/7","homeService":false,"description":"সিলেটে বিড়ালের বোর্ডিং সুবিধা।","verified":"Pending","rating":4.2,"reviewCount":34,"lastUpdated":"2026-09-27T10:33:03.129Z","createdAt":"2026-09-27T10:33:03.129Z","updatedAt":"2026-09-27T10:33:03.129Z"}],
  Article: [{"id":"cmuj8h3bq000roe50s4o6r55t","title":"The Complete Guide to Feeding Your Persian Cat in Bangladesh","category":"Nutrition","excerpt":"Persian cats need a specific diet to keep their coat glossy and digestion smooth. Here is everything Bangladeshi cat parents should know.","content":"Persian cats are prone to hairballs and have a unique jaw shape that makes eating difficult with standard kibble. Always choose breed-specific food like Royal Canin Persian. Provide fresh water daily — a ceramic fountain encourages drinking. Avoid raw fish and milk; many cats are lactose intolerant. Stick to measured portions twice daily to prevent obesity, which Persians are prone to.","author":"Dr. Farhana Yasmin","readTime":7,"image":null,"tags":"persian,diet,nutrition","featured":true,"createdAt":"2026-09-27T03:02:51.926Z","updatedAt":"2026-09-27T03:02:51.926Z"},{"id":"cmuj8h3bq000soe508ekphl4e","title":"Monsoon Care: Keeping Your Cat Safe During Bangladesh's Rainy Season","category":"Health","excerpt":"Damp weather brings fleas, fungal infections and chill. Here is how to keep your cat cosy and healthy through the barsa.","content":"During the monsoon, keep your cat indoors, dry and warm. Use a dehumidifier or air-conditioner to control moisture. Check ears weekly for fungal infections. Apply a flea-and-tick preventative like Seresto. Brush regularly to prevent matting. If your cat gets wet, towel-dry thoroughly — never use a human hair dryer on hot. Watch for sneezing or watery eyes, common signs of cat flu in humid weather.","author":"Dr. Mahmudul Hasan","readTime":6,"image":null,"tags":"monsoon,health,seasonal","featured":true,"createdAt":"2026-09-27T03:02:51.927Z","updatedAt":"2026-09-27T03:02:51.927Z"},{"id":"cmuj8h3br000toe50zb1vfbqh","title":"Litter Training Your Kitten in 5 Simple Steps","category":"Behavior","excerpt":"New kitten parent? Litter training is easier than you think. Follow this proven five-step method used by rescues across Dhaka.","content":"1) Choose a low-sided litter box kittens can climb into. 2) Use unscented clumping litter — kittens have sensitive noses. 3) Place the kitten in the box after meals and naps. 4) Gently scratch their front paws in the litter. 5) Praise and reward with a treat. Never punish accidents — clean them with an enzymatic cleaner. Most kittens are reliably trained within two weeks.","author":"BiralBond Editorial","readTime":5,"image":null,"tags":"kitten,training,litter","featured":false,"createdAt":"2026-09-27T03:02:51.928Z","updatedAt":"2026-09-27T03:02:51.928Z"},{"id":"cmuj8h3bs000uoe50g87g0x1d","title":"Grooming Your Long-Haired Cat: A Step-by-Step Routine","category":"Grooming","excerpt":"Long-haired beauties need daily attention. Build a stress-free grooming routine that you and your cat will both enjoy.","content":"Long-haired breeds like Persians, Maine Coons and Ragdolls need daily brushing to prevent painful mats. Use a wide stainless steel comb first, then a slicker brush. Bath every 4–6 weeks with a cat-specific shampoo. Check eyes daily — flat-faced breeds tear up. Trim nails every two weeks. Start grooming as kittens with short, treat-filled sessions so they learn to love it.","author":"Sadia Islam","readTime":8,"image":null,"tags":"grooming,long-hair,routine","featured":true,"createdAt":"2026-09-27T03:02:51.928Z","updatedAt":"2026-09-27T03:02:51.928Z"},{"id":"cmuj8h3bs000voe503zlctts6","title":"Vaccination Schedule Every Bangladeshi Cat Parent Must Follow","category":"Health","excerpt":"From FVRCP to rabies, here is the complete vaccination calendar recommended by Bangladeshi veterinarians.","content":"Core vaccines for cats in Bangladesh: FVRCP (feline distemper) at 8, 12 and 16 weeks, then annually. Rabies at 12–16 weeks, then annually. Optional but recommended: FeLV (feline leukaemia) for outdoor cats. Keep a vaccination card and never miss a booster. Vaccines cost ৳800–৳1,500 per dose and are available at all major vet clinics in Dhaka, Chattogram, Sylhet and Khulna.","author":"Dr. Anisur Rahman","readTime":6,"image":null,"tags":"vaccine,health,schedule","featured":false,"createdAt":"2026-09-27T03:02:51.929Z","updatedAt":"2026-09-27T03:02:51.929Z"},{"id":"cmuj8h3bt000woe5070rjv5mq","title":"Caring for Your Senior Cat (10+ Years)","category":"Senior","excerpt":"Older cats need gentler routines, softer food and more frequent vet visits. Here is how to make their golden years truly golden.","content":"Senior cats sleep more, eat less and may develop arthritis or kidney issues. Switch to senior-formula food, easily warmed. Add ramps to favourite spots so they don't jump. Provide a low-sided litter box. Schedule vet check-ups every 6 months with bloodwork. Watch for changes in drinking, appetite or litter box habits — early detection saves lives.","author":"Dr. Nazmun Nahar","readTime":7,"image":null,"tags":"senior,health,care","featured":false,"createdAt":"2026-09-27T03:02:51.930Z","updatedAt":"2026-09-27T03:02:51.930Z"}],
  CommunityPost: [{"id":"cmuj8h3bu000xoe5004agie4i","author":"Tahmid Karim","avatar":"TK","title":"My rescue kitten just discovered the mirror!","body":"After three weeks of hiding, Mochi finally noticed herself in the wardrobe mirror. The slow-blink fest that followed melted my heart. Sharing a photo below.","category":"Story","city":"Dhaka","likes":142,"comments":23,"image":null,"createdAt":"2026-09-27T03:02:51.930Z","updatedAt":"2026-09-27T03:02:51.930Z"},{"id":"cmuj8h3bv000yoe5009xand3s","author":"Labiba Rahman","avatar":"LR","title":"Best wet food brand for a picky Persian?","body":"My 3-year-old Persian Bagha has suddenly gone off his usual brand. Has anyone tried Sheba or Felix pouches available in Dhaka? Looking for recommendations that won't break the bank.","category":"Question","city":"Dhaka","likes":38,"comments":41,"image":null,"createdAt":"2026-09-27T03:02:51.931Z","updatedAt":"2026-09-27T03:02:51.931Z"},{"id":"cmuj8h3bv000zoe50sh8r7s32","author":"Sajid Ahmed","avatar":"SA","title":"Organising a cat meet-up in Chattogram this month","body":"Calling all Chattogram cat parents! Planning a casual meet-up at Polo Ground on the last Saturday of the month. Bring your kitties (in carriers) and let's swap stories. RSVP below.","category":"Discussion","city":"Chattogram","likes":87,"comments":56,"image":null,"createdAt":"2026-09-27T03:02:51.932Z","updatedAt":"2026-09-27T03:02:51.932Z"},{"id":"cmuj8h3bw0010oe50scdk1s28","author":"Maliha Chowdhury","avatar":"MC","title":"Tofu litter changed my life (and my nose)","body":"Switched from clay to tofu litter two weeks ago. Zero dust, zero smell, and it flushes straight down the toilet. My Bengal's asthma has improved dramatically. Highly recommend.","category":"Advice","city":"Sylhet","likes":213,"comments":34,"image":null,"createdAt":"2026-09-27T03:02:51.932Z","updatedAt":"2026-09-27T03:02:51.932Z"},{"id":"cmuj8h3bw0011oe50o77awvhj","author":"Rakib Hasan","avatar":"RH","title":"Photo dump: my tabby's first Eid","body":"Dressed Tuntuni in a tiny panjabi and she was NOT amused. Sharing the photographic evidence of her dignified suffering.","category":"Photo","city":"Khulna","likes":301,"comments":67,"image":null,"createdAt":"2026-09-27T03:02:51.933Z","updatedAt":"2026-09-27T03:02:51.933Z"}],
  LostFound: [{"id":"cmuj8h3bx0012oe505ozbexd7","type":"Lost","catName":"Buri","breed":"Local (Tabby)","color":"Brown tabby with white paws","location":"Road 7A, Dhanmondi","city":"Dhaka","date":"2025-01-12","contact":"Sumaiya Akter","phone":"+88017XXXXXXXX","description":"Buri slipped out through the gate during evening walks. She is shy but responds to \"Buri Buri\". Wearing a pink collar with a bell.","reward":"৳2,000","image":null,"status":"Active","createdAt":"2026-09-27T03:02:51.934Z","updatedAt":"2026-09-27T03:02:51.934Z"},{"id":"cmuj8h3by0013oe505blbk3hv","type":"Found","catName":null,"breed":"Local (Black)","color":"All black, yellow eyes","location":"Block C, Mirpur","city":"Dhaka","date":"2025-01-15","contact":"Arif Hossain","phone":"+88018XXXXXXXX","description":"Friendly black cat found crying near the market. Not microchipped, no collar. Currently being fostered. Looking for the owner.","reward":null,"image":null,"status":"Active","createdAt":"2026-09-27T03:02:51.934Z","updatedAt":"2026-09-27T03:02:51.934Z"},{"id":"cmuj8h3bz0014oe509uekk0el","type":"Lost","catName":"Tom","breed":"Persian","color":"White with grey patches","location":"GEC Circle, Chattogram","city":"Chattogram","date":"2025-01-10","contact":"Nadia Islam","phone":"+88019XXXXXXXX","description":"Tom is an indoor cat who escaped when guests left the door open. He is microchipped and neutered. Very friendly.","reward":"৳5,000","image":null,"status":"Active","createdAt":"2026-09-27T03:02:51.936Z","updatedAt":"2026-09-27T03:02:51.936Z"},{"id":"cmuj8h3c00015oe5097rub3tf","type":"Found","catName":null,"breed":"Local (Orange)","color":"Orange tabby, white chest","location":"Zindabazar","city":"Sylhet","date":"2025-01-14","contact":"Imran Ali","phone":"+88015XXXXXXXX","description":"Sweet orange boy found near the tea stall. Looks well-fed, probably someone's pet. Keeping him safe for now.","reward":null,"image":null,"status":"Reunited","createdAt":"2026-09-27T03:02:51.937Z","updatedAt":"2026-09-27T11:43:47.790Z"},{"id":"cmumf7qg80000r1kre3gq8xy8","type":"Lost","catName":"Test Cat","breed":"Local","color":"Orange tabby","location":"Dhanmondi","city":"Dhaka","date":"2025-09-28","contact":"Test Owner","phone":"+8801700000000","description":"এই বিড়ালটি হারিয়ে গেছে। দয়া করে যদি কেউ দেখেন যোগাযোগ করুন।","reward":null,"image":"/uploads/upload-1790679395736-kfg6ep.png","status":"Active","createdAt":"2026-09-29T08:34:51.176Z","updatedAt":"2026-09-29T10:56:40.047Z"}],
  Event: [{"id":"cmuj8h3c10016oe50fz67mhrz","title":"Dhaka Cat Show 2025 — Winter Edition","type":"Cat Show","date":"2025-02-15","time":"10:00 AM – 6:00 PM","venue":"Bangabandhu International Conference Centre","city":"Dhaka","description":"The biggest cat show of the year! 300+ cats competing across 12 categories. Bring your feline royalty or just come admire. Stalls, vets and free grooming demos.","capacity":500,"registered":347,"image":null,"createdAt":"2026-09-27T03:02:51.937Z","updatedAt":"2026-09-27T03:02:51.937Z"},{"id":"cmuj8h3c20017oe506e9gpz6z","title":"Kitten Care Workshop (Beginner)","type":"Workshop","date":"2025-02-08","time":"3:00 PM – 5:00 PM","venue":"Paws & Claws Vet Clinic, Gulshan","city":"Dhaka","description":"A hands-on workshop for new kitten parents. Learn feeding, litter training, grooming basics and what to keep in your home first-aid kit.","capacity":30,"registered":22,"image":null,"createdAt":"2026-09-27T03:02:51.938Z","updatedAt":"2026-09-27T03:02:51.938Z"},{"id":"cmuj8h3c20018oe50wdzps8gp","title":"Chattogram Cat Parents Meet-up","type":"Meetup","date":"2025-02-22","time":"4:00 PM – 6:00 PM","venue":"Polo Ground","city":"Chattogram","description":"A relaxed afternoon for Chattogram cat parents. Bring your kitties (in carriers) for photos, snacks and stories. Free entry.","capacity":80,"registered":41,"image":null,"createdAt":"2026-09-27T03:02:51.939Z","updatedAt":"2026-09-27T03:02:51.939Z"},{"id":"cmuj8h3c40019oe507sfjy782","title":"Adoption Drive — Rescue Kittens Need Homes","type":"Adoption Drive","date":"2025-02-11","time":"11:00 AM – 4:00 PM","venue":"Banalata Animal Hospital, Dhanmondi","city":"Dhaka","description":"20+ vaccinated, healthy rescue kittens looking for forever homes. Adoption fees waived. Bring a carrier and your heart.","capacity":100,"registered":58,"image":null,"createdAt":"2026-09-27T03:02:51.940Z","updatedAt":"2026-09-27T03:02:51.940Z"},{"id":"cmuj8h3c4001aoe502kexguoa","title":"Webinar: Feline Nutrition Demystified","type":"Webinar","date":"2025-02-18","time":"8:00 PM – 9:30 PM","venue":"Online (Zoom)","city":"Online","description":"Join Dr. Farhana Yasmin for a deep dive into cat nutrition — wet vs dry, raw diets, and what labels really mean. Q&A at the end.","capacity":200,"registered":134,"image":null,"createdAt":"2026-09-27T03:02:51.941Z","updatedAt":"2026-09-27T03:02:51.941Z"}],
  Gallery: [{"id":"cmuj8h3c5001boe50460v0emq","title":"Sunbeam snooze","cat":"Mochi","owner":"Tahmid K.","city":"Dhaka","image":"https://images.unsplash.com/photo-1518791841217-8f162f1e1131?w=800&q=80","likes":234,"caption":"Golden hour, golden boy.","createdAt":"2026-09-27T03:02:51.942Z","updatedAt":"2026-09-27T03:02:51.942Z"},{"id":"cmuj8h3c6001coe50dz9tqlbt","title":"The judge","cat":"Bagha","owner":"Nusrat J.","city":"Dhaka","image":"/cats/bagha.png","likes":189,"caption":null,"createdAt":"2026-09-27T03:02:51.942Z","updatedAt":"2026-09-27T03:18:39.417Z"},{"id":"cmuj8h3c6001doe50uioqeem9","title":"Window watcher","cat":"Tuntuni","owner":"Rescue","city":"Chattogram","image":"https://images.unsplash.com/photo-1606214174585-fe31582dc6ee?w=800&q=80","likes":312,"caption":null,"createdAt":"2026-09-27T03:02:51.943Z","updatedAt":"2026-09-27T03:02:51.943Z"},{"id":"cmuj8h3c6001eoe50ssqikiuy","title":"Yoga master","cat":"Mia","owner":"Tanvir H.","city":"Rajshahi","image":"https://images.unsplash.com/photo-1561948955-570b270e7c36?w=800&q=80","likes":167,"caption":null,"createdAt":"2026-09-27T03:02:51.943Z","updatedAt":"2026-09-27T03:02:51.943Z"},{"id":"cmuj8h3c7001foe50msqrz528","title":"Twin trouble","cat":"Kalo & Milk","owner":"Khulna PP","city":"Khulna","image":"https://images.unsplash.com/photo-1495360010541-f48722b34f7d?w=800&q=80","likes":401,"caption":null,"createdAt":"2026-09-27T03:02:51.944Z","updatedAt":"2026-09-27T03:02:51.944Z"},{"id":"cmuj8h3c8001goe50ed4ss6qk","title":"Royal repose","cat":"Simba","owner":"Sadia I.","city":"Sylhet","image":"https://images.unsplash.com/photo-1574158622682-e40e69881006?w=800&q=80","likes":278,"caption":null,"createdAt":"2026-09-27T03:02:51.944Z","updatedAt":"2026-09-27T03:02:51.944Z"}],
  Member: [{"id":"cmuj8tf9q0000oewfboqpopws","name":"Rifat Ahmed","email":"rifat@example.com","phone":"+8801712345678","city":"Dhaka","plan":"Gold","cats":1,"joinedAt":"2026-09-27T03:12:27.278Z","createdAt":"2026-09-27T03:12:27.278Z","updatedAt":"2026-09-27T03:12:27.278Z"}],
  Newsletter: [{"id":"cmuj8tqyh0001oewfk1d8wya4","email":"test@biralbond.bd","createdAt":"2026-09-27T03:12:42.426Z"}],
  Testimonial: [{"id":"cmuj93d2d0000oea7fxda1gyw","name":"Tahmid Karim","role":"Cat parent of 2","city":"Dhaka","avatar":"TK","rating":5,"quote":"I adopted my first rescue through BiralBond and the experience was seamless. The team guided me through every step — from choosing the right kitten to the first vet visit. Mochi is now the queen of our house!","catName":"Mochi","featured":true,"createdAt":"2026-09-27T03:20:10.982Z","updatedAt":"2026-09-27T03:20:10.982Z"},{"id":"cmuj93d2f0001oea7du1rd5sl","name":"Dr. Farhana Yasmin","role":"Veterinarian","city":"Dhaka","avatar":"FY","rating":5,"quote":"As a vet, I trust BiralBond's care guides more than any other Bangladeshi resource. They're accurate, practical and genuinely put cats first. I recommend the platform to every new cat parent I meet.","catName":null,"featured":true,"createdAt":"2026-09-27T03:20:10.983Z","updatedAt":"2026-09-27T03:20:10.983Z"},{"id":"cmuj93d2g0002oea7rigzcfpw","name":"Maliha Chowdhury","role":"Bengal cat owner","city":"Sylhet","avatar":"MC","rating":5,"quote":"The product range is incredible — finally a place in Bangladesh where I can get breed-specific food and organic litter without worrying about fakes. Delivery to Sylhet took just 2 days!","catName":"Leo","featured":true,"createdAt":"2026-09-27T03:20:10.984Z","updatedAt":"2026-09-27T03:20:10.984Z"},{"id":"cmuj93d2g0003oea795c3qtay","name":"Sajid Ahmed","role":"Community organiser","city":"Chattogram","avatar":"SA","rating":5,"quote":"I organised our first Chattogram meet-up through the community forum and 40 cat parents showed up! BiralBond has built something special — a real, warm community, not just a marketplace.","catName":null,"featured":false,"createdAt":"2026-09-27T03:20:10.985Z","updatedAt":"2026-09-27T03:20:10.985Z"},{"id":"cmuj93d2h0004oea7d5ppeis9","name":"Nadia Islam","role":"Persian cat parent","city":"Chattogram","avatar":"NI","rating":5,"quote":"When my Tom went missing, I posted on the Lost & Found board and within 6 hours a neighbour three streets away recognised him. BiralBond reunited us. I will be forever grateful.","catName":"Tom","featured":true,"createdAt":"2026-09-27T03:20:10.986Z","updatedAt":"2026-09-27T03:20:10.986Z"},{"id":"cmuj93d2i0005oea7rlblui5l","name":"Rakib Hasan","role":"First-time cat dad","city":"Khulna","avatar":"RH","rating":5,"quote":"The care articles taught me everything I needed to know as a first-time cat parent — from litter training to vaccination schedules. Written in simple language, very practical for Bangladesh.","catName":"Tuntuni","featured":false,"createdAt":"2026-09-27T03:20:10.986Z","updatedAt":"2026-09-27T03:20:10.986Z"},{"id":"cmuj93d2j0006oea7e0plvucx","name":"Sumaiya Akter","role":"Rescue volunteer","city":"Dhaka","avatar":"SA","rating":5,"quote":"Our rescue has placed over 50 kittens through BiralBond's adoption listings. The adoption-first philosophy is real — they genuinely care about matching the right cat to the right home.","catName":null,"featured":false,"createdAt":"2026-09-27T03:20:10.987Z","updatedAt":"2026-09-27T03:20:10.987Z"},{"id":"cmuj93d2j0007oea7ldt5hns1","name":"Imran Ali","role":"Scottish Fold owner","city":"Sylhet","avatar":"IA","rating":4,"quote":"Booked a grooming session through the Services directory and was blown away. The provider came home, was gentle with my anxious cat, and sent before/after photos. Premium service indeed.","catName":"Bagha","featured":false,"createdAt":"2026-09-27T03:20:10.988Z","updatedAt":"2026-09-27T03:20:10.988Z"}],
  CareTip: [{"id":"cmuj93d2k0008oea7tffnyw7h","title":"Fresh water, always","body":"Cats are notoriously bad drinkers. Place 2–3 water bowls around the house, away from food. A ceramic fountain increases intake by up to 30%.","icon":"💧","category":"Daily","createdAt":"2026-09-27T03:20:10.988Z","updatedAt":"2026-09-27T03:20:10.988Z"},{"id":"cmuj93d2k0009oea7mo02whc2","title":"Brush before meals","body":"Link grooming to mealtime. Your cat will associate brushing with something positive and mats stay away.","icon":"🪮","category":"Grooming","createdAt":"2026-09-27T03:20:10.989Z","updatedAt":"2026-09-27T03:20:10.989Z"},{"id":"cmuj93d2l000aoea7zu7oxf8b","title":"Slow blink = I love you","body":"When your cat slow-blinks at you, blink back slowly. It's the feline \"I trust you\" signal.","icon":"👁️","category":"Behavior","createdAt":"2026-09-27T03:20:10.989Z","updatedAt":"2026-09-27T03:20:10.989Z"},{"id":"cmuj93d2l000boea75rvvsbyw","title":"One litter box per cat + one","body":"Two cats? Three boxes. Place them in quiet, separate spots. Scoop daily, full change weekly.","icon":"🚽","category":"Daily","createdAt":"2026-09-27T03:20:10.990Z","updatedAt":"2026-09-27T03:20:10.990Z"},{"id":"cmuj93d2m000coea7bu3j6ya0","title":"Vertical space matters","body":"Cats feel safe up high. A cat tree, wall shelves or a cleared bookshelf top reduces stress dramatically in small Dhaka apartments.","icon":"📈","category":"Environment","createdAt":"2026-09-27T03:20:10.991Z","updatedAt":"2026-09-27T03:20:10.991Z"},{"id":"cmuj93d2n000doea7uky3dzwo","title":"Play before meals","body":"Hunt → catch → eat → groom → sleep. A 15-minute play session before food mimics this natural cycle and prevents midnight zoomies.","icon":"🪶","category":"Behavior","createdAt":"2026-09-27T03:20:10.991Z","updatedAt":"2026-09-27T03:20:10.991Z"}],
  Comment: [{"id":"cmuj9q6rl0000oe1vt0huax13","postId":"cmuj8h3bu000xoe5004agie4i","author":"Rifat A.","avatar":"RI","body":"This is so wholesome! Thanks for sharing. 🐱","createdAt":"2026-09-27T03:37:55.905Z"},{"id":"cmuj9q6rl0001oe1vl3g3lgjh","postId":"cmuj8h3bv000yoe5009xand3s","author":"Sadia I.","avatar":"SA","body":"Aww, that's adorable. Give Mochi a chin scratch from me!","createdAt":"2026-09-27T03:37:55.906Z"},{"id":"cmuj9q6rm0002oe1vzi12byqt","postId":"cmuj8h3bv000yoe5009xand3s","author":"Nusrat J.","avatar":"NU","body":"Have you tried Sheba pouches? My picky Persian loves them.","createdAt":"2026-09-27T03:37:55.906Z"},{"id":"cmuj9q6rm0003oe1vpyy7mmcm","postId":"cmuj8h3bv000yoe5009xand3s","author":"Labiba R.","avatar":"LA","body":"Great tip, switching to tofu litter this week.","createdAt":"2026-09-27T03:37:55.907Z"},{"id":"cmuj9q6rn0004oe1v0hq072fq","postId":"cmuj8h3bv000zoe50sh8r7s32","author":"Imran A.","avatar":"IM","body":"Count me in! Bringing my two tabbies.","createdAt":"2026-09-27T03:37:55.907Z"},{"id":"cmuj9q6ro0005oe1v6zwcyg72","postId":"cmuj8h3bv000zoe50sh8r7s32","author":"Sajid A.","avatar":"SA","body":"The panjabi photos are killing me 😂","createdAt":"2026-09-27T03:37:55.908Z"},{"id":"cmuj9q6ro0006oe1va7zkv822","postId":"cmuj8h3bv000zoe50sh8r7s32","author":"Rifat A.","avatar":"RI","body":"This is so wholesome! Thanks for sharing. 🐱","createdAt":"2026-09-27T03:37:55.909Z"},{"id":"cmuj9q6rp0007oe1v1dhiguu1","postId":"cmuj8h3bw0010oe50scdk1s28","author":"Maliha C.","avatar":"MA","body":"Mine did the exact same thing! So relatable.","createdAt":"2026-09-27T03:37:55.909Z"},{"id":"cmuj9q6rp0008oe1vyyq43b8u","postId":"cmuj8h3bw0010oe50scdk1s28","author":"Sadia I.","avatar":"SA","body":"Aww, that's adorable. Give Mochi a chin scratch from me!","createdAt":"2026-09-27T03:37:55.910Z"},{"id":"cmuj9q6rq0009oe1vbw4pobck","postId":"cmuj8h3bw0010oe50scdk1s28","author":"Nusrat J.","avatar":"NU","body":"Have you tried Sheba pouches? My picky Persian loves them.","createdAt":"2026-09-27T03:37:55.910Z"},{"id":"cmuj9q6rq000aoe1vkipdz4l8","postId":"cmuj8h3bw0011oe50o77awvhj","author":"Tanvir H.","avatar":"TA","body":"Same thing happened with my cat. They figure it out eventually!","createdAt":"2026-09-27T03:37:55.911Z"}],
  FeatureToggle: [{"id":"cmujok1ax000soecijj805lkz","feature":"ecommerce","enabled":false,"createdAt":"2026-09-27T10:33:03.130Z","updatedAt":"2026-09-27T10:33:03.130Z"},{"id":"cmujok1ay000toeci5doaikdh","feature":"catMarketplace","enabled":false,"createdAt":"2026-09-27T10:33:03.130Z","updatedAt":"2026-09-27T10:33:03.130Z"},{"id":"cmujok1az000uoeci431jhjm0","feature":"onlinePayment","enabled":false,"createdAt":"2026-09-27T10:33:03.132Z","updatedAt":"2026-09-27T10:33:03.132Z"},{"id":"cmujok1b0000voecimilldmow","feature":"silverMembership","enabled":false,"createdAt":"2026-09-27T10:33:03.133Z","updatedAt":"2026-09-27T10:33:03.133Z"},{"id":"cmujok1b1000woecivnyaf7jw","feature":"goldMembership","enabled":false,"createdAt":"2026-09-27T10:33:03.134Z","updatedAt":"2026-09-27T10:33:03.134Z"},{"id":"cmujok1b4000xoeciz8e2i93v","feature":"premiumMembership","enabled":false,"createdAt":"2026-09-27T10:33:03.136Z","updatedAt":"2026-09-27T10:33:03.136Z"},{"id":"cmujok1b5000yoeci79zgs4vl","feature":"freeMembership","enabled":true,"createdAt":"2026-09-27T10:33:03.138Z","updatedAt":"2026-09-27T10:33:03.138Z"}]
};
