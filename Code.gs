/**
 * OKV RMS - Google Apps Script back end
 * ------------------------------------------------------------------
 * Runs everything for OKV RMS using ONLY Google Sheets + Apps Script:
 *   - landing-page forms and in-app forms (free-trial download, activation request, contact)
 *   - storage of every submission, customer, activation code, alert and message (Google Sheet)
 *   - e-mail to customers and an alert e-mail to OKV Technology on every submission
 *   - activation codes (1,000+ ready), sent from the admin dashboard straight to the customer's mail
 *   - the OKV Technology admin API (admin.html on GitHub Pages calls this file; no HTML is hosted here)
 *
 * Install: see SETUP_GUIDE.md. Bound to the Google Sheet (Extensions > Apps Script).
 */

// MUST be identical to LICENSE_SECRET in okv_rms.py. Change BOTH together before you ship, then rebuild the app.
const LICENSE_SECRET = 'a7f3c9e1d4b6820f5c3a9e7d1b4f6a82';
const CODE_ALPHA = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

const SHEETS = {
  Submissions: ['ID', 'Timestamp', 'Type', 'Status', 'Name', 'Business', 'Business Type', 'Email', 'Phone', 'State', 'Source', 'Install ID', 'Amount', 'Pay Date', 'Pay Ref', 'Receipt URL', 'Message', 'Code Sent', 'Updated', 'Note'],
  Customers: ['Email', 'Name', 'Business', 'Business Type', 'Phone', 'State', 'Status', 'Trial Start', 'Activated On', 'Code', 'Install ID', 'Notes', 'Created', 'Updated'],
  Codes: ['Code', 'Status', 'Sent To', 'Sent On', 'Request ID', 'Bound Install', 'Used On', 'Batch'],
  Alerts: ['ID', 'Timestamp', 'Type', 'Title', 'Detail', 'Ref', 'Read'],
  Messages: ['Timestamp', 'Channel', 'To', 'Subject', 'Body', 'Status', 'Ref'],
  Payments: ['ID', 'Timestamp', 'Method', 'Reference', 'Name', 'Email', 'Phone', 'Amount', 'Currency', 'Status', 'Submission', 'Code', 'Receipt URL', 'Note', 'Updated'],
  Audit: ['Timestamp', 'Action', 'Detail'],
  Config: ['Key', 'Value']
};

const DEFAULT_CONFIG = {
  admin_email: 'technologyokv@gmail.com',
  from_name: 'OKV Technology',
  product_name: 'OKV RMS',
  price: '120000',
  trial_days: '12',
  currency: '\u20a6',
  whatsapp: '2348104141138',
  phone: '+234 915 210 3969',
  website: 'https://okvtechnology.com',
  download_file_id: '',
  download_url: '',
  admin_url: '',
  bank_name: '',
  account_name: '',
  account_number: '',
  termii_api_key: '',
  termii_sender_id: '',
  site_url: '',
  pricing_url: 'https://www.okvtechnology.com/okvrms',
  price_label: 'one-time payment',
  old_price: '',
  promo_text: '',
  download_file_id_mac: '',
  download_url_mac: '',
  paystack_auto_approve: '1',
  paystack_enabled: '0',
  paystack_secret_key: '',
  paystack_public_key: '',
  manual_enabled: '1',
  payment_instructions: 'Transfer the exact amount, then send your receipt. We confirm and email your activation code within 24 hours.',
  bank2_name: '',
  bank2_account_name: '',
  bank2_account_number: '',
  webhook_enabled: '0',
  webhook_url: '',
  app_notice: '',
  latest_version: '',
  update_url: '',
  install_service_text: 'OKV Technology Consults can install and set up OKV RMS for you, on-site or remotely, and train your staff.'
};

// ---------------------------------------------------------------- generic helpers
// The data spreadsheet. Works whether this script is attached to a sheet or stands alone:
// setup() creates the sheet by itself and remembers its ID. Nothing has to be made by hand.
function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const bound = SpreadsheetApp.getActiveSpreadsheet();
  if (bound) return bound;
  throw new Error('The data sheet is not set up yet. In the Apps Script editor, choose the function setup and click Run.');
}
function ensureBook_() {
  const props = PropertiesService.getScriptProperties();
  const bound = SpreadsheetApp.getActiveSpreadsheet();
  if (bound) { props.setProperty('SHEET_ID', bound.getId()); return bound; }
  const existing = props.getProperty('SHEET_ID');
  if (existing) { try { return SpreadsheetApp.openById(existing); } catch (e) { /* deleted: create a new one below */ } }
  const book = SpreadsheetApp.create('OKV RMS Data');
  props.setProperty('SHEET_ID', book.getId());
  return book;
}
function tz_() { return Session.getScriptTimeZone() || 'Africa/Lagos'; }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" is missing. Run setup() once from the script editor.');
  return s;
}
function toHex_(bytes) {
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().substr(0, max || 200);
}
function cleanLong_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').trim().substr(0, max || 2000);
}
function validEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]{2,}$/.test(e) && e.length <= 120; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (m) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]; }); }
function userError_(msg) { const e = new Error(msg); e.userFacing = true; return e; }

function cfg_() {
  const out = {};
  Object.keys(DEFAULT_CONFIG).forEach(function (k) { out[k] = DEFAULT_CONFIG[k]; });
  const s = ss_().getSheetByName('Config');
  if (s && s.getLastRow() > 1) {
    s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(function (r) { if (r[0]) out[String(r[0])] = String(r[1]); });
  }
  return out;
}

function readAll_(name) {
  const s = sh_(name), last = s.getLastRow(), cols = SHEETS[name].length;
  if (last < 2) return [];
  const vals = s.getRange(2, 1, last - 1, cols).getValues();
  return vals.map(function (v, i) {
    const o = { _row: i + 2 };
    SHEETS[name].forEach(function (h, j) { o[h] = v[j] === null || v[j] === undefined ? '' : (v[j] instanceof Date ? Utilities.formatDate(v[j], tz_(), 'yyyy-MM-dd HH:mm:ss') : String(v[j])); });
    return o;
  });
}
function appendRow_(name, obj) {
  const s = sh_(name);
  const row = SHEETS[name].map(function (h) { return obj[h] == null ? '' : String(obj[h]); });
  s.getRange(s.getLastRow() + 1, 1, 1, row.length).setValues([row]);
}
function updateRow_(name, rowIndex, patch) {
  const s = sh_(name), cols = SHEETS[name].length;
  const rng = s.getRange(rowIndex, 1, 1, cols);
  const row = rng.getValues()[0];
  SHEETS[name].forEach(function (h, j) { if (Object.prototype.hasOwnProperty.call(patch, h)) row[j] = patch[h] == null ? '' : String(patch[h]); });
  rng.setValues([row]);
}
function nextId_(prefix) {
  const p = PropertiesService.getScriptProperties();
  const key = 'seq_' + prefix;
  const n = Number(p.getProperty(key) || 0) + 1;
  p.setProperty(key, String(n));
  return prefix + '-' + ('0000' + n).slice(-4);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function rateLimit_(key, max, seconds) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(key) || 0);
  if (n >= max) throw userError_('Too many attempts. Please wait a little and try again.');
  c.put(key, String(n + 1), seconds);
}
function dashUrl_() { try { return String(cfg_().admin_url || ''); } catch (e) { return ''; } }

// ---------------------------------------------------------------- activation codes
function hmacBytes_(msg, key) { return Utilities.computeHmacSha256Signature(msg, key || LICENSE_SECRET); }
function codeTag_(serial) {
  const b = hmacBytes_('RMS1:' + serial);
  let s = '';
  for (let i = 0; i < 6; i++) s += CODE_ALPHA.charAt((((b[i] % 256) + 256) % 256) % 32);
  return s;
}
function normCode_(s) {
  let c = String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (c.indexOf('OKV') === 0) c = c.substr(3);
  return c;
}
function codeValid_(c16) {
  if (c16.length !== 16) return false;
  for (let i = 0; i < 16; i++) if (CODE_ALPHA.indexOf(c16.charAt(i)) < 0) return false;
  return codeTag_(c16.substr(0, 10)) === c16.substr(10);
}
function fmtCode_(c) { return 'OKV-' + c.substr(0, 4) + '-' + c.substr(4, 4) + '-' + c.substr(8, 4) + '-' + c.substr(12, 4); }
function randomSerial_() {
  let s = '';
  while (s.length < 10) {
    const b = hmacBytes_(Utilities.getUuid() + ':' + Math.random() + ':' + Date.now(), LICENSE_SECRET + '#r');
    for (let i = 0; i < b.length && s.length < 10; i++) s += CODE_ALPHA.charAt(((b[i] + 256) % 256) % 32);
  }
  return s;
}
function generateCodes_(n) {
  const sh = sh_('Codes');
  const last = sh.getLastRow();
  const existing = {};
  if (last > 1) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { existing[r[0]] = 1; });
  const batch = 'B' + Utilities.formatDate(new Date(), tz_(), 'yyMMdd-HHmm');
  const rows = [];
  let guard = 0;
  while (rows.length < n && guard++ < n * 3) {
    const serial = randomSerial_();
    const code = fmtCode_(serial + codeTag_(serial));
    if (existing[code]) continue;
    existing[code] = 1;
    rows.push([code, 'Available', '', '', '', '', '', batch]);
  }
  if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
  return rows.length;
}

// ---------------------------------------------------------------- one-time setup (run from the editor)
// Put the admin password here ONCE, run setup(), then set this back to ''. Leave it blank after setup.
const INITIAL_ADMIN_PASSWORD = '';

// Run this ONE function from the editor (choose "setup" > Run). It creates the data sheet, every tab,
// the settings and 1,200 activation codes. It does not need you to create anything by hand.
function setup() {
  const book = ensureBook_();
  Object.keys(SHEETS).forEach(function (name) {
    let s = book.getSheetByName(name);
    if (!s) s = book.insertSheet(name);
    const hdr = SHEETS[name];
    s.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#0f7b6c').setFontColor('#ffffff');
    s.getRange(1, 1, s.getMaxRows(), hdr.length).setNumberFormat('@');   // keep phone numbers, dates and codes as plain text
    s.setFrozenRows(1);
  });
  const c = sh_('Config');
  const haveKeys = {};
  if (c.getLastRow() > 1) c.getRange(2, 1, c.getLastRow() - 1, 1).getValues().forEach(function (r) { haveKeys[r[0]] = 1; });
  const missing = Object.keys(DEFAULT_CONFIG).filter(function (k) { return !haveKeys[k]; }).map(function (k) { return [k, DEFAULT_CONFIG[k]]; });
  if (missing.length) c.getRange(c.getLastRow() + 1, 1, missing.length, 2).setValues(missing);
  const first = book.getSheetByName('Sheet1');
  if (first && book.getSheets().length > 1 && first.getLastRow() === 0) book.deleteSheet(first);
  const have = Math.max(0, sh_('Codes').getLastRow() - 1);
  let made = 0;
  if (have < 1000) made = generateCodes_(1200 - have);
  if (INITIAL_ADMIN_PASSWORD) setAdminPassword(INITIAL_ADMIN_PASSWORD);   // sets or changes the admin password
  const msg = 'Setup complete. Data sheet: ' + book.getUrl() + ' | New activation codes: ' + made;
  Logger.log(msg);
  return msg;
}
function setAdminPassword(pw) {
  // Run ONCE from the editor: setAdminPassword('your-strong-password'), then delete the password from the editor.
  if (!pw || String(pw).length < 8) throw new Error('Use at least 8 characters.');
  const p = PropertiesService.getScriptProperties();
  p.setProperty('ADMIN_SALT', Utilities.getUuid());
  p.setProperty('ADMIN_HASH', hashPw_(pw));
  return 'Admin password saved.';
}
function hashPw_(pw) {
  const salt = PropertiesService.getScriptProperties().getProperty('ADMIN_SALT') || '';
  return toHex_(Utilities.computeHmacSha256Signature(String(pw), salt));
}

// ---------------------------------------------------------------- web app entry points
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'public') return json_(publicInfo_());
    if (p.action === 'bind') return json_(bind_(p.code, clean_(p.install, 40)));
    if (p.action === 'verify') return json_(verifyPaystack_(clean_(p.reference, 60)));
    if (p.action === 'app_info') { const c = cfg_(); return json_({ ok: true, notice: c.app_notice, latest_version: c.latest_version, update_url: c.update_url }); }
    return json_({ ok: true, service: 'OKV RMS API' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err.userFacing ? err.message : 'Server error' });
  }
}

function doPost(e) {
  let out;
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    if (raw.length > 9 * 1024 * 1024) throw userError_('That upload is too large.');
    const b = JSON.parse(raw);
    if (b.hp) return json_({ ok: true, ref: 'OK' });                 // honeypot: bots fill this hidden field
    if (b.type === 'trial') out = handleTrial_(b);
    else if (b.type === 'activation_request') out = handleActivation_(b);
    else if (b.type === 'contact') out = handleContact_(b);
    else if (b.type === 'install_request') out = handleInstall_(b);
    else if (b.type === 'paystack_init') out = handlePaystackInit_(b);
    else if (b.type === 'admin_call') out = adminCall_(b);
    else out = { ok: false, error: 'Unknown request' };
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    out = { ok: false, error: err && err.userFacing ? err.message : 'Sorry, something went wrong on our side. Please try again, or message us on WhatsApp.' };
  }
  return json_(out);
}

// ---------------------------------------------------------------- admin bridge (GitHub Pages admin page)
// Only these functions can be called from the admin page. Every one of them checks the admin token itself.
const ADMIN_CALLABLE_ = {
  adminLogin: adminLogin, adminLogout: adminLogout, adminChangePassword: adminChangePassword,
  adminOverview: adminOverview, adminPoll: adminPoll, adminList: adminList,
  adminUpdateSubmission: adminUpdateSubmission, adminSendCode: adminSendCode, adminGenerateCodes: adminGenerateCodes,
  adminSaveCustomer: adminSaveCustomer, adminDeleteCustomer: adminDeleteCustomer, adminSend: adminSend,
  adminMarkRead: adminMarkRead, adminGetConfig: adminGetConfig, adminSaveConfig: adminSaveConfig,
  adminPaymentAction: adminPaymentAction, adminPaystackTest: adminPaystackTest, adminReconcile: adminReconcile,
  adminVerifyReference: adminVerifyReference, adminAutoCheck: adminAutoCheck, adminAutoStatus: adminAutoStatus,
  adminTestWebhook: adminTestWebhook, adminTestEmail: adminTestEmail, adminApprove: adminApprove,
  adminRevokeCode: adminRevokeCode, adminApproveSubmission: adminApproveSubmission
};
function adminCall_(b) {
  const fn = ADMIN_CALLABLE_[String(b.fn || '')];
  if (!fn) throw userError_('That admin action is not available.');
  const args = Array.isArray(b.args) ? b.args : [];
  const result = fn.apply(null, args);
  return { ok: true, result: result === undefined ? null : result };
}

function publicInfo_() {
  const c = cfg_();
  return {
    ok: true, product: c.product_name, price: Number(c.price) || 120000, trial_days: Number(c.trial_days) || 12, currency: c.currency,
    bank_name: c.bank_name, account_name: c.account_name, account_number: c.account_number,
    whatsapp: c.whatsapp, phone: c.phone, email: c.admin_email, website: c.website,
    paystack_enabled: c.paystack_enabled === '1' && !!c.paystack_secret_key && !!c.site_url, manual_enabled: c.manual_enabled !== '0',
    payment_instructions: c.payment_instructions, bank2_name: c.bank2_name, bank2_account_name: c.bank2_account_name, bank2_account_number: c.bank2_account_number,
    install_service_text: c.install_service_text, price_label: c.price_label, old_price: Number(c.old_price) || 0, promo_text: c.promo_text, pricing_url: c.pricing_url, currency_symbol: c.currency
  };
}

// ---------------------------------------------------------------- form handlers
function common_(b, required) {
  const f = {
    name: clean_(b.name, 80), business: clean_(b.business, 100), email: clean_(b.email, 120).toLowerCase(), phone: clean_(b.phone, 30),
    business_type: clean_(b.business_type, 30), state: clean_(b.state, 40), source: clean_(b.source, 20) || 'landing',
    install_id: clean_(b.install_id, 40), message: cleanLong_(b.message, 2000), subject: clean_(b.subject, 80),
    amount: clean_(b.amount, 20), pay_date: clean_(b.pay_date, 20), pay_ref: clean_(b.pay_ref, 80), heard: clean_(b.heard, 60)
  };
  if (required.indexOf('name') > -1 && f.name.length < 2) throw userError_('Please enter your full name.');
  if (required.indexOf('business') > -1 && f.business.length < 2) throw userError_('Please enter your business name.');
  if (!validEmail_(f.email)) throw userError_('Please enter a valid email address.');
  if (required.indexOf('phone') > -1 && f.phone.replace(/\D/g, '').length < 7) throw userError_('Please enter a valid phone / WhatsApp number.');
  return f;
}

function upsertCustomer_(f, patch) {
  const rows = readAll_('Customers');
  let hit = null;
  for (let i = 0; i < rows.length; i++) if (rows[i].Email.toLowerCase() === f.email) { hit = rows[i]; break; }
  const t = now_();
  const base = { Name: f.name, Business: f.business, 'Business Type': f.business_type, Phone: f.phone, State: f.state, Updated: t };
  if (f.install_id) base['Install ID'] = f.install_id;
  Object.keys(base).forEach(function (k) { if (!base[k]) delete base[k]; });
  Object.keys(patch || {}).forEach(function (k) { base[k] = patch[k]; });
  if (hit) updateRow_('Customers', hit._row, base);
  else { base.Email = f.email; base.Created = t; if (!base.Status) base.Status = 'Lead'; appendRow_('Customers', base); }
}

function alert_(type, title, detail, ref) {
  appendRow_('Alerts', { ID: nextId_('ALT'), Timestamp: now_(), Type: type, Title: title, Detail: detail, Ref: ref || '', Read: 'FALSE' });
  fireWebhook_('alert', { type: type, title: title, detail: detail, ref: ref || '', time: now_() });
}
function audit_(action, detail) {
  try { appendRow_('Audit', { Timestamp: now_(), Action: action, Detail: String(detail || '').substr(0, 300) }); } catch (e) { }
}
function fireWebhook_(event, payload) {
  try {
    const c = cfg_();
    if (c.webhook_enabled !== '1' || !/^https:\/\//.test(c.webhook_url)) return;
    UrlFetchApp.fetch(c.webhook_url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify({ source: 'okv-rms', event: event, data: payload }) });
  } catch (e) { console.error('webhook', e); }
}

function mail_(to, subject, html, opts) {
  const c = cfg_();
  const o = { to: to, subject: subject, htmlBody: wrapMail_(html), name: c.from_name, replyTo: c.admin_email };
  if (opts && opts.cc) o.cc = opts.cc;
  MailApp.sendEmail(o);
}
function wrapMail_(inner) {
  const c = cfg_();
  return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:600px;margin:auto;color:#1b2430">' +
    '<div style="background:#14312c;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0;font-size:18px;font-weight:700">' + esc_(c.product_name) + ' <span style="font-weight:400;font-size:13px;opacity:.8">by OKV Technology</span></div>' +
    '<div style="border:1px solid #d9dfe6;border-top:0;padding:20px;border-radius:0 0 10px 10px;line-height:1.55;font-size:15px">' + inner +
    '<hr style="border:0;border-top:1px solid #e5e9ee;margin:22px 0 12px"><div style="font-size:13px;color:#667386">OKV Technology Consults &middot; ' + esc_(c.phone) + ' &middot; ' + esc_(c.admin_email) + ' &middot; <a href="' + esc_(c.website) + '">' + esc_(c.website) + '</a></div></div></div>';
}
function safeMail_(to, subject, html, refForAlert) {
  try { mail_(to, subject, html); return true; }
  catch (err) { console.error(err); try { alert_('Email failed', 'Could not email ' + to, String(err.message || err).substr(0, 300), refForAlert); } catch (e2) { } return false; }
}
function adminAlertMail_(subject, rowsHtml, id) {
  const c = cfg_();
  const link = dashUrl_();
  safeMail_(c.admin_email, '[OKV RMS] ' + subject,
    '<h3 style="margin:0 0 10px">' + esc_(subject) + '</h3><table style="border-collapse:collapse;width:100%;font-size:14px">' + rowsHtml + '</table>' +
    (link ? '<p><a href="' + esc_(link) + '" style="background:#0f7b6c;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none;display:inline-block">Open the admin dashboard</a></p>' : ''), id);
}
function rowsHtml_(pairs) {
  return pairs.filter(function (p) { return p[1]; }).map(function (p) {
    return '<tr><td style="padding:5px 10px 5px 0;color:#667386;vertical-align:top;white-space:nowrap">' + esc_(p[0]) + '</td><td style="padding:5px 0"><b>' + esc_(p[1]) + '</b></td></tr>';
  }).join('');
}

function downloadUrl_() {
  const c = cfg_();
  if (c.download_url) return c.download_url;
  if (c.download_file_id) return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(c.download_file_id);
  return '';
}

function downloadUrlMac_() {
  const c = cfg_();
  if (c.download_url_mac) return c.download_url_mac;
  if (c.download_file_id_mac) return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(c.download_file_id_mac);
  return '';
}
function handleTrial_(b) {
  const f = common_(b, ['name', 'business', 'phone']);
  rateLimit_('trial:' + f.email, 5, 3600);
  let id;
  withLock_(function () {
    id = nextId_('TRL');
    appendRow_('Submissions', { ID: id, Timestamp: now_(), Type: 'Trial Download', Status: 'New', Name: f.name, Business: f.business, 'Business Type': f.business_type, Email: f.email, Phone: f.phone, State: f.state, Source: f.source, Message: f.heard ? 'Heard about us: ' + f.heard : '', Updated: now_() });
    upsertCustomer_(f, { Status: 'Trial', 'Trial Start': now_().substr(0, 10) });
    alert_('Trial download', f.name + ' downloaded the free trial', f.business + ' \u00b7 ' + f.email + ' \u00b7 ' + f.phone, id);
  });
  const c = cfg_();
  const url = downloadUrl_();
  safeMail_(f.email, 'Your OKV RMS free trial download',
    '<p>Hello ' + esc_(f.name) + ',</p><p>Thank you for trying <b>OKV RMS</b>, the restaurant and supermarket management system. Your <b>' + esc_(c.trial_days) + '-day free trial</b> starts the first time you open the program.</p>' +
    (url ? '<p><a href="' + esc_(url) + '" style="background:#0f7b6c;color:#fff;padding:12px 18px;border-radius:6px;text-decoration:none;display:inline-block;font-weight:600">Download for Windows</a></p>' : '') + (downloadUrlMac_() ? '<p>Using a Mac? <a href="' + esc_(downloadUrlMac_()) + '">Download for Mac</a></p>' : '') +
    '<p><b>Install in 4 steps</b></p><ol><li>Open the download. On Windows run <b>OKV_RMS_Setup.exe</b> and press Next. On Mac open the .dmg and drag OKV RMS into Applications.</li><li>Nothing else needs to be installed. The program opens in your browser. Sign in with <b>admin</b> / <b>admin123</b> and change the password.</li><li>Open <b>Settings</b>, choose Restaurant, Supermarket or Both, and start selling.</li></ol>' +
    '<p>User manuals for every role (PDF) are inside the ZIP, and each staff member can download theirs from the program.</p>' +
    '<p>Full version: <b>' + esc_(c.currency) + Number(c.price).toLocaleString('en-US') + ' one-time</b>, no monthly fees. You can activate from inside the program at any time.</p><p>Need help? Reply to this email or WhatsApp ' + esc_(c.phone) + '.</p>', id);
  adminAlertMail_('New free-trial download ' + id, rowsHtml_([['Reference', id], ['Name', f.name], ['Business', f.business], ['Type', f.business_type], ['Email', f.email], ['Phone', f.phone], ['State', f.state], ['Heard about us', f.heard]]), id);
  return { ok: true, ref: id, downloadUrl: url, downloadUrlMac: downloadUrlMac_() };
}

function saveReceipt_(id, who, receipt) {
  if (!receipt || !receipt.data) return '';
  const mime = String(receipt.mime || '');
  if (!/^(image\/(jpeg|png|webp)|application\/pdf)$/.test(mime)) throw userError_('The receipt must be a JPG, PNG or PDF file.');
  const bytes = Utilities.base64Decode(String(receipt.data));
  if (bytes.length > 6 * 1024 * 1024) throw userError_('The receipt file is too big (6 MB maximum).');
  const ext = mime === 'application/pdf' ? 'pdf' : (mime === 'image/png' ? 'png' : (mime === 'image/webp' ? 'webp' : 'jpg'));
  const blob = Utilities.newBlob(bytes, mime, id + '_' + who.replace(/[^A-Za-z0-9]+/g, '_').substr(0, 30) + '.' + ext);
  return receiptFolder_().createFile(blob).getUrl();
}
function receiptFolder_() {
  const p = PropertiesService.getScriptProperties();
  const fid = p.getProperty('RECEIPT_FOLDER_ID');
  if (fid) { try { return DriveApp.getFolderById(fid); } catch (e) { } }
  const f = DriveApp.createFolder('OKV RMS Receipts');
  p.setProperty('RECEIPT_FOLDER_ID', f.getId());
  return f;
}

function handleActivation_(b) {
  const f = common_(b, ['name', 'phone']);
  const amount = Number(String(f.amount).replace(/[^0-9.]/g, ''));
  if (!amount || amount < 1) throw userError_('Please enter the amount you paid.');
  if (!f.pay_ref && !(b.receipt && b.receipt.data)) throw userError_('Please add the payment reference or upload the receipt.');
  rateLimit_('act:' + f.email, 6, 3600);
  let id, receiptUrl = '';
  id = withLock_(function () { return nextId_('ACT'); });
  receiptUrl = saveReceipt_(id, f.name, b.receipt);
  withLock_(function () {
    appendRow_('Submissions', { ID: id, Timestamp: now_(), Type: 'Activation Request', Status: 'New', Name: f.name, Business: f.business, 'Business Type': f.business_type, Email: f.email, Phone: f.phone, State: f.state, Source: f.source, 'Install ID': f.install_id, Amount: amount, 'Pay Date': f.pay_date, 'Pay Ref': f.pay_ref, 'Receipt URL': receiptUrl, Message: f.message, Updated: now_() });
    upsertCustomer_(f, { 'Install ID': f.install_id });
    appendRow_('Payments', { ID: nextId_('PAY'), Timestamp: now_(), Method: 'Manual', Reference: f.pay_ref || id, Name: f.name, Email: f.email, Phone: f.phone, Amount: amount, Currency: 'NGN', Status: 'Pending', Submission: id, 'Receipt URL': receiptUrl, Updated: now_() });
    alert_('Activation request', 'Activation request from ' + f.name + ' (' + id + ')', (f.business || '') + ' \u00b7 ' + f.email + ' \u00b7 paid ' + amount + (receiptUrl ? ' \u00b7 receipt attached' : ''), id);
  });
  const c = cfg_();
  safeMail_(f.email, 'We received your OKV RMS activation request (' + id + ')',
    '<p>Hello ' + esc_(f.name) + ',</p><p>Thank you. We have received your payment details for <b>OKV RMS</b>.</p><table style="font-size:14px">' + rowsHtml_([['Reference', id], ['Amount', c.currency + amount], ['Payment ref', f.pay_ref], ['Install ID', f.install_id]]) + '</table>' +
    '<p>Our team will confirm the payment and send your <b>activation code to this email address within 24 hours</b> (usually much sooner).</p><p>When it arrives, open OKV RMS, go to <b>License</b>, type the code in the box and press <b>Activate</b>. Your data stays untouched.</p>', id);
  adminAlertMail_('New activation request ' + id + ' - ' + f.name, rowsHtml_([['Reference', id], ['Name', f.name], ['Business', f.business], ['Email', f.email], ['Phone', f.phone], ['Amount', String(amount)], ['Pay date', f.pay_date], ['Pay ref', f.pay_ref], ['Install ID', f.install_id], ['Source', f.source], ['Receipt', receiptUrl], ['Note', f.message]]), id);
  return { ok: true, ref: id };
}

function handleContact_(b) {
  const f = common_(b, ['name']);
  if (f.message.length < 5) throw userError_('Please type your message.');
  rateLimit_('ct:' + f.email, 6, 3600);
  let id;
  withLock_(function () {
    id = nextId_('MSG');
    appendRow_('Submissions', { ID: id, Timestamp: now_(), Type: 'Contact', Status: 'New', Name: f.name, Business: f.business, Email: f.email, Phone: f.phone, Source: f.source, 'Install ID': f.install_id, Message: (f.subject ? '[' + f.subject + '] ' : '') + f.message, Updated: now_() });
    upsertCustomer_(f, {});
    alert_('Contact message', 'Message from ' + f.name + (f.subject ? ' (' + f.subject + ')' : ''), f.message.substr(0, 160), id);
  });
  safeMail_(f.email, 'We received your message (' + id + ')', '<p>Hello ' + esc_(f.name) + ',</p><p>Thank you for contacting OKV Technology. We received your message and will reply to this email address shortly.</p><p style="color:#667386">Your message:<br>' + esc_(f.message).replace(/\n/g, '<br>') + '</p>', id);
  adminAlertMail_('New message ' + id + ' from ' + f.name, rowsHtml_([['Reference', id], ['Name', f.name], ['Business', f.business], ['Email', f.email], ['Phone', f.phone], ['Topic', f.subject], ['Source', f.source], ['Message', f.message]]), id);
  return { ok: true, ref: id };
}

// Called by the installed program when a code is entered (needs internet; the program also checks codes offline).
function bind_(rawCode, install) {
  rateLimit_('bind:' + (install || 'x'), 20, 3600);
  const c16 = normCode_(rawCode);
  if (!codeValid_(c16)) return { ok: false, error: 'This activation code is not valid. Please check it and try again.' };
  const code = fmtCode_(c16);
  return withLock_(function () {
    const rows = readAll_('Codes');
    let row = null;
    for (let i = 0; i < rows.length; i++) if (rows[i].Code === code) { row = rows[i]; break; }
    if (!row) return { ok: false, error: 'This code is not registered with OKV Technology. Please contact us.' };
    if (row.Status === 'Revoked') return { ok: false, error: 'This code has been cancelled. Please contact OKV Technology.' };
    if (row.Status === 'Used' && row['Bound Install'] && row['Bound Install'] !== install) {
      return { ok: false, error: 'This code has already been used on another computer. Please contact OKV Technology.' };
    }
    if (row.Status !== 'Used') updateRow_('Codes', row._row, { Status: 'Used', 'Bound Install': install, 'Used On': now_() });
    const cust = readAll_('Customers');
    for (let i = 0; i < cust.length; i++) {
      if (cust[i].Code === code || (row['Sent To'] && cust[i].Email.toLowerCase() === row['Sent To'].toLowerCase())) {
        updateRow_('Customers', cust[i]._row, { Status: 'Active', 'Activated On': now_().substr(0, 10), Code: code, 'Install ID': install, Updated: now_() });
        break;
      }
    }
    if (row.Status !== 'Used') {
      alert_('Code activated', 'Activation code used: ' + code, 'Sent to ' + (row['Sent To'] || 'n/a') + ' \u00b7 install ' + install, row['Request ID']);
      adminAlertMail_('Activation code used ' + code, rowsHtml_([['Code', code], ['Sent to', row['Sent To']], ['Install ID', install], ['Request', row['Request ID']]]), row['Request ID']);
    }
    return { ok: true };
  });
}

// ================================================================ ADMIN DASHBOARD API (google.script.run)
function auth_(t) {
  if (!t || !CacheService.getScriptCache().get('tok_' + t)) throw new Error('SESSION');
}
function adminLogin(password) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('adm_fail') || 0);
  if (fails >= 8) throw new Error('Too many wrong attempts. Wait 15 minutes and try again.');
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH');
  if (!stored) throw new Error('Admin password not set yet. Run setAdminPassword() once in the script editor.');
  if (hashPw_(password) !== stored) { cache.put('adm_fail', String(fails + 1), 900); throw new Error('Wrong password.'); }
  cache.remove('adm_fail');
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('tok_' + token, '1', 21600);
  return token;
}
function adminLogout(token) { CacheService.getScriptCache().remove('tok_' + token); return true; }
function adminChangePassword(token, oldPw, newPw) {
  auth_(token);
  if (hashPw_(oldPw) !== PropertiesService.getScriptProperties().getProperty('ADMIN_HASH')) throw new Error('Current password is wrong.');
  setAdminPassword(newPw);
  return true;
}

function adminOverview(token) {
  auth_(token);
  const subs = readAll_('Submissions'), cust = readAll_('Customers'), codes = readAll_('Codes'), alerts = readAll_('Alerts');
  const count = function (arr, fn) { let n = 0; arr.forEach(function (x) { if (fn(x)) n++; }); return n; };
  const trialDays = Number(cfg_().trial_days) || 12;
  const today = new Date();
  const daily = [];
  for (let d = 13; d >= 0; d--) {
    const dt = new Date(today.getTime() - d * 86400000);
    const key = Utilities.formatDate(dt, tz_(), 'yyyy-MM-dd');
    daily.push({ date: key, trials: count(subs, function (s) { return s.Timestamp.substr(0, 10) === key && s.Type === 'Trial Download'; }), requests: count(subs, function (s) { return s.Timestamp.substr(0, 10) === key && s.Type === 'Activation Request'; }) });
  }
  const strip = function (o) { const c = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') c[k] = o[k]; }); return c; };
  let quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { }
  return {
    counts: {
      newSubmissions: count(subs, function (s) { return s.Status === 'New'; }),
      pendingActivations: count(subs, function (s) { return s.Type === 'Activation Request' && (s.Status === 'New' || s.Status === 'In Progress'); }),
      trialDownloads: count(subs, function (s) { return s.Type === 'Trial Download'; }),
      contactNew: count(subs, function (s) { return s.Type === 'Contact' && s.Status === 'New'; }),
      customers: cust.length, active: count(cust, function (c) { return c.Status === 'Active'; }), trial: count(cust, function (c) { return c.Status === 'Trial'; }),
      codesAvailable: count(codes, function (c) { return c.Status === 'Available'; }), codesSent: count(codes, function (c) { return c.Status === 'Sent'; }), codesUsed: count(codes, function (c) { return c.Status === 'Used'; }),
      unreadAlerts: count(alerts, function (a) { return a.Read !== 'TRUE'; }), trialDays: trialDays
    },
    pending: subs.filter(function (s) { return s.Type === 'Activation Request' && (s.Status === 'New' || s.Status === 'In Progress'); }).reverse().slice(0, 8).map(strip),
    recent: subs.slice(-8).reverse().map(strip),
    alerts: alerts.slice(-8).reverse().map(strip),
    daily: daily, quota: quota, extra: extraStats_(subs, cust, codes),
    pendingPayments: readAll_('Payments').filter(function (p) { return !p.Code && ((p.Status === 'Pending' && p.Method === 'Manual') || p.Status === 'Paid'); }).reverse().slice(0, 8).map(strip)
  };
}

function adminPoll(token) {
  auth_(token);
  const alerts = readAll_('Alerts');
  const unread = alerts.filter(function (a) { return a.Read !== 'TRUE'; });
  return { unread: unread.length, latest: unread.slice(-5).reverse().map(function (a) { return { ID: a.ID, Title: a.Title, Type: a.Type, Timestamp: a.Timestamp }; }), newestId: alerts.length ? alerts[alerts.length - 1].ID : '' };
}

function adminList(token, kind, opts) {
  auth_(token);
  opts = opts || {};
  if (['Submissions', 'Customers', 'Codes', 'Alerts', 'Messages', 'Payments', 'Audit'].indexOf(kind) < 0) throw new Error('Unknown list');
  let rows = readAll_(kind);
  if (opts.status) rows = rows.filter(function (r) { return r.Status === opts.status; });
  if (opts.type) rows = rows.filter(function (r) { return r.Type === opts.type; });
  if (opts.q) {
    const q = String(opts.q).toLowerCase();
    rows = rows.filter(function (r) { return Object.keys(r).some(function (k) { return k !== '_row' && String(r[k]).toLowerCase().indexOf(q) > -1; }); });
  }
  if (['Submissions', 'Alerts', 'Messages', 'Customers', 'Payments', 'Audit'].indexOf(kind) > -1) rows = rows.reverse();
  const size = Math.min(Math.max(Number(opts.size) || 25, 5), 100), page = Math.max(Number(opts.page) || 1, 1);
  const total = rows.length;
  return {
    total: total, page: page, size: size,
    rows: rows.slice((page - 1) * size, page * size).map(function (r) { const c = {}; Object.keys(r).forEach(function (k) { if (k !== '_row') c[k] = r[k]; }); return c; })
  };
}

function adminUpdateSubmission(token, id, patch) {
  auth_(token);
  const rows = readAll_('Submissions');
  for (let i = 0; i < rows.length; i++) if (rows[i].ID === id) {
    const p = { Updated: now_() };
    if (patch.status && ['New', 'In Progress', 'Code Sent', 'Done', 'Rejected'].indexOf(patch.status) > -1) p.Status = patch.status;
    if (patch.note != null) p.Note = cleanLong_(patch.note, 1000);
    updateRow_('Submissions', rows[i]._row, p);
    return true;
  }
  throw new Error('Submission not found');
}

function adminSendCode(token, p) {
  auth_(token);
  const email = clean_(p.email, 120).toLowerCase();
  const name = clean_(p.name, 80);
  if (!validEmail_(email)) throw new Error('Enter a valid customer email address.');
  const c = cfg_();
  const picked = withLock_(function () {
    const codes = readAll_('Codes');
    let row = null;
    if (p.code) {
      for (let i = 0; i < codes.length; i++) if (codes[i].Code === p.code) { row = codes[i]; break; }
      if (!row) throw new Error('That code is not in the list.');
      if (row.Status === 'Used') throw new Error('That code is already used.');
      if (row.Status === 'Sent' && row['Sent To'].toLowerCase() !== email) throw new Error('That code was already sent to ' + row['Sent To'] + '.');
    } else {
      for (let i = 0; i < codes.length; i++) if (codes[i].Status === 'Available') { row = codes[i]; break; }
      if (!row) throw new Error('No ready codes left. Open Activation Codes and generate more.');
    }
    const prev = { Status: row.Status, 'Sent To': row['Sent To'], 'Sent On': row['Sent On'], 'Request ID': row['Request ID'] };
    updateRow_('Codes', row._row, { Status: 'Sent', 'Sent To': email, 'Sent On': now_(), 'Request ID': p.submissionId || row['Request ID'] });
    return { row: row, prev: prev };
  });
  const code = picked.row.Code;
  try {
    mail_(email, 'Your OKV RMS activation code', codeMailHtml_(name, code, p.note));
  } catch (err) {
    withLock_(function () { updateRow_('Codes', picked.row._row, picked.prev); });
    throw new Error('The email could not be sent (' + String(err.message || err).substr(0, 120) + '). The code was put back in the ready list.');
  }
  if (p.submissionId) {
    const subs = readAll_('Submissions');
    for (let i = 0; i < subs.length; i++) if (subs[i].ID === p.submissionId) { updateRow_('Submissions', subs[i]._row, { Status: 'Code Sent', 'Code Sent': code, Updated: now_() }); break; }
  }
  upsertCustomer_({ name: name, email: email, business: '', phone: '', business_type: '', state: '', install_id: '' }, { Code: code });
  appendRow_('Messages', { Timestamp: now_(), Channel: 'Email', To: email, Subject: 'Your OKV RMS activation code', Body: 'Activation code ' + code, Status: 'Sent', Ref: p.submissionId || '' });
  audit_('Activation code sent', code + ' to ' + email + (p.submissionId ? ' for ' + p.submissionId : ''));
  return { code: code, emailed: true };
}

function adminGenerateCodes(token, n) {
  auth_(token);
  n = Math.floor(Number(n));
  if (!(n >= 1 && n <= 500)) throw new Error('Choose a number from 1 to 500.');
  audit_('Codes generated', n);
  return withLock_(function () { return { created: generateCodes_(n) }; });
}

function adminSaveCustomer(token, c) {
  auth_(token);
  const email = clean_(c.Email, 120).toLowerCase();
  if (!validEmail_(email)) throw new Error('Enter a valid email address.');
  const allowed = ['Name', 'Business', 'Business Type', 'Phone', 'State', 'Status', 'Notes'];
  const patch = { Updated: now_() };
  allowed.forEach(function (k) { if (c[k] != null) patch[k] = clean_(c[k], k === 'Notes' ? 500 : 120); });
  if (['Lead', 'Trial', 'Active', 'Expired'].indexOf(patch.Status) < 0) patch.Status = 'Lead';
  return withLock_(function () {
    const rows = readAll_('Customers');
    for (let i = 0; i < rows.length; i++) if (rows[i].Email.toLowerCase() === email) { updateRow_('Customers', rows[i]._row, patch); return true; }
    patch.Email = email; patch.Created = now_();
    appendRow_('Customers', patch);
    return true;
  });
}
function adminDeleteCustomer(token, email) {
  auth_(token);
  return withLock_(function () {
    const rows = readAll_('Customers');
    for (let i = 0; i < rows.length; i++) if (rows[i].Email.toLowerCase() === String(email).toLowerCase()) { sh_('Customers').deleteRow(rows[i]._row); return true; }
    throw new Error('Customer not found');
  });
}

function normPhone_(p) {
  let d = String(p || '').replace(/\D/g, '');
  if (d.indexOf('00') === 0) d = d.substr(2);
  if (d.indexOf('0') === 0 && d.length === 11) d = '234' + d.substr(1);
  else if (/^[789]\d{9}$/.test(d)) d = '234' + d;
  return d;
}
function fill_(tpl, r) {
  return String(tpl).replace(/\{name\}/gi, r.name || 'there').replace(/\{business\}/gi, r.business || 'your business');
}
function recipients_(p) {
  let list = [];
  if (p.segment) {
    const tdays = Number(cfg_().trial_days) || 12;
    readAll_('Customers').forEach(function (c) {
      const age = c['Trial Start'] ? Math.floor((new Date() - new Date(c['Trial Start'] + 'T00:00:00')) / 86400000) : 0;
      const seg = c.Status === 'Active' ? 'active' : (c.Status === 'Expired' || (c.Status === 'Trial' && age > tdays)) ? 'expired' : c.Status === 'Trial' ? 'trial' : 'lead';
      if (p.segment === 'all' || p.segment === seg) list.push({ name: c.Name, business: c.Business, email: c.Email, phone: c.Phone });
    });
  }
  (p.to || []).forEach(function (r) { list.push({ name: clean_(r.name, 80), business: clean_(r.business, 100), email: clean_(r.email, 120), phone: clean_(r.phone, 30) }); });
  const seen = {};
  return list.filter(function (r) { const k = (r.email || r.phone || '').toLowerCase(); if (!k || seen[k]) return false; seen[k] = 1; return true; });
}
function adminSend(token, p) {
  auth_(token);
  const channel = p.channel;
  const body = cleanLong_(p.body, 4000), subject = clean_(p.subject, 150);
  if (!body) throw new Error('Type a message first.');
  const list = recipients_(p);
  if (!list.length) throw new Error('No recipients selected.');
  const out = { channel: channel, sent: 0, failed: [], links: [] };
  const c = cfg_();
  if (channel === 'email') {
    if (!subject) throw new Error('Add a subject for the email.');
    const withMail = list.filter(function (r) { return validEmail_(r.email); });
    if (!withMail.length) throw new Error('None of the selected people has a valid email address.');
    const quota = MailApp.getRemainingDailyQuota();
    if (withMail.length > quota) throw new Error('Google allows ' + quota + ' more emails today. You selected ' + withMail.length + '. Send fewer, or try tomorrow.');
    withMail.forEach(function (r) {
      try {
        mail_(r.email, fill_(subject, r), '<p>' + esc_(fill_(body, r)).replace(/\n/g, '<br>') + '</p>');
        appendRow_('Messages', { Timestamp: now_(), Channel: 'Email', To: r.email, Subject: fill_(subject, r), Body: fill_(body, r), Status: 'Sent' });
        out.sent++;
      } catch (err) { out.failed.push(r.email + ': ' + String(err.message || err).substr(0, 80)); appendRow_('Messages', { Timestamp: now_(), Channel: 'Email', To: r.email, Subject: subject, Body: body, Status: 'Failed' }); }
    });
  } else if (channel === 'whatsapp') {
    list.forEach(function (r) {
      const ph = normPhone_(r.phone);
      if (ph.length < 10) { out.failed.push((r.name || r.email) + ': no valid phone'); return; }
      const text = fill_(body, r);
      out.links.push({ name: r.name || r.email, phone: ph, url: 'https://wa.me/' + ph + '?text=' + encodeURIComponent(text) });
      appendRow_('Messages', { Timestamp: now_(), Channel: 'WhatsApp', To: ph, Subject: '', Body: text, Status: 'Link opened by admin' });
    });
  } else if (channel === 'sms') {
    list.forEach(function (r) {
      const ph = normPhone_(r.phone);
      if (ph.length < 10) { out.failed.push((r.name || r.email) + ': no valid phone'); return; }
      const text = fill_(body, r);
      if (c.termii_api_key && c.termii_sender_id) {
        try {
          const res = UrlFetchApp.fetch('https://api.ng.termii.com/api/sms/send', { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify({ to: ph, from: c.termii_sender_id, sms: text, type: 'plain', channel: 'generic', api_key: c.termii_api_key }) });
          const ok = res.getResponseCode() >= 200 && res.getResponseCode() < 300;
          appendRow_('Messages', { Timestamp: now_(), Channel: 'SMS', To: ph, Subject: '', Body: text, Status: ok ? 'Sent' : 'Failed ' + res.getResponseCode() });
          if (ok) out.sent++; else out.failed.push(ph + ': gateway error ' + res.getResponseCode());
        } catch (err) { out.failed.push(ph + ': ' + String(err.message || err).substr(0, 80)); }
      } else {
        out.links.push({ name: r.name || r.email, phone: ph, url: 'sms:+' + ph + '?body=' + encodeURIComponent(text) });
        appendRow_('Messages', { Timestamp: now_(), Channel: 'SMS', To: ph, Subject: '', Body: text, Status: 'Link opened by admin' });
      }
    });
  } else throw new Error('Choose Email, WhatsApp or SMS.');
  return out;
}

function adminMarkRead(token, ids) {
  auth_(token);
  const rows = readAll_('Alerts');
  const all = ids === 'all';
  const set = {};
  if (!all) (ids || []).forEach(function (i) { set[i] = 1; });
  rows.forEach(function (r) { if (r.Read !== 'TRUE' && (all || set[r.ID])) updateRow_('Alerts', r._row, { Read: 'TRUE' }); });
  return true;
}

const SECRET_KEYS = ['termii_api_key', 'paystack_secret_key'];
function adminGetConfig(token) {
  auth_(token);
  const c = cfg_();
  SECRET_KEYS.forEach(function (k) { if (c[k]) c[k] = '********'; });
  return c;
}
function adminSaveConfig(token, kv) {
  auth_(token);
  const s = sh_('Config');
  const rows = s.getLastRow() > 1 ? s.getRange(2, 1, s.getLastRow() - 1, 2).getValues() : [];
  const idx = {}; rows.forEach(function (r, i) { idx[r[0]] = i + 2; });
  Object.keys(kv).forEach(function (k) {
    if (!Object.prototype.hasOwnProperty.call(DEFAULT_CONFIG, k)) return;
    if (SECRET_KEYS.indexOf(k) > -1 && kv[k] === '********') return;
    if (k === 'price' && !(/^\d+(\.\d+)?$/.test(String(kv[k]).trim()) && Number(kv[k]) > 0)) throw new Error('Price must be a number greater than zero.');
    if (k === 'old_price' && String(kv[k]).trim() !== '' && !/^\d+(\.\d+)?$/.test(String(kv[k]).trim())) throw new Error('Old price must be a number or empty.');
    if (k === 'site_url' && kv[k] && !/^https?:\/\//.test(kv[k])) throw new Error('Website address must start with https://');
    const v = (k === 'price' || k === 'old_price') ? String(Number(String(kv[k]).replace(/[^0-9.]/g, '')) || '') : clean_(kv[k], k === 'payment_instructions' || k === 'install_service_text' || k === 'app_notice' ? 600 : 300);
    if (idx[k]) s.getRange(idx[k], 2).setValue(v); else s.getRange(s.getLastRow() + 1, 1, 1, 2).setValues([[k, v]]);
  });
  audit_('Settings saved', Object.keys(kv).filter(function (k) { return SECRET_KEYS.indexOf(k) < 0; }).join(', '));
  return true;
}


// ================================================================ INSTALLATION SERVICE
function handleInstall_(b) {
  const f = common_(b, ['name', 'phone']);
  rateLimit_('inst:' + f.email, 4, 3600);
  const loc = clean_(b.location, 120), devices = clean_(b.devices, 20), mode = clean_(b.mode, 20), when = clean_(b.preferred_date, 20);
  let id;
  withLock_(function () {
    id = nextId_('INS');
    appendRow_('Submissions', { ID: id, Timestamp: now_(), Type: 'Installation Request', Status: 'New', Name: f.name, Business: f.business, 'Business Type': f.business_type, Email: f.email, Phone: f.phone, State: f.state, Source: f.source,
      Message: 'Location: ' + loc + ' | Devices/tills: ' + devices + ' | Preferred: ' + mode + ' | Date: ' + when + (f.message ? ' | Note: ' + f.message : ''), Updated: now_() });
    upsertCustomer_(f, {});
    alert_('Installation request', 'Installation request from ' + f.name + ' (' + id + ')', (f.business || '') + ' \u00b7 ' + loc + ' \u00b7 ' + mode + ' \u00b7 ' + f.phone, id);
  });
  safeMail_(f.email, 'We received your installation request (' + id + ')',
    '<p>Hello ' + esc_(f.name) + ',</p><p>Thank you. OKV Technology Consults has received your request for help installing <b>OKV RMS</b>.</p><table style="font-size:14px">' + rowsHtml_([['Reference', id], ['Where', loc], ['Tills / devices', devices], ['You prefer', mode], ['Preferred date', when]]) + '</table><p>A member of our team will contact you on <b>' + esc_(f.phone) + '</b> to agree the time and the price.</p>', id);
  adminAlertMail_('New installation request ' + id + ' - ' + f.name, rowsHtml_([['Reference', id], ['Name', f.name], ['Business', f.business], ['Phone', f.phone], ['Email', f.email], ['Location', loc], ['Devices', devices], ['Preferred', mode], ['Date', when], ['Note', f.message]]), id);
  return { ok: true, ref: id };
}

// ================================================================ ACTIVATION CODE DELIVERY (shared)
function codeMailHtml_(name, code, note) {
  const url = downloadUrl_();
  return '<p>Hello ' + esc_(name || 'there') + ',</p><p>Thank you for your payment. Your <b>OKV RMS activation code</b> is:</p>' +
    '<p style="font:700 24px Consolas,\'Courier New\',monospace;letter-spacing:1px;background:#f1f6f5;border:1px dashed #0f7b6c;border-radius:8px;padding:14px 18px;display:inline-block">' + esc_(code) + '</p>' +
    '<p><b>How to activate</b></p><ol><li>Open OKV RMS and sign in as the <b>admin</b>.</li><li>Click <b>License</b> in the top menu.</li><li>Type the code in the box under <i>Already have your activation code?</i> and press <b>Activate</b>.</li></ol>' +
    '<p>The code works for one computer (one installation). Keep this email safe. Your sales, stock and settings are untouched by activation.</p>' +
    (url ? '<p><b>Need to install OKV RMS on this or another computer?</b><br>Download the complete package (Setup.exe, instructions and manuals): <a href="' + esc_(url) + '">Download OKV RMS</a><br>It installs with a simple wizard on Windows 7, 8, 10 and 11 (32-bit and 64-bit). Nothing else needs to be installed.</p>' : '') +
    (note ? '<p>' + esc_(cleanLong_(note, 500)).replace(/\n/g, '<br>') + '</p>' : '');
}
function deliverCode_(email, name, subId, payRef, note) {
  const picked = withLock_(function () {
    const codes = readAll_('Codes');
    let row = null;
    for (let i = 0; i < codes.length; i++) if (codes[i].Status === 'Available') { row = codes[i]; break; }
    if (!row) throw new Error('No ready activation codes left. Generate more in the dashboard.');
    const prev = { Status: row.Status, 'Sent To': row['Sent To'], 'Sent On': row['Sent On'], 'Request ID': row['Request ID'] };
    updateRow_('Codes', row._row, { Status: 'Sent', 'Sent To': email, 'Sent On': now_(), 'Request ID': subId || payRef || '' });
    return { row: row, prev: prev };
  });
  const code = picked.row.Code;
  try { mail_(email, 'Your OKV RMS activation code', codeMailHtml_(name, code, note || '')); }
  catch (err) { withLock_(function () { updateRow_('Codes', picked.row._row, picked.prev); }); throw err; }
  if (subId) {
    const subs = readAll_('Submissions');
    for (let i = 0; i < subs.length; i++) if (subs[i].ID === subId) { updateRow_('Submissions', subs[i]._row, { Status: 'Code Sent', 'Code Sent': code, Updated: now_() }); break; }
  }
  upsertCustomer_({ name: name, email: email, business: '', phone: '', business_type: '', state: '', install_id: '' }, { Code: code });
  appendRow_('Messages', { Timestamp: now_(), Channel: 'Email', To: email, Subject: 'Your OKV RMS activation code', Body: 'Activation code ' + code, Status: 'Sent', Ref: subId || payRef || '' });
  return code;
}

// ================================================================ PAYSTACK (online payment, verified server-side)
function paystackFetch_(path, method, payload) {
  const c = cfg_();
  const o = { method: method, headers: { Authorization: 'Bearer ' + c.paystack_secret_key }, muteHttpExceptions: true };
  if (payload) { o.contentType = 'application/json'; o.payload = JSON.stringify(payload); }
  const res = UrlFetchApp.fetch('https://api.paystack.co' + path, o);
  let d = null; try { d = JSON.parse(res.getContentText()); } catch (e) { }
  return { code: res.getResponseCode(), data: d };
}
function ageMin_(ts) { const t = new Date(String(ts).replace(' ', 'T')).getTime(); return isNaN(t) ? 9999 : (Date.now() - t) / 60000; }

function handlePaystackInit_(b) {
  const c = cfg_();
  if (c.paystack_enabled !== '1' || !c.paystack_secret_key || !c.site_url) throw userError_('Online payment is not available right now. Please pay by bank transfer instead.');
  const f = common_(b, ['name', 'phone']);
  rateLimit_('ps:' + f.email, 6, 3600);
  const price = Number(c.price) || 120000;
  const ref = 'OKVRMS-' + Utilities.getUuid().replace(/-/g, '').substr(0, 14).toUpperCase();
  const cb = c.site_url.replace(/\/?$/, '/') + 'pricing.html?paystack=1';
  const r = paystackFetch_('/transaction/initialize', 'post', { email: f.email, amount: Math.round(price * 100), currency: 'NGN', reference: ref, callback_url: cb, metadata: { name: f.name, phone: f.phone, business: f.business, product: 'OKV RMS' } });
  if (!r.data || !r.data.status || !r.data.data || !r.data.data.authorization_url) { console.error('paystack init', r.code, JSON.stringify(r.data)); throw userError_('We could not start the online payment. Please try again, or pay by bank transfer.'); }
  withLock_(function () {
    const subId = nextId_('ACT');
    appendRow_('Submissions', { ID: subId, Timestamp: now_(), Type: 'Activation Request', Status: 'New', Name: f.name, Business: f.business, Email: f.email, Phone: f.phone, Source: 'paystack', Amount: price, 'Pay Ref': ref, Message: 'Online payment started with Paystack', Updated: now_() });
    appendRow_('Payments', { ID: nextId_('PAY'), Timestamp: now_(), Method: 'Paystack', Reference: ref, Name: f.name, Email: f.email, Phone: f.phone, Amount: price, Currency: 'NGN', Status: 'Pending', Submission: subId, Updated: now_() });
    upsertCustomer_(f, {});
    alert_('Online payment started', f.name + ' started a Paystack payment (' + subId + ')', f.email + ' \u00b7 ' + ref, subId);
  });
  return { ok: true, authorization_url: r.data.data.authorization_url, reference: ref };
}

function verifyPaystack_(ref) {
  if (!/^OKVRMS-[A-Z0-9]{8,20}$/.test(ref)) return { ok: false, error: 'This payment reference is not valid.' };
  const c = cfg_();
  const claim = withLock_(function () {
    const rows = readAll_('Payments');
    let row = null;
    for (let i = 0; i < rows.length; i++) if (rows[i].Reference === ref && rows[i].Method === 'Paystack') { row = rows[i]; break; }
    if (!row) return { err: 'We could not find this payment.' };
    if (row.Code && (row.Status === 'Paid' || row.Status === 'Verified')) return { done: row };
    if (row.Status === 'Issuing' && ageMin_(row.Updated) < 5) return { busy: true };
    if (row.Status === 'Rejected') return { err: 'This payment was rejected. Please contact us.' };
    if (row.Status === 'Paid' && !row.Code && c.paystack_auto_approve === '0') return { waiting: true };
    return { row: row };
  });
  if (claim.err) return { ok: false, error: claim.err };
  if (claim.done) return { ok: true, status: 'success', code: claim.done.Code, email: claim.done.Email, downloadUrl: downloadUrl_(), downloadUrlMac: downloadUrlMac_() };
  if (claim.busy) return { ok: true, status: 'pending' };
  if (claim.waiting) return { ok: true, status: 'paid_pending' };
  if (!c.paystack_secret_key) return { ok: false, error: 'Online payment is not set up.' };
  const v = paystackFetch_('/transaction/verify/' + encodeURIComponent(ref), 'get');
  if (!v.data || !v.data.status || !v.data.data) return { ok: true, status: 'pending' };
  const t = v.data.data, row = claim.row;
  if (t.status === 'success') {
    const need = Math.round((Number(c.price) || 120000) * 100);
    if (String(t.currency) !== 'NGN' || Number(t.amount) < need) {
      withLock_(function () { updateRow_('Payments', row._row, { Status: 'Review', Note: 'Amount or currency does not match: ' + t.amount + ' ' + t.currency, Updated: now_() }); });
      alert_('Payment needs review', 'Paystack amount mismatch for ' + ref, String(t.amount / 100) + ' ' + t.currency + ' (expected ' + need / 100 + ' NGN)', row.Submission);
      return { ok: true, status: 'review' };
    }
    if (c.paystack_auto_approve === '0') {
      withLock_(function () { updateRow_('Payments', row._row, { Status: 'Paid', Amount: t.amount / 100, Note: 'Paystack ' + (t.channel || '') + ' (awaiting approval)', Updated: now_() }); });
      alert_('Online payment awaiting approval', 'Paid online, approve to send code: ' + row.Name + ' (' + ref + ')', 'NGN ' + t.amount / 100 + ' \u00b7 ' + row.Email, row.Submission);
      adminAlertMail_('Online payment awaiting your approval - ' + row.Name, rowsHtml_([['Reference', ref], ['Name', row.Name], ['Email', row.Email], ['Amount', 'NGN ' + t.amount / 100]]), row.Submission);
      return { ok: true, status: 'paid_pending' };
    }
    const won = withLock_(function () {
      const fresh = readAll_('Payments').filter(function (r) { return r.ID === row.ID; })[0];
      if (fresh.Code || (fresh.Status === 'Issuing' && ageMin_(fresh.Updated) < 5)) return false;
      updateRow_('Payments', fresh._row, { Status: 'Issuing', Updated: now_() });
      return true;
    });
    if (!won) { const again = readAll_('Payments').filter(function (r) { return r.ID === row.ID; })[0]; return again.Code ? { ok: true, status: 'success', code: again.Code, email: again.Email, downloadUrl: downloadUrl_(), downloadUrlMac: downloadUrlMac_() } : { ok: true, status: 'pending' }; }
    let code;
    try { code = deliverCode_(row.Email, row.Name, row.Submission, ref); }
    catch (err) { withLock_(function () { updateRow_('Payments', row._row, { Status: 'Pending', Note: 'Delivery failed: ' + String(err.message || err).substr(0, 150), Updated: now_() }); }); alert_('Code delivery failed', 'Paid online but code not sent: ' + ref, String(err.message || err).substr(0, 200), row.Submission); throw err; }
    withLock_(function () { updateRow_('Payments', row._row, { Status: 'Paid', Code: code, Amount: t.amount / 100, Note: 'Paystack ' + (t.channel || ''), Updated: now_() }); });
    alert_('Online payment received', 'Paid online: ' + row.Name + ' (' + ref + ')', 'NGN ' + t.amount / 100 + ' \u00b7 code ' + code + ' sent to ' + row.Email, row.Submission);
    adminAlertMail_('Online payment received - ' + row.Name, rowsHtml_([['Reference', ref], ['Name', row.Name], ['Email', row.Email], ['Amount', 'NGN ' + t.amount / 100], ['Channel', t.channel], ['Code sent', code]]), row.Submission);
    audit_('Paystack payment verified', ref + ' -> ' + code);
    return { ok: true, status: 'success', code: code, email: row.Email, downloadUrl: downloadUrl_(), downloadUrlMac: downloadUrlMac_() };
  }
  if (t.status === 'failed' || t.status === 'reversed') {
    withLock_(function () { updateRow_('Payments', row._row, { Status: 'Failed', Note: 'Paystack: ' + t.status, Updated: now_() }); });
    return { ok: true, status: 'failed' };
  }
  return { ok: true, status: 'pending' };
}

function reconcilePaystack() {
  // Runs by itself every 10 minutes once "Auto-check" is switched on (Integrations page).
  let checked = 0, paid = 0;
  readAll_('Payments').forEach(function (r) {
    if (r.Method === 'Paystack' && (r.Status === 'Pending' || r.Status === 'Issuing') && ageMin_(r.Timestamp) >= 2 && ageMin_(r.Timestamp) < 7 * 1440) {
      try { const v = verifyPaystack_(r.Reference); checked++; if (v.status === 'success') paid++; } catch (e) { console.error(e); }
    }
  });
  return { checked: checked, paid: paid };
}

// ================================================================ ADMIN: payments, integrations, stats
function extraStats_(subs, cust, codes) {
  const pays = readAll_('Payments');
  let total = 0, manual = 0, ps = 0, pendingManual = 0, pendingAmt = 0, failed = 0;
  const monthly = {};
  pays.forEach(function (p) {
    const amt = Number(p.Amount) || 0;
    if (p.Status === 'Paid' || p.Status === 'Verified') { total += amt; if (p.Method === 'Paystack') ps += amt; else manual += amt; const m = String(p.Timestamp).substr(0, 7); monthly[m] = (monthly[m] || 0) + amt; }
    else if (p.Status === 'Pending' && p.Method === 'Manual') { pendingManual++; pendingAmt += amt; }
    else if (p.Status === 'Failed' || p.Status === 'Review') failed++;
  });
  const status = { Lead: 0, Trial: 0, Active: 0, Expired: 0 };
  cust.forEach(function (c) { status[c.Status] = (status[c.Status] || 0) + 1; });
  const types = {}; subs.forEach(function (s) { types[s.Type] = (types[s.Type] || 0) + 1; });
  const trials = types['Trial Download'] || 0;
  return { revenue: { total: total, manual: manual, paystack: ps, pendingManual: pendingManual, pendingAmt: pendingAmt, needsReview: failed }, monthly: Object.keys(monthly).sort().slice(-6).map(function (k) { return { month: k, amount: monthly[k] }; }),
    status: status, types: types, installs: types['Installation Request'] || 0, conversion: trials ? Math.round(status.Active / trials * 100) : 0 };
}
function adminPaymentAction(token, id, action, note) {
  auth_(token);
  const map = { verify: 'Verified', reject: 'Rejected', pending: 'Pending' };
  if (!map[action]) throw new Error('Unknown action');
  const rows = readAll_('Payments');
  for (let i = 0; i < rows.length; i++) if (rows[i].ID === id) {
    updateRow_('Payments', rows[i]._row, { Status: map[action], Note: note != null ? clean_(note, 200) : rows[i].Note, Updated: now_() });
    if (action === 'reject' && rows[i].Submission) {
      const subs = readAll_('Submissions');
      for (let j = 0; j < subs.length; j++) if (subs[j].ID === rows[i].Submission) { updateRow_('Submissions', subs[j]._row, { Status: 'Rejected', Updated: now_() }); break; }
    }
    audit_('Payment ' + action, id + ' ' + rows[i].Reference);
    return true;
  }
  throw new Error('Payment not found');
}
function adminPaystackTest(token) {
  auth_(token);
  const c = cfg_();
  if (!c.paystack_secret_key) throw new Error('Save your Paystack secret key first.');
  const r = paystackFetch_('/balance', 'get');
  if (r.code === 401) throw new Error('Paystack rejected this key (401). Copy the secret key again from your Paystack dashboard (Settings > API Keys).');
  if (r.code >= 500 || r.code === 0) throw new Error('Could not reach Paystack (HTTP ' + r.code + '). Try again in a moment.');
  audit_('Paystack key tested', r.code);
  return { ok: true, mode: /^sk_live_/.test(c.paystack_secret_key) ? 'LIVE' : (/^sk_test_/.test(c.paystack_secret_key) ? 'TEST' : 'unknown'), http: r.code, enabled: c.paystack_enabled === '1', siteUrl: !!c.site_url };
}
function adminReconcile(token) { auth_(token); const r = reconcilePaystack(); audit_('Paystack reconcile', JSON.stringify(r)); return r; }
function adminVerifyReference(token, ref) { auth_(token); const r = verifyPaystack_(clean_(ref, 60)); audit_('Paystack verify', ref + ' -> ' + (r.status || r.error)); return r; }
function adminAutoCheck(token, on) {
  auth_(token);
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'reconcilePaystack') ScriptApp.deleteTrigger(t); });
  if (on) ScriptApp.newTrigger('reconcilePaystack').timeBased().everyMinutes(10).create();
  audit_('Paystack auto-check', on ? 'on' : 'off');
  return { on: !!on };
}
function adminAutoStatus(token) { auth_(token); return { on: ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'reconcilePaystack'; }) }; }
function adminTestWebhook(token) {
  auth_(token);
  const c = cfg_();
  if (!/^https:\/\//.test(c.webhook_url)) throw new Error('Save a webhook URL that starts with https:// first.');
  const res = UrlFetchApp.fetch(c.webhook_url, { method: 'post', contentType: 'application/json', muteHttpExceptions: true, payload: JSON.stringify({ source: 'okv-rms', event: 'test', data: { message: 'Test from the OKV RMS admin dashboard', time: now_() } }) });
  audit_('Webhook tested', res.getResponseCode());
  return { http: res.getResponseCode() };
}
function adminTestEmail(token) {
  auth_(token);
  const c = cfg_();
  mail_(c.admin_email, '[OKV RMS] Test email', '<p>This is a test email from the OKV RMS admin dashboard. If you can read this, e-mail alerts work.</p>');
  audit_('Test email sent', c.admin_email);
  return { to: c.admin_email, quota: MailApp.getRemainingDailyQuota() };
}

// ================================================================ SUPER ADMIN: approve, revoke
function adminApprove(token, id, note) {
  auth_(token);
  const rows = readAll_('Payments');
  let row = null;
  for (let i = 0; i < rows.length; i++) if (rows[i].ID === id) { row = rows[i]; break; }
  if (!row) throw new Error('Payment not found');
  if (row.Status === 'Rejected') throw new Error('This payment was rejected. Set it back to Pending first.');
  if (!validEmail_(row.Email)) throw new Error('This payment has no valid email address to send the code to.');
  if (row.Code) {
    mail_(row.Email, 'Your OKV RMS activation code', codeMailHtml_(row.Name, row.Code, note));
    audit_('Activation code re-sent', row.Code + ' to ' + row.Email);
    return { code: row.Code, emailed: true, resent: true };
  }
  const code = deliverCode_(row.Email, row.Name, row.Submission, row.Reference || row.ID, note);
  updateRow_('Payments', row._row, { Status: 'Verified', Code: code, Note: note != null && note !== '' ? clean_(note, 200) : row.Note, Updated: now_() });
  audit_('Payment approved, code sent', id + ' -> ' + code + ' to ' + row.Email);
  return { code: code, emailed: true, resent: false };
}
function adminRevokeCode(token, code) {
  auth_(token);
  const rows = readAll_('Codes');
  for (let i = 0; i < rows.length; i++) if (rows[i].Code === code) {
    updateRow_('Codes', rows[i]._row, { Status: 'Revoked' });
    audit_('Code revoked', code);
    return true;
  }
  throw new Error('Code not found');
}

// One-click approval straight from a submission (finds the payment record, or creates one for older requests)
function adminApproveSubmission(token, subId, note) {
  auth_(token);
  const pays = readAll_('Payments');
  for (let i = 0; i < pays.length; i++) if (pays[i].Submission === subId) return adminApprove(token, pays[i].ID, note);
  const subs = readAll_('Submissions');
  for (let i = 0; i < subs.length; i++) if (subs[i].ID === subId) {
    const s = subs[i];
    if (s.Type !== 'Activation Request') throw new Error('Only activation requests can be approved.');
    const id = nextId_('PAY');
    appendRow_('Payments', { ID: id, Timestamp: now_(), Method: 'Manual', Reference: s['Pay Ref'] || subId, Name: s.Name, Email: s.Email, Phone: s.Phone, Amount: s.Amount, Currency: 'NGN', Status: 'Pending', Submission: subId, 'Receipt URL': s['Receipt URL'], Updated: now_() });
    return adminApprove(token, id, note);
  }
  throw new Error('Submission not found');
}
