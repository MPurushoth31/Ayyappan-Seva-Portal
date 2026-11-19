/**
 * Ayyappan Seva Portal - Backend (Google Apps Script)
 * 1) Paste this file  2) Run setup() once  3) Deploy > Web app (Execute as: Me, Access: Anyone)
 * Existing Poojai / Silai_Booking data is kept: columns are matched by header name.
 */
const PROOF_FOLDER_ID = ''; // Drive folder ID for proof files ('' = auto-create "Silai_Proofs")
const MIMES = ['application/pdf', 'image/jpeg', 'image/png'];
const S = {
  poojai: { tab: 'Poojai', pre: 'POJ', cols: [['ID','id'],['Date','date'],['Location','location'],['Timing','timing'],['Guru','guru'],['Status','status'],['Created At','createdAt'],['Created By','by']], req: ['date','location','timing','guru'] },
  silai: { tab: 'Silai_Booking', pre: 'AYS', file: 1, cols: [['ID','id'],['Booking Date','bookingDate'],['Name','name'],['Location','location'],['Address','address'],['Phone','phone'],['Proof Type','proofType'],['Proof Number','proofNo'],['Status','status'],['Created At','createdAt'],['Proof File','proofFile'],['Created By','by']], req: ['bookingDate','name','location','address','proofType'] },
  users: { tab: 'Users', cols: [['Name','name'],['Password','password'],['Role','role'],['Active','active'],['Created At','createdAt']] },
  log: { tab: 'Activity_Log', cols: [['Time','time'],['User','user'],['Action','action'],['Record','record']] },
  settings: { tab: 'Settings', cols: [['Key','key'],['Value','value']] }
};

function setup() {
  Object.keys(S).forEach(sh_);
  const ss = SpreadsheetApp.getActiveSpreadsheet(), d = ss.getSheetByName('Sheet1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  getFolder_();
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', 'MGN');
}

// ---------- sheet helpers (header based) ----------
function sh_(t) {
  const c = S[t], ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(c.tab) || ss.insertSheet(c.tab);
  const lc = sh.getLastColumn();
  const hd = lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(String) : [];
  let add = false;
  c.cols.forEach(x => { if (hd.indexOf(x[0]) < 0) { hd.push(x[0]); add = true; } });
  if (add) {
    sh.getRange(1, 1, 1, hd.length).setValues([hd]).setFontWeight('bold').setBackground('#f4a300');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, 2000, hd.length).setNumberFormat('@');
  }
  const map = {}; c.cols.forEach(x => map[x[1]] = hd.indexOf(x[0]) + 1);
  return { sh: sh, map: map, n: hd.length };
}
function rows_(t) {
  const x = sh_(t), last = x.sh.getLastRow();
  if (last < 2) return [];
  const tz = tz_(), first = S[t].cols[0][1];
  return x.sh.getRange(2, 1, last - 1, x.n).getValues()
    .filter(r => r[x.map[first] - 1] !== '')
    .map(r => { const o = {}; for (const k in x.map) { const v = r[x.map[k] - 1]; o[k] = v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v); } return o; });
}
function find_(t, key, val) {
  const x = sh_(t), last = x.sh.getLastRow();
  if (last < 2) return 0;
  const v = x.sh.getRange(1, x.map[key], last, 1).getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(val)) return i + 1;
  return 0;
}
function tz_() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
function put_(t, o, row) {
  const x = sh_(t), r = row || x.sh.getLastRow() + 1, rng = x.sh.getRange(r, 1, 1, x.n);
  const cur = row ? rng.getValues()[0].map(v => v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : v) : new Array(x.n).fill('');
  for (const k in o) if (x.map[k]) cur[x.map[k] - 1] = o[k];
  rng.setNumberFormat('@'); // keep dates / phone numbers as plain text
  rng.setValues([cur]);
}
function nextId_(t) {
  const yr = new Date().getFullYear(), re = new RegExp('^' + S[t].pre + '-' + yr + '-(\\d+)$');
  let m = 0;
  rows_(t).forEach(r => { const k = re.exec(r.id); if (k) m = Math.max(m, +k[1]); });
  return S[t].pre + '-' + yr + '-' + ('000' + (m + 1)).slice(-4);
}
function log_(user, action, record) { put_('log', { time: new Date().toISOString(), user: user, action: action, record: record }); }

// ---------- drive ----------
function getFolder_() {
  if (PROOF_FOLDER_ID) return DriveApp.getFolderById(PROOF_FOLDER_ID);
  const it = DriveApp.getFoldersByName('Silai_Proofs');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Silai_Proofs');
}
function saveFile_(d) {
  if (MIMES.indexOf(d.fileMime) < 0) throw new Error('Only PDF, JPG, PNG allowed');
  if (d.fileData.length * 0.75 > 5.2 * 1024 * 1024) throw new Error('File must be under 5 MB');
  const safe = String(d.fileName || 'proof').replace(/[^\w.\-]/g, '_');
  const nm = String(d.name || 'booking').replace(/[^\w]/g, '_') + '_' + d.phone + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '_' + safe;
  return getFolder_().createFile(Utilities.newBlob(Utilities.base64Decode(d.fileData), d.fileMime, nm)).getUrl();
}

// ---------- API ----------
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  try {
    const s = rows_('settings').filter(r => r.key === 'options')[0];
    return out({ ok: true, poojai: rows_('poojai'), silai: rows_('silai'), opt: s ? JSON.parse(s.value) : null });
  } catch (e) { return out({ ok: false, error: String(e) }); }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const r = JSON.parse(e.postData.contents), cache = CacheService.getScriptCache();
    if (r.action === 'login') {
      const name = String(r.name || '').trim(), pw = String(r.password || '').trim();
      const master = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || 'MGN';
      let u = null;
      if ((!name || name.toLowerCase() === 'admin') && pw.toUpperCase() === String(master).toUpperCase()) u = { name: 'Admin', role: 'Admin' };
      else { const f = rows_('users').filter(x => x.name.toLowerCase() === name.toLowerCase() && x.password === pw && x.active !== 'No')[0]; if (f) u = { name: f.name, role: f.role }; }
      if (!u) return out({ ok: false, error: 'Wrong name or password' });
      const token = Utilities.getUuid();
      cache.put('tok_' + token, JSON.stringify(u), 21600);
      log_(u.name, 'LOGIN', '-');
      return out({ ok: true, token: token, name: u.name, role: u.role });
    }
    const u = JSON.parse(cache.get('tok_' + r.token) || 'null');
    if (!u) return out({ ok: false, auth: true, error: 'Session expired. Please login again.' });
    const adm = u.role === 'Admin';
    lock.waitLock(15000);

    if (r.action === 'add' || r.action === 'edit') {
      const t = r.type, c = S[t], d = r.data || {};
      if (!c || !c.pre) throw new Error('Invalid type');
      c.req.forEach(k => { if (!String(d[k] || '').trim()) throw new Error('Please fill all required fields'); });
      if (t === 'silai' && !/^\d{10}$/.test(String(d.phone || '').trim())) throw new Error('Phone must be 10 digits');
      const skip = ['id', 'createdAt', 'by', 'proofFile'], o = {};
      c.cols.forEach(x => { if (skip.indexOf(x[1]) < 0 && d[x[1]] !== undefined) o[x[1]] = String(d[x[1]]).trim(); });
      if (c.file && d.fileData) o.proofFile = saveFile_(d);
      if (r.action === 'add') {
        if (!o.status) o.status = t === 'poojai' ? 'Scheduled' : 'Pending';
        o.id = nextId_(t); o.createdAt = new Date().toISOString(); o.by = u.name;
        put_(t, o); log_(u.name, 'CREATE', o.id);
        return out({ ok: true, id: o.id });
      }
      const row = find_(t, 'id', d.id);
      if (!row) throw new Error('Record not found');
      put_(t, o, row); log_(u.name, 'UPDATE', d.id);
      return out({ ok: true });
    }
    if (r.action === 'delete') {
      if (!adm) throw new Error('Only Admin can delete');
      if (!S[r.type] || !S[r.type].pre) throw new Error('Invalid type');
      const row = find_(r.type, 'id', r.id);
      if (!row) throw new Error('Record not found');
      sh_(r.type).sh.deleteRow(row); log_(u.name, 'DELETE', r.id);
      return out({ ok: true });
    }
    if (!adm) throw new Error('Admin only');
    if (r.action === 'getAdmin') {
      return out({ ok: true, users: rows_('users').map(x => ({ name: x.name, role: x.role, active: x.active !== 'No' })), log: rows_('log').slice(-150).reverse() });
    }
    if (r.action === 'userSave') {
      const n = String(r.name || '').trim();
      if (!n || n.toLowerCase() === 'admin') throw new Error('Choose another user name');
      const row = find_('users', 'name', n), o = { name: n, role: r.role === 'Admin' ? 'Admin' : 'Editor' };
      if (r.active !== undefined) o.active = r.active ? 'Yes' : 'No';
      if (r.password) o.password = String(r.password);
      if (!row) { if (!o.password) throw new Error('Password required'); o.active = o.active || 'Yes'; o.createdAt = new Date().toISOString(); put_('users', o); }
      else put_('users', o, row);
      log_(u.name, 'USER', n);
      return out({ ok: true });
    }
    if (r.action === 'userDelete') {
      const row = find_('users', 'name', r.name);
      if (row) sh_('users').sh.deleteRow(row);
      log_(u.name, 'USER-DELETE', r.name);
      return out({ ok: true });
    }
    if (r.action === 'saveSettings') {
      const row = find_('settings', 'key', 'options'), o = { key: 'options', value: JSON.stringify(r.options) };
      put_('settings', o, row || 0); log_(u.name, 'SETTINGS', '-');
      return out({ ok: true });
    }
    return out({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return out({ ok: false, error: String(err.message || err) });
  } finally { try { lock.releaseLock(); } catch (x) {} }
}
