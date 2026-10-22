/* บันทึกขายอะไหล่ — หน้าบ้าน (ทำงานออฟไลน์ได้ แล้วส่งข้อมูลขึ้น Google Sheets เมื่อมีเน็ต) */
'use strict';

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
const HASH_ROUNDS = 300;
const NUMBER_BLOCK = 100;      // จองเลขบิลครั้งละกี่เลข
const NUMBER_LOW = 30;         // เหลือน้อยกว่านี้ให้จองเพิ่ม
const SYNC_EVERY_MS = 30000;

/* ================================================================ ตัวช่วย */

const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const num = (v) => { const n = Number(v); return isNaN(n) ? 0 : n; };
const round2 = (n) => Math.round(n * 100) / 100;
const money = (n) => num(n).toLocaleString('th-TH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const normCode = (c) => String(c || '').toUpperCase().replace(/[^0-9A-Z]/g, '');
const pad = (n, w) => String(n).padStart(w, '0');
function todayISO() { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`; }
function nowISO() { const d = new Date(); return `${todayISO()}T${pad(d.getHours(), 2)}:${pad(d.getMinutes(), 2)}:${pad(d.getSeconds(), 2)}`; }
function nowTime() { const d = new Date(); return `${pad(d.getHours(), 2)}:${pad(d.getMinutes(), 2)}`; }
function thDate(iso) { if (!iso) return ''; const [y, m, d] = String(iso).slice(0, 10).split('-'); return `${d}/${m}/${num(y) + 543}`; }
function uid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

async function sha256hex(s) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
  return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, '0')).join('');
}
async function hashPin(salt, pin) {
  let s = salt + ':' + pin;
  for (let i = 0; i < HASH_ROUNDS; i++) s = await sha256hex(s);
  return s;
}

// แปลงตัวเลขเป็นคำอ่านภาษาไทย (เหมือน BAHTTEXT ใน Excel)
function bahtText(amount) {
  const n = round2(Math.abs(num(amount)));
  const digits = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
  const units = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
  function read(intStr) {
    let out = '';
    if (intStr.length > 6) { out += read(intStr.slice(0, -6)) + 'ล้าน'; intStr = intStr.slice(-6); }
    const len = intStr.length;
    for (let i = 0; i < len; i++) {
      const d = num(intStr[i]), pos = len - i - 1;
      if (d === 0) continue;
      if (pos === 1 && d === 1) out += 'สิบ';
      else if (pos === 1 && d === 2) out += 'ยี่สิบ';
      else if (pos === 0 && d === 1 && len > 1 && intStr.slice(0, -1).replace(/0/g, '') !== '') out += 'เอ็ด';
      else out += digits[d] + units[pos];
    }
    return out;
  }
  const baht = Math.floor(n), satang = Math.round((n - baht) * 100);
  let s = '';
  if (baht > 0) s += read(String(baht)) + 'บาท';
  if (satang > 0) s += read(String(satang)) + 'สตางค์';
  else s += (baht > 0 ? '' : 'ศูนย์บาท') + 'ถ้วน';
  return (num(amount) < 0 ? 'ลบ' : '') + s;
}

/* ================================================================ IndexedDB */

const DB = {
  db: null,
  open() {
    return new Promise((res, rej) => {
      const r = indexedDB.open('parts-app', 2);
      r.onupgradeneeded = () => {
        const d = r.result, make = (name, opt) => { if (!d.objectStoreNames.contains(name)) d.createObjectStore(name, opt); };
        make('kv');
        make('parts', { keyPath: 'k' });
        make('customers', { keyPath: 'id' });
        make('bills', { keyPath: 'bill_no' });
        make('queue', { keyPath: 'op_id' });
        make('estimates', { keyPath: 'est_no' });   // v2: ใบประเมินราคา
      };
      r.onsuccess = () => { DB.db = r.result; res(); };
      r.onerror = () => rej(r.error);
    });
  },
  tx(store, mode, fn) {
    return new Promise((res, rej) => {
      const t = DB.db.transaction(store, mode);
      const s = t.objectStore(store);
      let out;
      Promise.resolve(fn(s)).then((v) => { out = v; });
      t.oncomplete = () => res(out);
      t.onerror = () => rej(t.error);
      t.onabort = () => rej(t.error);
    });
  },
  req(r) { return new Promise((res, rej) => { r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); },
  get(store, key) { return DB.tx(store, 'readonly', (s) => DB.req(s.get(key))); },
  all(store) { return DB.tx(store, 'readonly', (s) => DB.req(s.getAll())); },
  put(store, val, key) { return DB.tx(store, 'readwrite', (s) => { key === undefined ? s.put(val) : s.put(val, key); }); },
  putMany(store, vals) { return DB.tx(store, 'readwrite', (s) => { vals.forEach((v) => s.put(v)); }); },
  del(store, key) { return DB.tx(store, 'readwrite', (s) => { s.delete(key); }); },
  clear(store) { return DB.tx(store, 'readwrite', (s) => { s.clear(); }); },
  count(store) { return DB.tx(store, 'readonly', (s) => DB.req(s.count())); },
  kv: { get: (k) => DB.get('kv', k), set: (k, v) => DB.put('kv', v, k) }
};

/* ================================================================ สถานะแอป */

const S = {
  apiUrl: '', token: '', user: null, branches: [], settings: {}, approvers: [],
  branch: '', parts: new Map(), partsList: [], customers: [],
  cart: { lines: [] }, online: navigator.onLine, syncing: false, pending: 0, failed: 0,
  report: null, printBill: null
};
const isAdmin = () => S.user && S.user.role === 'admin';
const branchInfo = (code) => S.branches.find((b) => b.code === code) || { code, name: code, letter: '?' };

/* ================================================================ ติดต่อหลังบ้าน */

class NetError extends Error {}
// หน้าจอรอ: ขึ้นเมื่อผู้ใช้กดแล้วต้องรอระบบ (งานเบื้องหลัง เช่น ส่งบิล/โหลดอะไหล่ ไม่ต้องขึ้น)
const QUIET_ACTIONS = new Set(['sync', 'parts', 'customers', 'reserve', 'me']);
const BUSY_LABEL = {
  login: 'กำลังเข้าสู่ระบบ', bill: 'กำลังโหลดบิล', bills: 'กำลังโหลดรายงาน', estimates: 'กำลังโหลดใบประเมิน',
  settled: 'กำลังโหลดใบคุม', credits: 'กำลังโหลดใบคุม', audit: 'กำลังโหลดประวัติ', users: 'กำลังโหลดรายชื่อผู้ใช้',
  suppliers: 'กำลังโหลดรายชื่อร้าน', pos: 'กำลังโหลดใบสั่งซื้อ', device: 'กำลังตรวจเลขเครื่อง',
  payPO: 'กำลังบันทึกการจ่ายเงิน', voidPO: 'กำลังยกเลิก'
};
let busyCount = 0, busyTimer = null, slowTimer = null;
function busy(on, label) {
  const el = $('#busy');
  if (!el) return;
  if (on) $('.busy-what', el).textContent = label || 'กำลังทำงาน';
  busyCount = Math.max(0, busyCount + (on ? 1 : -1));
  if (busyCount === 1 && on) {
    // รอเกิน 0.3 วินาทีค่อยขึ้น จะได้ไม่กระพริบตอนเสร็จเร็ว แต่กันกดซ้ำทันที
    document.body.style.pointerEvents = 'none';
    busyTimer = setTimeout(() => { el.hidden = false; }, 300);
    slowTimer = setTimeout(() => { $('.busy-slow', el).hidden = false; }, 10000);
  } else if (!busyCount) {
    clearTimeout(busyTimer); clearTimeout(slowTimer);
    el.hidden = true; $('.busy-slow', el).hidden = true;
    document.body.style.pointerEvents = '';
  }
}

async function api(action, body = {}, label) {
  if (QUIET_ACTIONS.has(action)) return apiCall(action, body);
  busy(true, label || BUSY_LABEL[action] || (/^save/.test(action) ? 'กำลังบันทึก' : ''));
  try { return await apiCall(action, body); } finally { busy(false); }
}

async function apiCall(action, body) {
  if (!S.apiUrl) throw new NetError('ยังไม่ได้ตั้งค่าลิงก์ระบบ');
  let res;
  try {
    res = await fetch(S.apiUrl, {
      method: 'POST', redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({ action, token: S.token }, body))
    });
  } catch (e) {
    setOnline(false);
    throw new NetError('ติดต่อระบบไม่ได้ (ไม่มีอินเทอร์เน็ต?)');
  }
  if (!res.ok) { setOnline(false); throw new NetError('ระบบตอบกลับผิดพลาด ' + res.status); }
  const j = await res.json();
  setOnline(true);
  if (!j.ok) {
    if (/เข้าสู่ระบบ/.test(j.error)) { toast(j.error); }
    throw new Error(j.error);
  }
  return j.data;
}

function setOnline(v) {
  if (S.online === v) return;
  S.online = v;
  renderSync();
}

/* ================================================================ หน้าจอ */

function show(id) { $$('.screen').forEach((s) => s.classList.add('hidden')); $('#' + id).classList.remove('hidden'); }

function toast(msg, ms = 3500) {
  let t = $('#toast');
  if (!t) {
    t = document.createElement('div'); t.id = 'toast'; t.className = 'no-print';
    Object.assign(t.style, { position: 'fixed', bottom: '18px', left: '50%', transform: 'translateX(-50%)', background: '#222', color: '#fff',
      padding: '10px 18px', borderRadius: '8px', zIndex: 99, maxWidth: '90vw' });
    document.body.appendChild(t);
  }
  t.textContent = msg; t.style.display = 'block';
  clearTimeout(t._h); t._h = setTimeout(() => { t.style.display = 'none'; }, ms);
}

function openModal(html) {
  $('#modal-box').innerHTML = html;
  $('#modal').classList.remove('hidden');
  const f = $('#modal-box [autofocus]') || $('#modal-box input');
  if (f) setTimeout(() => f.focus(), 30);
  return $('#modal-box');
}
function closeModal() { $('#modal').classList.add('hidden'); $('#modal-box').innerHTML = ''; }

/* ================================================================ เริ่มต้น */

// หน้า index.html กับ app.js ต้องเป็นรุ่นเดียวกัน ถ้าเบราว์เซอร์ยังจำหน้าเก่าไว้ ให้โหลดใหม่เอง (ไม่เกิน 3 ครั้ง)
const APP_BUILD = '37';
// โลโก้ร้านฝังไว้ในโปรแกรม (ไม่ต้องมีไฟล์รูปแยก)
const LOGO = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/6xaJSlACEQAAAAEAABZ/anVtYgAAAB5qdW1kYzJwYQARABCAAACqADibcQNjMnBhAAAAFllqdW1iAAAAR2p1bWRjMm1hABEAEIAAAKoAOJtxA3VybjpjMnBhOjFmOWQ4MTdlLTBjZjktNDZiMy05YmQzLWNiYmJmMWM0OTg2MwAAAAOUanVtYgAAAClqdW1kYzJhcwARABCAAACqADibcQNjMnBhLmFzc2VydGlvbnMAAAAAuWp1bWIAAABEanVtZGNib3IAEQAQgAAAqgA4m3ETYzJwYS5pbmdyZWRpZW50LnYzAAAAABhjMnNoOMP/6VrtiQ0haBwQm//9lQAAAG1jYm9yo2lkYzpmb3JtYXRqaW1hZ2UvanBlZ2ppbnN0YW5jZUlEeCx4bXA6aWlkOjg2YjBhNjE3LWRiZjctNDdlOS05NjA3LTE2MDcyODIzNzkwYmxyZWxhdGlvbnNoaXBocGFyZW50T2YAAAHianVtYgAAAEFqdW1kY2JvcgARABCAAACqADibcRNjMnBhLmFjdGlvbnMudjIAAAAAGGMyc2gZjFNYNytUUHi3diaCx5OzAAABmWNib3KiZ2FjdGlvbnOComZhY3Rpb25rYzJwYS5vcGVuZWRqcGFyYW1ldGVyc6FraW5ncmVkaWVudHOBomN1cmx4LXNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuaW5ncmVkaWVudC52M2RoYXNoWCAg2J9GwXh7G2es8654UOFNYYlQzcDflv6b6H99IEz7J6RmYWN0aW9ueB1jb20uYW50aHJvcGljLmNsYXVkZS5wcm92aWRlZGpwYXJhbWV0ZXJzoXgfY29tLmFudGhyb3BpYy5vcmlnaW4tY29uZmlkZW5jZWd1bmtub3dua2Rlc2NyaXB0aW9ueGZDbGF1ZGUgcHJvdmlkZWQgdGhpcyBmaWxlIGF0IHRoZSByZXF1ZXN0IG9mIGEgdXNlciBhbmQgbWF5IGhhdmUgY3JlYXRlZCBvciBtb2RpZmllZCB0aGUgZmlsZSBjb250ZW50cy5tc29mdHdhcmVBZ2VudKFkbmFtZWZDbGF1ZGVyYWxsQWN0aW9uc0luY2x1ZGVk9QAAAMhqdW1iAAAAQGp1bWRjYm9yABEAEIAAAKoAOJtxE2MycGEuaGFzaC5kYXRhAAAAABhjMnNo9k7BnVtM6C6uOgA4de2k+AAAAIBjYm9ypWNhbGdmc2hhMjU2Y3BhZE4AAAAAAAAAAAAAAAAAAGRoYXNoWCCKPAeIqnebPOxi0Nf0c6mkHCzdR2lLGCJAhIOeQ1Br8mRuYW1lbmp1bWJmIG1hbmlmZXN0amV4Y2x1c2lvbnOBomVzdGFydBRmbGVuZ3RoGRaLAAACPmp1bWIAAAAnanVtZGMyY2wAEQAQgAAAqgA4m3EDYzJwYS5jbGFpbS52MgAAAAIPY2JvcqVjYWxnZnNoYTI1NmlzaWduYXR1cmV4TXNlbGYjanVtYmY9L2MycGEvdXJuOmMycGE6MWY5ZDgxN2UtMGNmOS00NmIzLTliZDMtY2JiYmYxYzQ5ODYzL2MycGEuc2lnbmF0dXJlamluc3RhbmNlSUR4LHhtcDppaWQ6NmRiNmY3ZmUtMzUwOS00MGJhLWIxMTMtOTg5MTBmMmUxMmFkcmNyZWF0ZWRfYXNzZXJ0aW9uc4OiY3VybHgtc2VsZiNqdW1iZj1jMnBhLmFzc2VydGlvbnMvYzJwYS5pbmdyZWRpZW50LnYzZGhhc2hYICDYn0bBeHsbZ6zzrnhQ4U1hiVDNwN+W/pvof30gTPsnomN1cmx4KnNlbGYjanVtYmY9YzJwYS5hc3NlcnRpb25zL2MycGEuYWN0aW9ucy52MmRoYXNoWCA0nGrGMsW0+kSQ79AbsSE+ZiIVZqxtlrCfp4ArGnSoBKJjdXJseClzZWxmI2p1bWJmPWMycGEuYXNzZXJ0aW9ucy9jMnBhLmhhc2guZGF0YWRoYXNoWCBD8gytc3YUWdmaqIphGf/++bUfxlgFwn7eQq+T2FBVT3RjbGFpbV9nZW5lcmF0b3JfaW5mb6NkbmFtZW9BbnRocm9waWMgRmlsZXNndmVyc2lvbmUxLjAuMGtzcGVjVmVyc2lvbmUyLjQuMAAAEDhqdW1iAAAAKGp1bWRjMmNzABEAEIAAAKoAOJtxA2MycGEuc2lnbmF0dXJlAAAAEAhjYm9y0oRZAhKiASYYIVkCCjCCAgYwggGNoAMCAQICFEDloAruwjnQvriD+gZCBT1nVRMAMAoGCCqGSM49BAMDMEkxFzAVBgNVBAoTDkFudGhyb3BpYywgUEJDMS4wLAYDVQQDEyVBbnRocm9waWMgQ29udGVudCBDcmVkZW50aWFscyBSb290IENBMB4XDTI2MDgwNzE4NDM1NloXDTI4MDgwNjE5NDM1NlowRDEXMBUGA1UEChMOQW50aHJvcGljLCBQQkMxKTAnBgNVBAMTIEFudGhyb3BpYyBDbGF1ZGUgQ29udGVudCBTaWduaW5nMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAEmHoKa8tQGAUU1TS9QqU5W0Tp2N3XsvlK7BfQt6YWKwEzd2R3/dzKPEUDdCjlLjp9fT+KFjRVnuZ9v0oXvTe3k6NYMFYwDgYDVR0PAQH/BAQDAgeAMBUGA1UdJQQOMAwGCisGAQQBg+heAgEwDAYDVR0TAQH/BAIwADAfBgNVHSMEGDAWgBTOUeIEgU5kWyP448TPmj6cwddcwjAKBggqhkjOPQQDAwNnADBkAjAxcx0UngF60stVjs5G4T2eiptsBk5mf9oCtfJPAUBl8qs/PEXa8+gk1/X5QJ2DVcYCMHBfXN31YapiSqYvlIWrDVDJKOvXMl+kkz37Wt0PBI8sw486Mq6JeOhT+lRR4b1HCaFjcGFkWQ2eAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA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';
function buildMismatch() {
  const m = document.querySelector('meta[name="app-build"]');
  if (m && m.content === APP_BUILD) { try { sessionStorage.removeItem('reloadTry'); } catch (e) { /* */ } return false; }
  let n = 0;
  try { n = Number(sessionStorage.getItem('reloadTry') || 0); sessionStorage.setItem('reloadTry', n + 1); } catch (e) { n = 3; }
  document.body.innerHTML = '<p style="padding:24px;font-size:18px">กำลังอัปเดตแอปเป็นรุ่นใหม่ กรุณารอสักครู่…</p>'
    + (n >= 3 ? '<p style="padding:0 24px;color:#b3261e">ถ้าหน้านี้ไม่หายไป ให้กด Ctrl+Shift+R หรือปิดแท็บแล้วเปิดลิงก์ใหม่</p>' : '');
  if (n < 3) setTimeout(() => location.replace(location.pathname + '?r=' + Date.now()), 1500);
  return true;
}

async function boot() {
  if (buildMismatch()) { if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {}); return; }
  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
  await DB.open();
  const cfgUrl = window.APP_CONFIG && window.APP_CONFIG.API_URL;
  if (cfgUrl) await DB.kv.set('apiUrl', cfgUrl);
  S.apiUrl = cfgUrl || (await DB.kv.get('apiUrl')) || '';
  window.addEventListener('online', () => { setOnline(true); syncNow(); });
  window.addEventListener('offline', () => setOnline(false));
  bindStatic();
  if (!S.apiUrl) return show('scr-setup');
  const sess = await DB.kv.get('session');
  if (sess && sess.token) { await startSession(sess); return; }
  showLogin();
}

function showLogin() {
  show('scr-login');
  $('#login-net').textContent = navigator.onLine ? '' : 'ตอนนี้ออฟไลน์: เข้าได้เฉพาะผู้ใช้ที่เคยเข้าระบบในเครื่องนี้';
  $('#login-user').focus();
}

async function doLogin(username, pin) {
  username = username.trim();
  const users = (await DB.kv.get('users')) || {};
  try {
    const d = await api('login', { username, pin });
    users[d.user.username.toLowerCase()] = { salt: d.verifier.salt, hash: d.verifier.hash, token: d.token, boot: d };
    await DB.kv.set('users', users);
    return d;
  } catch (e) {
    if (!(e instanceof NetError)) throw e;
    // ออฟไลน์: ตรวจรหัสกับข้อมูลที่เครื่องจำไว้
    const u = users[username.toLowerCase()];
    if (!u) throw new Error(navigator.onLine ? 'ติดต่อระบบหลังบ้านไม่ได้: ตรวจลิงก์ Web app ใน config.js และตั้งผู้มีสิทธิ์เข้าถึงเป็น "ทุกคน"' : 'ออฟไลน์อยู่ และผู้ใช้นี้ยังไม่เคยเข้าระบบในเครื่องนี้');
    if ((await hashPin(u.salt, pin)) !== u.hash) throw new Error('ชื่อผู้ใช้หรือรหัสไม่ถูกต้อง');
    return Object.assign({}, u.boot, { token: u.token });
  }
}

async function startSession(sess) {
  S.token = sess.token; S.user = sess.user; S.branches = sess.branches || []; S.settings = sess.settings || {}; S.approvers = sess.approvers || [];
  S.branch = (await DB.kv.get('branch:' + S.user.username)) || S.user.branch;
  if (!isAdmin()) S.branch = S.user.branch;
  show('scr-app');
  $$('.admin-only').forEach((el) => el.classList.toggle('hidden', !isAdmin()));
  $('#user-label').textContent = S.user.name + (isAdmin() ? ' (แอดมิน)' : '');
  $('#shop-name').textContent = S.settings.shop_name || '';
  fillBranchSelects();
  renderSync();
  await loadPartsFromDb();
  await loadCustomers();
  newSale();
  await refreshPending();
  renderSync();
  // ทำงานเบื้องหลัง: อัปเดตข้อมูลจากหลังบ้าน
  (async () => {
    try {
      const me = await api('me');
      Object.assign(sess, { user: me.user, branches: me.branches, settings: me.settings, approvers: me.approvers });
      S.branches = me.branches; S.settings = me.settings; S.approvers = me.approvers;
      await DB.kv.set('session', sess);
      const users = (await DB.kv.get('users')) || {};
      const key = S.user.username.toLowerCase();
      if (users[key]) { Object.assign(users[key].boot, me); await DB.kv.set('users', users); }
      fillBranchSelects();
    } catch (e) { if (!(e instanceof NetError)) toast('อัปเดตข้อมูลผู้ใช้ไม่สำเร็จ: ' + e.message, 8000); }
    await syncParts();
    await syncCustomers();
    await syncNow();
  })();
  clearInterval(S._timer);
  S._timer = setInterval(syncNow, SYNC_EVERY_MS);
}

function fillBranchSelects() {
  const opts = S.branches.map((b) => `<option value="${esc(b.code)}">${esc(b.name)} (${esc(b.letter)})</option>`).join('');
  const sel = $('#branch-select');
  sel.innerHTML = opts; sel.value = S.branch; sel.disabled = !isAdmin();
  $('#rep-branch').innerHTML = '<option value="*">ทุกสาขา</option>' + opts;
  $('#rep-branch').value = S.branch;
  $('#ctl-branch').innerHTML = opts; $('#ctl-branch').value = S.branch;
  $('#po-branch').innerHTML = opts; $('#po-branch').value = S.branch;
  $('#user-branch').innerHTML = opts;
}

/* ================================================================ ข้อมูลอะไหล่ */

const codeKey = (c) => String(c || '').trim().toUpperCase();

async function loadPartsFromDb() {
  const rows = await DB.all('parts');
  S.parts = new Map(rows.map((p) => [p.k, p]));
  // ค้นแบบไม่สนขีด/ช่องว่าง เช่น 51490KGH901 = 51490-KGH-901 (ถ้าซ้ำกันใช้ตัวแรก)
  S.partsNorm = new Map();
  rows.forEach((p) => { const n = normCode(p.code); if (n && !S.partsNorm.has(n)) S.partsNorm.set(n, p); });
  S.partsList = rows;
}

function partFromRow(r) {
  return { k: codeKey(r[0]), code: r[0], name: r[1], model: r[2], price: r[3], cost: r[4], u: r[5],
    maxDisc: r[6] == null ? '' : r[6], status: r[7] || '', addedBy: r[8] || '', noPoint: !!r[9] };
}
const PART_PENDING = 'รอตรวจ';

/* ---------------------------------------------------------------- ส่วนลด */
// % ส่วนลดของสาขาที่กำลังขาย (แอดมินตั้งในแท็บ แอดมิน > ส่วนลด)
function branchPct(code = S.branch) { const b = S.branches.find((x) => x.code === code); return num(b && b.discount_pct); }
// เพดานส่วนลดของอะไหล่แต่ละตัว: ตามสาขา หรือต่ำกว่าถ้าแอดมินตั้งไว้, อะไหล่รอตรวจลดไม่ได้
function partCap(code, pct = branchPct()) {
  const p = findPart(code);
  if (!p || p.status === PART_PENDING) return 0;
  return p.maxDisc === '' || p.maxDisc == null ? pct : Math.min(pct, num(p.maxDisc));
}
function allowedDiscount(lines) { return round2(lines.reduce((s, l) => s + num(l.qty) * num(l.unit) * partCap(l.code) / 100, 0)); }

async function syncParts() {
  const since = await DB.kv.get('partsSince');
  const count = S.partsList.length;
  const bar = $('#parts-progress');
  try {
    if (!since || !count) {
      // ดาวน์โหลดทั้งหมดครั้งแรก
      let offset = 0, total = 0, maxU = '';
      bar.classList.remove('hidden');
      for (;;) {
        bar.textContent = `กำลังโหลดข้อมูลอะไหล่ลงเครื่อง ${offset.toLocaleString()} / ${total ? total.toLocaleString() : '...'} รายการ (ครั้งแรกครั้งเดียว)`;
        const d = await api('parts', { offset, limit: 10000 });
        total = d.total;
        const items = d.rows.map(partFromRow);
        items.forEach((p) => { if (p.u > maxU) maxU = p.u; });
        await DB.putMany('parts', items);
        offset += 10000;
        if (d.done) break;
      }
      await DB.kv.set('partsSince', maxU || '0');
      bar.classList.add('hidden');
    } else {
      const d = await api('parts', { since });
      const gone = d.deleted || [];
      if (d.rows.length || gone.length) {
        const items = d.rows.map(partFromRow);
        let maxU = since;
        items.forEach((p) => { if (p.u > maxU) maxU = p.u; });
        gone.forEach(([, at]) => { if (at > maxU) maxU = at; });
        // ลบก่อนแล้วค่อยใส่ใหม่ (รหัสที่ถูกลบแล้วเพิ่มกลับมาทีหลังจะยังอยู่)
        for (const [code] of gone) await DB.del('parts', codeKey(code));
        await DB.putMany('parts', items);
        await DB.kv.set('partsSince', maxU);
      }
    }
    await loadPartsFromDb();
  } catch (e) {
    bar.classList.add('hidden');
    if (!(e instanceof NetError)) toast('โหลดอะไหล่ไม่สำเร็จ: ' + e.message);
  }
}

function findPart(code) { return S.parts.get(codeKey(code)) || (normCode(code) ? S.partsNorm.get(normCode(code)) : undefined); }

function searchParts(q, limit = 80) {
  q = q.trim().toLowerCase();
  if (!q) return [];
  const words = q.split(/\s+/);
  const nq = normCode(q);
  const out = [];
  for (const p of S.partsList) {
    const hay = (p.code + ' ' + p.name + ' ' + p.model).toLowerCase();
    if ((nq.length >= 3 && normCode(p.code).includes(nq)) || words.every((w) => hay.includes(w))) {
      out.push(p);
      if (out.length >= limit) break;
    }
  }
  return out;
}

/* ================================================================ ลูกค้า */

async function loadCustomers() {
  const all = await DB.all('customers');
  S.customers = all.filter((c) => c.branch === S.branch);
  // ตัวเลือกแสดงทั้งรหัสและชื่อ พิมพ์ส่วนไหนก็ค้นเจอ
  $('#customer-list').innerHTML = S.customers.map((c) => `<option value="${esc(customerLabel(c))}">${esc(c.phone || '')}</option>`).join('');
  renderCustomers();
}

async function syncCustomers() {
  try {
    const rows = await api('customers', { branch: S.branch });
    // แทนที่ข้อมูลสาขานี้ทั้งหมด ยกเว้นที่ยังรอส่ง
    const local = await DB.all('customers');
    const pendingIds = new Set(local.filter((c) => c._pending).map((c) => c.id));
    await DB.tx('customers', 'readwrite', (s) => {
      local.filter((c) => c.branch === S.branch && !c._pending).forEach((c) => s.delete(c.id));
      (rows || []).forEach((c) => { if (!pendingIds.has(c.id)) s.put(c); });
    });
    await loadCustomers();
  } catch (e) { if (!(e instanceof NetError)) toast('โหลดรายชื่อลูกค้าไม่สำเร็จ: ' + e.message, 8000); }
}

// รหัสลูกค้า = ส่วนหลังรหัสสาขาใน id เช่น HQ-0003 -> 0003 (พิมพ์ 3 ก็หาเจอ)
function customerCode(c) { return String(c.id || '').split('-').slice(1).join('-').replace(/^0+(?=\d)/, ''); }
const codeNorm = (v) => String(v || '').trim().toUpperCase().replace(/^0+(?=\d)/, '');
function customerLabel(c) { const code = customerCode(c); return code ? `${code} · ${c.display}` : c.display; }
function customerByName(name) {
  name = String(name || '').trim();
  return S.customers.find((c) => customerLabel(c) === name) || S.customers.find((c) => c.display === name)
    || S.customers.find((c) => codeNorm(customerCode(c)) === codeNorm(name));
}
// ชื่อที่จะพิมพ์ลงบิล: ถ้าเป็นลูกค้าในระบบใช้ชื่อในระบบ ถ้าพิมพ์ชื่อใหม่ใช้ตามที่พิมพ์
function customerNameFor(text) { const c = customerByName(text); return c ? c.display : String(text || '').trim(); }
// สมาชิกสะสมแต้ม: ติ๊กในข้อมูลลูกค้า เบอร์มือถือใช้เป็นรหัสในแอปสะสมแต้ม
function isMember(c) { return !!c && (c.points === true || /^(TRUE|Y|1)$/i.test(String(c.points == null ? '' : c.points).trim())); }
function mobileOf(phone) { const m = String(phone || '').replace(/\D/g, ''); return /^0[689]\d{8}$/.test(m) ? m : ''; }
function customerFull(c) { return c ? [c.title, c.first, c.last].filter(Boolean).join(' ').trim() : ''; }

function renderCustomers() {
  const q = ($('#cus-search').value || '').toLowerCase();
  const rows = S.customers.filter((c) => !q || JSON.stringify(c).toLowerCase().includes(q)).slice(0, 300);
  $('#cus-table tbody').innerHTML = rows.map((c) => `<tr><td>${esc(customerCode(c))}</td>
    <td>${esc(c.display)} ${c._pending ? '<span class="badge pending">รอส่ง</span>' : ''}</td><td>${esc(customerFull(c))}</td>
    <td>${esc(c.phone)}</td><td>${isMember(c) ? '<span class="badge">สมาชิก</span>' : ''}</td><td>${esc(c.address)}</td><td><button data-cus="${esc(c.id)}">แก้ไข</button></td></tr>`).join('');
}

function customerForm(c) {
  const isNew = !c;
  c = c || { branch: S.branch };
  const box = openModal(`<h2>${isNew ? 'เพิ่มลูกค้า' : 'แก้ไขลูกค้า'}</h2>
    <form class="form-grid" id="cus-form">
      <label>รหัสลูกค้า *<input name="code" required ${isNew ? 'autofocus' : 'readonly'} value="${esc(isNew ? '' : customerCode(c))}" placeholder="เช่น 0431"></label>
      <label>ชื่อที่ใช้ในบิล *<input name="display" required value="${esc(c.display)}"></label>
      <label>คำนำหน้า<input name="title" value="${esc(c.title)}"></label>
      <label>ชื่อ<input name="first" value="${esc(c.first)}"></label>
      <label>นามสกุล<input name="last" value="${esc(c.last)}"></label>
      <label>เบอร์โทร<input name="phone" value="${esc(c.phone)}"></label>
      <label>เลขบัตรประชาชน<input name="idcard" value="${esc(c.idcard)}"></label>
      <label class="check"><input name="points" type="checkbox" ${isMember(c) ? 'checked' : ''}> สมาชิกสะสมแต้ม (ต้องมีเบอร์มือถือ 10 หลัก)</label>
      <label style="grid-column:1/-1">ที่อยู่<input name="address" value="${esc(c.address)}"></label>
      <div class="actions"><button type="button" data-close>ยกเลิก</button><button class="primary" type="submit">บันทึก</button></div>
    </form>`);
  $('#cus-form', box).onsubmit = async (ev) => {
    ev.preventDefault();
    const f = Object.fromEntries(new FormData(ev.target));
    const code = String(f.code || '').trim().replace(/[^0-9A-Za-z._-]/g, '');
    delete f.code;
    if (!code) return toast('ใส่รหัสลูกค้าเป็นตัวเลขหรือตัวอักษรอังกฤษ');
    if (isNew && S.customers.some((x) => codeNorm(customerCode(x)) === codeNorm(code))) return toast('รหัสลูกค้านี้มีอยู่แล้ว');
    f.points = !!f.points;
    if (f.points && !mobileOf(f.phone)) return toast('สมาชิกสะสมแต้มต้องมีเบอร์มือถือ 10 หลัก เช่น 081-234-5678');
    const rec = Object.assign({}, c, f, isNew ? { id: S.branch + '-' + code } : {});
    if (S.customers.some((x) => x.display === rec.display && x.id !== rec.id)) return toast('ชื่อนี้มีอยู่แล้ว');
    delete rec._pending;
    await enqueue({ type: 'saveCustomer', customer: rec });
    rec._pending = true;
    await DB.put('customers', rec);
    closeModal();
    await loadCustomers();
    toast('บันทึกลูกค้าแล้ว');
  };
}

/* ================================================================ เลขที่บิล */

// เลขที่จองไว้ต้องขึ้นต้นตามรูปแบบปัจจุบัน (เช่น PA69) ถ้าเปลี่ยนรูปแบบ เลขเก่าในเครื่องจะถูกทิ้ง
// kind 'est' = เลขใบประเมินราคา (Q + ตัวอักษรสาขา) จองแยกจากเลขบิล
// เลขบิลแยกตามเครื่อง เช่น PA1-00001 = สำนักงานใหญ่ เครื่องที่ 1 (ใบประเมินยังเป็น QA + ปี + เลข)
function numberPrefix(branch, kind, dev) {
  const pre = kind === 'est' ? 'Q' : S.settings && 'bill_prefix' in S.settings ? String(S.settings.bill_prefix).trim().toUpperCase() : 'P';
  return pre + branchInfo(branch).letter + (kind === 'est' ? '' : dev + '-');
}
const numbersKey = (branch, kind) => (kind === 'est' ? 'estnums:' : 'numbers:') + branch;
async function numberBlocks(branch, kind) {
  if (kind === 'est') {
    const head = numberPrefix(branch, kind);
    return ((await DB.kv.get(numbersKey(branch, kind))) || []).filter((b) => b.prefix.indexOf(head) === 0 && /^\d{2}$/.test(b.prefix.slice(head.length)));
  }
  const dev = await DB.kv.get('deviceNo:' + branch);
  if (!dev) return [];
  const head = numberPrefix(branch, kind, dev);
  return ((await DB.kv.get(numbersKey(branch, kind))) || []).filter((b) => b.prefix === head);
}

// รหัสประจำเครื่อง (สุ่มครั้งแรก) ใช้ขอเลขเครื่องจากระบบกลาง
async function deviceId() {
  let id = await DB.kv.get('deviceId');
  if (!id) { id = uid() + uid(); await DB.kv.set('deviceId', id); }
  return id;
}
async function ensureDevice(branch) {
  const dev = await DB.kv.get('deviceNo:' + branch);
  if (dev) return dev;
  const r = await api('device', { branch, device_id: await deviceId() });
  await DB.kv.set('deviceNo:' + branch, r.device_no);
  return r.device_no;
}
async function showDevice() {
  const dev = await DB.kv.get('deviceNo:' + S.branch);
  $('#sale-dev').textContent = dev ? 'เครื่องที่ ' + dev : '';
}
// เปลี่ยนเลขเครื่อง เช่น เปลี่ยนคอมใหม่แทนเครื่องที่ 1 (ต้องใช้รหัสอนุมัติแอดมิน) เลขบิลจะรันต่อจากเลขสุดท้ายของเครื่องนั้น
async function changeDevice() {
  const cur = await DB.kv.get('deviceNo:' + S.branch);
  const v = prompt(`เครื่องนี้เป็นเครื่องที่ ${cur || '-'} ของ${branchInfo(S.branch).name}\nต้องการเปลี่ยนเป็นเครื่องที่ (1-99)\n(ใส่เลขเดิม = ทิ้งเลขบิลที่จองไว้ แล้วรันต่อจากเลขในชีต Counters)`, cur || '');
  if (v == null || String(v).trim() === '') return;
  const n = Number(v);
  if (!(n >= 1 && n <= 99 && Math.floor(n) === n)) return toast('เลขเครื่องต้องเป็น 1 ถึง 99');
  let approval;
  const same = n === Number(cur);
  if (!isAdmin()) { const ap = await askApproval(same ? 'จองเลขบิลใหม่ของเครื่องที่ ' + n : 'เปลี่ยนเลขเครื่องเป็น ' + n, 'บิลใหม่ของเครื่องนี้จะใช้หัวเลขของเครื่องที่ ' + n, false); if (!ap) return; approval = ap.approval; }
  try {
    const r = await api('device', { branch: S.branch, device_id: await deviceId(), device_no: n, approval });
    await DB.kv.set('deviceNo:' + S.branch, r.device_no);
    await withNumbers(() => DB.kv.set(numbersKey(S.branch), []));   // ทิ้งเลขที่จองไว้ของหัวเดิม
    // ล้างสำเนาบิลที่ส่งเข้าระบบแล้วในเครื่อง (ของจริงอยู่ในชีต) เผื่อมีการแก้เลขบิลในชีต จะได้ไม่ขึ้นบิลเลขเก่าซ้ำในรายงาน
    for (const b of await DB.all('bills')) if (b._sync === 'synced') await DB.del('bills', b.bill_no);
    toast(same ? 'จองเลขบิลใหม่ของเครื่องที่ ' + r.device_no + ' แล้ว' : 'ตั้งเป็นเครื่องที่ ' + r.device_no + ' แล้ว');
    if (!S.cart || !S.cart.lines.length) await newSale(); else await showDevice();
  } catch (e) { toast(e.message, 6000); }
}
function remaining(blocks) { return blocks.reduce((s, b) => s + (b.end - b.next + 1), 0); }

// งานจอง/ใช้เลขที่ทำทีละงาน ไม่ให้ซ้อนกัน (เคยซ้อนกันแล้วเขียนทับชุดเลขที่เพิ่งจอง เลขบิลเลยกระโดด เช่น 00300 -> 00701)
let numberQueue = Promise.resolve();
function withNumbers(fn) {
  // navigator.locks กันได้ข้ามแท็บด้วย (เปิดโปรแกรมไว้หลายแท็บในเครื่องเดียวกัน)
  const job = navigator.locks ? () => navigator.locks.request('parts-app-numbers', fn) : fn;
  const run = numberQueue.then(job, job);
  numberQueue = run.catch(() => {});
  return run;
}

function ensureNumbers(branch, kind) { return withNumbers(() => reserveIfLow(branch, kind)); }

async function reserveIfLow(branch, kind) {
  const blocks = await numberBlocks(branch, kind);
  if (remaining(blocks) >= (kind === 'est' ? 10 : NUMBER_LOW)) return;
  try {
    const device = kind === 'est' ? undefined : await ensureDevice(branch);
    const r = await api('reserve', { branch, count: kind === 'est' ? 30 : NUMBER_BLOCK, kind, device });
    blocks.push({ prefix: r.prefix, next: r.start, end: r.end });
    await DB.kv.set(numbersKey(branch, kind), blocks);
  } catch (e) { if (!(e instanceof NetError)) toast('จองเลขที่ไม่สำเร็จ: ' + e.message); }
}

async function peekNumber(branch, kind) {
  const b = (await numberBlocks(branch, kind)).find((x) => x.next <= x.end);
  return b ? b.prefix + pad(b.next, 5) : null;
}

function takeNumber(branch, kind) { return withNumbers(() => takeNumberNow(branch, kind)); }

async function takeNumberNow(branch, kind) {
  let blocks = await numberBlocks(branch, kind);
  if (!blocks.some((x) => x.next <= x.end)) { await reserveIfLow(branch, kind); blocks = await numberBlocks(branch, kind); }
  const b = blocks.find((x) => x.next <= x.end);
  if (!b) throw new Error('เลขที่สำรองในเครื่องหมดแล้ว กรุณาต่ออินเทอร์เน็ตสักครู่');
  const no = b.prefix + pad(b.next, 5);
  b.next++;
  await DB.kv.set(numbersKey(branch, kind), blocks.filter((x) => x.next <= x.end));
  return no;
}

/* ================================================================ คิวส่งข้อมูล */

async function enqueue(op) {
  op.op_id = op.op_id || uid() + '-' + S.user.username;
  op.seq = Date.now() + Math.random();
  op.state = 'pending';
  await DB.put('queue', op);
  await refreshPending();
  syncNow();
  return op;
}

async function refreshPending() {
  const q = await DB.all('queue');
  S.pending = q.filter((o) => o.state === 'pending').length;
  S.failed = q.filter((o) => o.state === 'error').length;
  renderSync();
}

function renderSync() {
  const c = $('#sync-chip');
  if (!c) return;
  c.className = 'chip ' + (!S.online ? 'offline' : S.pending || S.failed ? 'pending' : 'online');
  let t = S.online ? 'ออนไลน์' : 'ออฟไลน์';
  if (S.pending) t += ` · รอส่ง ${S.pending}`;
  if (S.failed) t += ` · ส่งไม่สำเร็จ ${S.failed}`;
  if (S.syncing) t += ' · กำลังส่ง…';
  c.textContent = t;
}

async function syncNow() {
  if (!S.token) return;
  if (S.syncing) { S.syncAgain = true; return; }   // มีงานใหม่เข้าคิวระหว่างส่ง ส่งต่อทันทีเมื่อรอบนี้จบ
  S.syncing = true; S.syncAgain = false; renderSync();
  try {
    await ensureNumbers(S.branch);
    const q = (await DB.all('queue')).filter((o) => o.state === 'pending').sort((a, b) => a.seq - b.seq);
    for (let i = 0; i < q.length; i += 20) {
      const batch = q.slice(i, i + 20);
      const send = batch.map((o) => { const c = Object.assign({}, o); delete c.state; delete c.seq; delete c.error; return c; });
      const results = await api('sync', { ops: send });
      for (const r of results) {
        const op = batch.find((o) => o.op_id === r.op_id);
        if (!op) continue;
        if (r.ok) {
          await DB.del('queue', op.op_id);
          await markSynced(op, true);
        } else {
          op.state = 'error'; op.error = r.error;
          await DB.put('queue', op);
          await markSynced(op, false, r.error);
        }
      }
    }
  } catch (e) {
    if (!(e instanceof NetError)) toast('ส่งข้อมูลไม่สำเร็จ: ' + e.message);
  } finally {
    S.syncing = false;
    if (S.syncAgain && S.online) setTimeout(syncNow, 300);
    await refreshPending();
  }
}

async function markSynced(op, ok, error) {
  const no = op.bill ? op.bill.bill_no : op.bill_no;
  if (no) {
    const b = await DB.get('bills', no);
    if (b) { b._sync = ok ? 'synced' : 'error'; b._error = error || ''; await DB.put('bills', b); }
  }
  if (op.type === 'saveEstimate') {
    const e = await DB.get('estimates', op.estimate.est_no);
    if (e) { e._sync = ok ? 'synced' : 'error'; e._error = error || ''; await DB.put('estimates', e); }
  }
  if (op.type === 'saveCustomer') {
    const c = await DB.get('customers', op.customer.id);
    if (c && ok) { delete c._pending; await DB.put('customers', c); await loadCustomers(); }
  }
}

function showQueue() {
  DB.all('queue').then((q) => {
    q.sort((a, b) => a.seq - b.seq);
    const label = { createBill: 'บิลขาย', editBill: 'แก้ไขบิล', voidBill: 'ยกเลิกบิล', returnBill: 'บิลคืน', saveCustomer: 'ลูกค้า', addPart: 'อะไหล่ด่วน', saveEstimate: 'ใบประเมิน' };
    const box = openModal(`<h2>ข้อมูลที่ยังไม่ได้ส่งขึ้นระบบ</h2>
      <p class="muted">${S.online ? 'ออนไลน์' : 'ออฟไลน์อยู่ ระบบจะส่งให้อัตโนมัติเมื่อมีอินเทอร์เน็ต'}</p>
      ${q.length ? `<table class="grid small"><thead><tr><th>รายการ</th><th>บิล/ลูกค้า</th><th>สถานะ</th><th></th></tr></thead><tbody>
      ${q.map((o) => `<tr><td>${label[o.type] || o.type}</td><td>${esc(o.bill ? o.bill.bill_no : o.bill_no || (o.customer && o.customer.display) || (o.part && o.part.code) || (o.estimate && o.estimate.est_no))}</td>
        <td>${o.state === 'error' ? `<span class="badge error">ไม่สำเร็จ</span> ${esc(o.error)}` : '<span class="badge pending">รอส่ง</span>'}</td>
        <td>${o.state === 'error' ? `<button data-retry="${esc(o.op_id)}">ลองใหม่</button>` : ''}
            ${o.state === 'error' && isAdmin() ? `<button class="danger" data-drop="${esc(o.op_id)}">ทิ้ง</button>` : ''}</td></tr>`).join('')}
      </tbody></table>` : '<p class="ok">ส่งครบแล้ว</p>'}
      <p class="small muted">รายการที่ไม่สำเร็จเพราะรหัสอนุมัติผิด ให้แอดมินกด "ทิ้ง" แล้วทำรายการใหม่</p>
      <div class="actions"><button data-close>ปิด</button><button class="primary" id="q-sync">ส่งตอนนี้</button></div>`);
    $('#q-sync', box).onclick = async () => { await syncNow(); showQueue(); };
    $$('[data-retry]', box).forEach((b) => b.onclick = async () => {
      const op = await DB.get('queue', b.dataset.retry); op.state = 'pending'; await DB.put('queue', op); await syncNow(); showQueue();
    });
    $$('[data-drop]', box).forEach((b) => b.onclick = async () => {
      if (!confirm('ทิ้งรายการนี้? ข้อมูลนี้จะไม่ถูกส่งขึ้นระบบ')) return;
      await DB.del('queue', b.dataset.drop); await refreshPending(); showQueue();
    });
  });
}

/* ================================================================ อนุมัติโดยแอดมิน */

// คืนค่า {approver, code, reason} หรือ null ถ้ากดยกเลิก
function askApproval(title, detail, needReason = true) {
  return new Promise((resolve) => {
    if (!S.approvers.length) { alert('ยังไม่มีแอดมินที่ตั้งรหัสอนุมัติ'); return resolve(null); }
    const box = openModal(`<h2>${esc(title)}</h2>
      <p>${detail}</p>
      <form class="stack" id="ap-form" autocomplete="off">
        <label>แอดมินผู้อนุมัติ<select name="approver">${S.approvers.map((a) => `<option value="${esc(a.username)}">${esc(a.name)}</option>`).join('')}</select></label>
        <label>รหัสอนุมัติ<input name="code" type="password" required autofocus autocomplete="new-password"></label>
        ${needReason ? '<label>เหตุผล<input name="reason" required></label>' : ''}
        <p class="err" id="ap-err"></p>
        <div class="actions"><button type="button" id="ap-cancel">ยกเลิก</button><button class="primary" type="submit">อนุมัติ</button></div>
      </form>`);
    $('#ap-cancel', box).onclick = () => { closeModal(); resolve(null); };
    $('#ap-form', box).onsubmit = async (ev) => {
      ev.preventDefault();
      const f = Object.fromEntries(new FormData(ev.target));
      const a = S.approvers.find((x) => x.username === f.approver);
      if (!a || (await hashPin(a.salt, f.code)) !== a.hash) { $('#ap-err', box).textContent = 'รหัสอนุมัติไม่ถูกต้อง'; return; }
      closeModal();
      resolve({ approval: { approver: f.approver, code: f.code }, reason: f.reason || '' });
    };
  });
}

/* ================================================================ หน้าขาย */

async function newSale() {
  S.cart = { lines: [] };
  $('#sale-date').value = todayISO();
  $('#sale-date').disabled = !isAdmin();
  $('#sale-customer').value = (customerByName('เงินสด') && customerLabel(customerByName('เงินสด'))) || 'เงินสด';
  $('#sale-discount').value = 0;
  $('#sale-discount-pct').value = 0;
  $('#sale-deposit').checked = false; $('#sale-pay [value="เงินสด"]').checked = true; $('#sale-payref').value = ''; $('#sale-payref').disabled = true;
  $('#scan-msg').textContent = '';
  renderCart();
  await ensureNumbers(S.branch);
  $('#sale-no').textContent = (await peekNumber(S.branch)) || 'รอจองเลข (ต่อเน็ต)';
  await showDevice();
  $('#scan-input').focus();
}

function addToCart(p, qty = 1) {
  const ex = S.cart.lines.find((l) => codeKey(l.code) === p.k);
  if (ex) ex.qty = num(ex.qty) + qty;
  else S.cart.lines.push({ code: p.code, name: p.name, model: p.model, qty, unit: num(p.price), listPrice: num(p.price) });
  renderCart();
}

function partBadge(code) {
  const p = findPart(code);
  if (!p) return '';
  if (p.status === PART_PENDING) return ' <span class="badge pending" title="พนักงานเพิ่มด่วน ลดราคาไม่ได้จนกว่าแอดมินตรวจ">รอตรวจ</span>';
  if (p.maxDisc !== '' && p.maxDisc != null && num(p.maxDisc) < branchPct()) return ` <span class="badge" title="อะไหล่ตัวนี้ลดได้น้อยกว่าปกติ">${num(p.maxDisc) ? 'ลดได้ ' + num(p.maxDisc) + '%' : 'ลดไม่ได้'}</span>`;
  return '';
}

function renderCart() {
  const tb = $('#sale-table tbody');
  tb.innerHTML = S.cart.lines.map((l, i) => `<tr>
    <td>${i + 1}</td><td>${esc(l.code)}</td><td>${esc(l.name)}${partBadge(l.code)}</td><td>${esc(l.model)}</td>
    <td class="num"><input type="number" min="1" step="1" data-i="${i}" data-f="qty" value="${l.qty}"></td>
    <td class="num"><input type="number" min="0" step="0.01" data-i="${i}" data-f="unit" value="${l.unit}" class="${l.unit !== l.listPrice ? 'changed' : ''}" title="ราคาตามฐานข้อมูล ${money(l.listPrice)}"></td>
    <td class="num">${money(l.qty * l.unit)}</td>
    <td><button class="ghost" data-del="${i}" title="ลบรายการนี้">✕</button></td></tr>`).join('') ||
    '<tr><td colspan="8" class="muted">ยังไม่มีรายการ ยิงบาร์โค้ดได้เลย</td></tr>';
  const total = round2(S.cart.lines.reduce((s, l) => s + num(l.qty) * num(l.unit), 0));
  // ช่อง % กับช่องบาทผูกกัน: แก้ช่องไหน อีกช่องคำนวณตาม
  if (S.cart.discBy === 'pct') $('#sale-discount').value = round2(total * num($('#sale-discount-pct').value) / 100);
  const disc = num($('#sale-discount').value);
  if (S.cart.discBy !== 'pct') $('#sale-discount-pct').value = total ? round2(disc / total * 100) : 0;
  const allowed = allowedDiscount(S.cart.lines), pct = branchPct();
  const odd = S.cart.lines.filter((l) => partCap(l.code, pct) < pct).length;
  $('#sale-disc-hint').innerHTML = `ลดไปแล้ว <b>${total ? round2(disc / total * 100) : 0}%</b> · สาขานี้ลดได้ ${pct}%`
    + (total ? ` · บิลนี้ลดได้สูงสุด <b>${money(allowed)}</b> บาท` : '')
    + (odd ? ` (มีอะไหล่ ${odd} รายการที่ลดได้น้อยกว่าปกติ)` : '')
    + (disc > allowed + 0.001 ? ' <span class="err">เกิน ต้องให้แอดมินอนุมัติ</span>' : '');
  $('#sale-total').textContent = money(total);
  $('#sale-net').textContent = money(total - disc);
  $('#sale-bahttext').textContent = S.cart.lines.length ? '(' + bahtText(total - disc) + ')' : '';
}

function scan(code) {
  code = code.trim();
  if (!code) return;
  let qty = 1;
  const m = code.match(/^(\d+)\*(.+)$/);   // พิมพ์ 3*รหัส เพื่อใส่จำนวน 3
  if (m) { qty = num(m[1]); code = m[2]; }
  const p = findPart(code);
  $('#scan-msg').textContent = '';
  if (p) { addToCart(p, qty); return; }
  if (!S.partsList.length) { $('#scan-msg').textContent = 'ยังไม่มีข้อมูลอะไหล่ในเครื่อง ต่ออินเทอร์เน็ตเพื่อโหลดก่อน'; return; }
  $("#scan-msg").textContent = `ไม่พบรหัส "${code}" ค้นหาจากชื่อ หรือกด เพิ่มอะไหล่ด่วน`;
  openPartSearch(code, (pp) => addToCart(pp, qty));
}

function openPartSearch(q, onPick, opt = {}) {
  const box = openModal(`<h2>ค้นหาอะไหล่</h2>
    <input id="ps-q" placeholder="พิมพ์ชื่อ รุ่น หรือรหัสบางส่วน" style="width:100%" value="${esc(q || '')}" autofocus>
    <div class="search-results"><table class="grid small"><thead><tr><th>รหัส</th><th>ชื่ออะไหล่</th><th>รุ่น</th><th class="num">ราคา</th></tr></thead><tbody id="ps-body"></tbody></table></div>
    <div class="actions">${opt.quickAdd === false ? '' : '<button type="button" id="ps-add" title="รหัสไม่มีในระบบ ลูกค้ารอ">➕ เพิ่มอะไหล่ด่วน</button>'}<button data-close>ปิด</button></div>`);
  let found = [];
  if ($('#ps-add', box)) $('#ps-add', box).onclick = () => { const v = $('#ps-q', box).value.trim(); closeModal(); quickAddPart(/\s/.test(v) ? '' : v, v, onPick); };
  const run = () => {
    found = searchParts($('#ps-q', box).value);
    $('#ps-body', box).innerHTML = found.map((p, i) => `<tr class="clickable" data-pick="${i}"><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td>${esc(p.model)}</td><td class="num">${money(p.price)}</td></tr>`).join('')
      || '<tr><td colspan="4" class="muted">ไม่พบ</td></tr>';
  };
  $('#ps-q', box).oninput = run;
  $('#ps-q', box).onkeydown = (e) => { if (e.key === 'Enter' && found[0]) { closeModal(); onPick(found[0]); $('#scan-input').focus(); } };
  $('#ps-body', box).onclick = (e) => {
    const tr = e.target.closest('[data-pick]');
    if (tr) { closeModal(); onPick(found[num(tr.dataset.pick)]); $('#scan-input').focus(); }
  };
  run();
}

// พนักงานเพิ่มอะไหล่เองกรณีฉุกเฉิน: ใช้ได้ทันที (แม้ออฟไลน์) ติดสถานะรอแอดมินตรวจ และลดราคาไม่ได้
function quickAddPart(code, name, onAdded) {
  const box = openModal(`<h2>เพิ่มอะไหล่ด่วน</h2>
    <p class="small muted">ใช้เมื่อรหัสยังไม่มีในระบบและต้องรีบออกบิล อะไหล่จะขึ้นว่า "รอตรวจ" จนกว่าแอดมินยืนยัน และระหว่างนั้นลดราคาไม่ได้</p>
    <form class="form-grid" id="qa-form">
      <label>รหัสอะไหล่ *<input name="code" required value="${esc(code)}"></label>
      <label>ชื่ออะไหล่ *<input name="name" required value="${esc(code ? '' : name)}"></label>
      <label>รุ่น<input name="model"></label>
      <label>ราคาขาย (รวมภาษี) *<input name="price" type="number" step="0.01" min="0.01" required></label>
      <div class="actions"><button type="button" data-close>ยกเลิก</button><button class="primary" type="submit">เพิ่มและใส่ในบิล</button></div>
    </form>`);
  $(code ? '[name=name]' : '[name=code]', box).focus();
  $('#qa-form', box).onsubmit = async (ev) => {
    ev.preventDefault();
    const f = Object.fromEntries(new FormData(ev.target));
    f.code = f.code.trim(); f.name = f.name.trim();
    if (findPart(f.code)) { closeModal(); toast('รหัสนี้มีอยู่แล้วในระบบ'); onAdded(findPart(f.code)); return; }
    if (!(num(f.price) > 0)) return toast('ใส่ราคาขาย');
    const p = { k: codeKey(f.code), code: f.code, name: f.name, model: f.model.trim(), price: num(f.price), cost: '', u: '',
      maxDisc: 0, status: PART_PENDING, addedBy: S.user.name };
    await DB.put('parts', p);
    S.parts.set(p.k, p); S.partsList.push(p);
    const n = normCode(p.code); if (n && !S.partsNorm.has(n)) S.partsNorm.set(n, p);
    await enqueue({ type: 'addPart', part: { code: p.code, name: p.name, model: p.model, price: p.price } });
    closeModal();
    toast('เพิ่มอะไหล่ ' + p.code + ' แล้ว (รอแอดมินตรวจ)');
    onAdded(p);
  };
}

async function saveSale(print) {
  const lines = S.cart.lines.filter((l) => num(l.qty) > 0);
  if (!lines.length) return toast('ยังไม่มีรายการอะไหล่');
  const custText = $('#sale-customer').value.trim() || 'เงินสด';
  const cust = customerByName(custText);
  const custName = customerNameFor(custText);
  const total = round2(lines.reduce((s, l) => s + num(l.qty) * num(l.unit), 0));
  const discount = round2(num($('#sale-discount').value));
  if (discount > total) return toast('ส่วนลดมากกว่ายอดรวม');
  const allowed = allowedDiscount(lines);
  const priceChanged = lines.some((l) => Math.abs(l.unit - l.listPrice) > 0.001);
  const overDisc = discount > allowed + 0.001;
  const pay = payPicker($('#sale-pay'));
  const credit = pay.pay_method === PAY_CREDIT;
  if (credit && (!cust || custName === 'เงินสด' || /^มัดจำ/.test(custName))) return toast('ขายค้างจ่ายต้องเลือกชื่อลูกค้าที่มีในระบบ (ไม่ใช่ "เงินสด")', 5000);
  const needAp = priceChanged || overDisc || credit;
  let ap = null;
  if (needAp) {
    const why = [credit ? `ขายค้างจ่าย ${money(total - discount)} บาท ให้ ${custName}` : '', priceChanged ? 'มีการแก้ราคาขาย' : '', overDisc ? `ส่วนลด ${money(discount)} บาท เกินที่ลดได้ ${money(allowed)} บาท` : ''].filter(Boolean).join(' และ ');
    ap = await askApproval('ต้องให้แอดมินอนุมัติ', esc(why), false);
    if (!ap) return;
  }
  const order = { order_type: $('#sale-deposit').checked ? DEPOSIT : '', paid: '' };
  if (pay.pay_method === 'โอนเงิน' && !pay.pay_ref && !confirm('ยังไม่ได้ใส่รายละเอียดการโอน จะบันทึกต่อเลยไหม?')) return $('#sale-payref').focus();
  let no;
  try { no = await takeNumber(S.branch); } catch (e) { return toast(e.message, 6000); }
  const bill = {
    bill_no: no, branch: S.branch, date: $('#sale-date').value || todayISO(), time: nowTime(), type: 'ขาย', ref_bill: '',
    customer_id: cust ? cust.id : '', customer_name: custName, total, discount, net: round2(total - discount), status: 'ปกติ',
    created_by: S.user.username, created_at: nowISO(), note: '', order_type: order.order_type, paid: order.paid,
    pay_method: pay.pay_method, pay_ref: pay.pay_ref, approved_by: ap ? ap.approval.approver : '',
    lines: lines.map((l, i) => ({ line: i + 1, code: l.code, name: l.name, model: l.model, qty: num(l.qty), unit: num(l.unit), amount: round2(num(l.qty) * num(l.unit)) }))
  };
  await DB.put('bills', Object.assign({ _sync: 'pending' }, bill));
  await enqueue({ type: 'createBill', bill, approval: ap ? ap.approval : undefined });
  if (S.cart.fromEst) await markEstimateBilled(S.cart.fromEst, no);
  toast('บันทึกบิล ' + no + ' แล้ว');
  S.printBill = bill;
  renderPrintPreview(bill);
  if (print) printBill(bill);
  newSale();
}

/* ================================================================ หมายเหตุใบเบิก (มัดจำจ่ายเงินแล้วรอของ) */

// ทุกบิลรับเงินครบ ไม่มีค้างจ่าย สั่งด่วน/สั่งรายสัปดาห์ เลิกใช้แล้ว แต่บิลเก่ายังแสดงชื่อเดิม
const DEPOSIT = 'มัดจำรอของ';
const ORDER_LABEL = { [DEPOSIT]: 'มัดจำจ่ายเงินแล้วรอของ', 'สั่งด่วน': 'รายการอะไหล่สั่งด่วน', 'สั่งรายสัปดาห์': 'รายการสั่งรายสัปดาห์' };
// วิธีชำระเงิน: เงินสด / โอนเงิน + รายละเอียดที่พนักงานพิมพ์เอง
function payPicker(scope) {
  const method = (scope.querySelector('input[type=radio]:checked') || {}).value || 'เงินสด';
  const ref = scope.querySelector('.payref');
  ref.disabled = method !== 'โอนเงิน';
  return { pay_method: method, pay_ref: method === 'โอนเงิน' ? ref.value.trim() : '' };
}
const PAY_CREDIT = 'ค้างจ่าย';
function payText(b) {
  if (!b || !b.pay_method) return '';
  let t = b.pay_method + (b.pay_method === 'โอนเงิน' && b.pay_ref ? ' (' + b.pay_ref + ')' : '');
  if (b.pay_method === PAY_CREDIT) t += b.settled_date ? ` → รับชำระแล้ว ${thDate(b.settled_date)} ${b.settled_method}${b.settled_ref ? ' (' + b.settled_ref + ')' : ''}` : ' (ยังไม่ได้รับเงิน)';
  return t;
}
function payPickerHtml(name, b) {
  const m = (b && b.pay_method) || 'เงินสด';
  return `<div class="order-opts" id="${name}"><span>ชำระโดย</span>
    ${['เงินสด', 'โอนเงิน', PAY_CREDIT].map((v) => `<label><input type="radio" name="${name}" value="${v}" ${m === v ? 'checked' : ''}> ${v}</label>`).join('')}
    <input class="payref" placeholder="รายละเอียดการโอน" value="${esc((b && b.pay_ref) || '')}" ${m === 'โอนเงิน' ? '' : 'disabled'}></div>`;
}
function orderText(b, long = true) {
  if (!b || !b.order_type) return '';
  return long ? ORDER_LABEL[b.order_type] || b.order_type : b.order_type;
}

/* ================================================================ ใบเสร็จ */

// part = แผ่นที่พิมพ์ (บิลยาวแบ่งแผ่นละ RC_LINES รายการ ครึ่ง A4 ต่อแผ่นเสมอ)
const RC_LINES = 10;
function receiptHtml(bill, copy, part = { start: 0, page: 1, pages: 1 }) {
  const c = S.customers.find((x) => x.id === bill.customer_id);
  const isRet = bill.type === 'คืน';
  // ทุกบิลรับเงินครบแล้ว ตัดคำว่า (ค้างจ่าย) ที่อาจค้างอยู่ในชีต Settings ออก
  const title = isRet ? 'ใบคืนอะไหล่' : String(S.settings.receipt_title || 'ใบเบิกอะไหล่').replace(/\s*\(?ค้างจ่าย\)?/g, '').trim() || 'ใบเบิกอะไหล่';
  const all = bill.lines || [];
  const lines = all.slice(part.start, part.start + RC_LINES);
  const last = part.page === part.pages;
  const br = branchInfo(bill.branch);
  // วิธีชำระ / ประเภทการสั่ง แสดงเป็นป้ายเล็ก ๆ แทนช่องหมายเหตุแบบเดิม
  const tags = [];
  if (!isRet && bill.pay_method === PAY_CREDIT) tags.push(`<span class="rc-chip warn">ค้างจ่าย</span>` + (bill.settled_date ? ` ชำระแล้ว ${thDate(bill.settled_date)} ${esc(bill.settled_method + (bill.settled_ref ? ' ' + bill.settled_ref : ''))}` : ''));
  else if (!isRet && bill.pay_method) tags.push(`<span class="rc-chip">${esc(bill.pay_method)}</span>` + (bill.pay_method === 'โอนเงิน' && bill.pay_ref ? ' ' + esc(bill.pay_ref) : ''));
  if (bill.order_type) tags.push(`<span class="rc-chip warn">${esc(orderText(bill))}</span>`);
  const cus = [customerFull(c), c && c.address].filter(Boolean).join(' · ');
  const rows = lines.map((l, i) => `<tr><td class="c">${part.start + i + 1}</td><td class="code">${esc(l.code)}</td><td>${esc(l.name)}</td><td class="mdl">${esc(l.model)}</td>
    <td class="num">${l.qty}</td><td class="num">${money(l.unit)}</td><td class="num">${money(l.amount)}</td></tr>`).join('');
  const meta = [isRet ? `<span>อ้างอิงบิล</span><b>${esc(bill.ref_bill)}</b>` : '', ...tags].filter(Boolean);
  return `<div class="receipt rc${part.brk ? ' rc-break' : ''}${bill.status === 'ยกเลิก' ? ' rc-void' : ''}">
    <div class="rc-top">
      <img class="rc-logo" src="${LOGO}" alt="">
      <div class="rc-co"><div class="rc-co-name">${esc(S.settings.shop_name || 'หจก.ประสบชัยกลการ')}</div>
        <div class="rc-co-sub">สาขา${esc(br.name)}${br.address ? ' · ' + esc(br.address) : ''} · โทร ${esc(br.phone || S.settings.shop_phone || '')}</div></div>
      <div class="rc-head"><div class="rc-title">${esc(title)}</div><div class="rc-tag">${copy}${part.pages > 1 ? ` · แผ่น ${part.page}/${part.pages}` : ''}</div></div>
    </div>
    <div class="rc-bar">
      <div class="rc-cus"><span>ลูกค้า</span><b>${esc(bill.customer_name)}</b></div>
      <div class="rc-meta"><div><span>เลขที่</span><b>${esc(bill.bill_no)}</b></div><div><span>วันที่</span><b>${thDate(bill.date)}</b></div></div>
    </div>
    <div class="rc-info">
      <div class="rc-sub">${cus && cus !== bill.customer_name ? esc(cus) : ''}</div>
      <div class="rc-right">${meta.join(' &nbsp; ')}${isRet ? '' : '<span class="rc-paydate"><span>จ่ายเงินวันที่</span><i></i></span>'}</div>
    </div>
    <table class="rc-table"><colgroup><col style="width:5%"><col style="width:20%"><col><col style="width:14%"><col style="width:6%"><col style="width:12%"><col style="width:13%"></colgroup>
      <thead><tr><th class="c">#</th><th>รหัสอะไหล่</th><th>ชื่ออะไหล่</th><th>รุ่น</th><th class="num">จำนวน</th><th class="num">หน่วยละ</th><th class="num">จำนวนเงิน</th></tr></thead>
      <tbody>${rows}</tbody></table>
    ${!last ? `<div class="rc-sum"><div class="rc-words">มีรายการต่อแผ่นถัดไป (แผ่น ${part.page + 1}/${part.pages})</div></div>` : `<div class="rc-sum">
      <div class="rc-words">${num(bill.net) ? '(' + esc(bahtText(bill.net)) + ')' : ''}</div>
      <div class="rc-totals"><span>รวม</span><b>${all.reduce((a, l) => a + num(l.qty), 0).toLocaleString()}</b> ชิ้น${num(bill.discount) ? `<span>ก่อนลด</span><b>${money(bill.total)}</b><span>ส่วนลด</span><b>${money(bill.discount)}</b>` : ''}</div>
      <div class="rc-net"><span>${isRet ? 'ยอดคืน' : 'เป็นเงิน'}</span><b>${money(bill.net)}</b></div>
    </div>
    <div class="rc-sign">
      <div><div class="rc-line"></div>ผู้รับสินค้า${bill.customer_name ? ' (' + esc(bill.customer_name) + ')' : ''}</div>
      <div><div class="rc-line"></div>ผู้จัดอะไหล่${bill.created_by ? ' (' + esc(bill.created_by) + ')' : ''}</div>
      <div><div class="rc-line"></div>ผู้ตรวจสอบ</div>
    </div>`}
    ${last && copy === 'สำหรับลูกค้า' ? `<div class="rc-note">${esc(S.settings.receipt_note || '')}<br><b>${esc(S.settings.receipt_footer || '')}</b></div>` : ''}
    ${bill.status === 'ยกเลิก' ? '<div class="po-stamp">ยกเลิก</div>' : ''}
  </div>`;
}

function billDocHtml(bill) {
  // แผ่น A4 ละ 1 ชุด: ครึ่งบนร้านเก็บ ครึ่งล่างลูกค้า รายการเกิน 10 ขึ้น A4 แผ่นใหม่
  const pages = Math.max(1, Math.ceil((bill.lines || []).length / RC_LINES));
  let html = '';
  for (let p = 1; p <= pages; p++) {
    const part = { start: (p - 1) * RC_LINES, page: p, pages };
    html += receiptHtml(bill, 'สำหรับร้านเก็บ', part) + receiptHtml(bill, 'สำหรับลูกค้า', Object.assign({ brk: p < pages }, part));
  }
  return html;
}
function renderPrintPreview(bill) { $('#print-preview').innerHTML = bill ? billDocHtml(bill) : ''; if (bill) $('#print-no').value = bill.bill_no; }
// ขอบกระดาษ 0 เฉพาะตอนพิมพ์ใบนี้ Chrome จะไม่พิมพ์วันที่ ชื่อเว็บ เลขหน้า (ตัวเอกสารเว้นขอบเอง)
function setPrintPage(css) {
  let el = $('#page-style');
  if (!el) { el = document.createElement('style'); el.id = 'page-style'; document.head.appendChild(el); }
  el.textContent = css || '';
}
window.addEventListener('afterprint', () => setPrintPage(''));
function printBill(bill) {
  $('#print-area').innerHTML = billDocHtml(bill);
  $('#print-area').className = 'print-only pa-tight';
  setPrintPage('@page { size: A4 portrait; margin: 0; }');
  setTimeout(() => window.print(), 50);
}

async function loadBill(no) {
  no = String(no || '').trim().toUpperCase();
  const local = await DB.get('bills', no);
  if (S.online || !local) {
    try {
      const b = await api('bill', { bill_no: no });
      const keep = local && local._sync !== 'synced' ? local : null;
      if (keep) return keep;     // ยังมีการเปลี่ยนแปลงในเครื่องที่ยังไม่ได้ส่ง ใช้ของในเครื่อง
      b._sync = 'synced';
      await DB.put('bills', b);
      return b;
    } catch (e) { if (!local) throw e; }
  }
  return local;
}

/* ================================================================ รายละเอียดบิล แก้ไข ยกเลิก คืน */

async function openBill(no) {
  let b;
  try { b = await loadBill(no); } catch (e) { return toast('เปิดบิล ' + no + ' ไม่สำเร็จ: ' + e.message, 8000); }
  if (!isAdmin() && b.branch !== S.user.branch) return toast('ไม่มีสิทธิ์ดูบิลสาขาอื่น');
  const syncBadge = b._sync === 'pending' ? '<span class="badge pending">รอส่ง</span>' : b._sync === 'error' ? `<span class="badge error">ส่งไม่สำเร็จ: ${esc(b._error)}</span>` : '';
  const canChange = b.status === 'ปกติ';
  const box = openModal(`<h2>บิล ${esc(b.bill_no)} ${b.type === 'คืน' ? '<span class="badge ret">บิลคืน</span>' : ''}
      ${b.status === 'ยกเลิก' ? '<span class="badge void">ยกเลิก</span>' : ''} ${syncBadge}</h2>
    <p>สาขา ${esc(branchInfo(b.branch).name)} · วันที่ ${thDate(b.date)} ${esc(b.time || '')} · ลูกค้า <b>${esc(b.customer_name)}</b> · ผู้ขาย ${esc(b.created_by)}
      ${b.ref_bill ? ' · อ้างอิง ' + esc(b.ref_bill) : ''}${b.approved_by ? ' · อนุมัติโดย ' + esc(b.approved_by) : ''}${b.pay_method ? ' · ชำระโดย ' + esc(payText(b)) : ''}</p>
    ${b.note ? `<p class="muted">${esc(b.note)}</p>` : ''}
    <table class="grid small"><thead><tr><th>รหัส</th><th>ชื่อ</th><th>รุ่น</th><th class="num">จำนวน</th><th class="num">หน่วยละ</th><th class="num">จำนวนเงิน</th></tr></thead>
    <tbody>${(b.lines || []).map((l) => `<tr class="${l.status === 'ยกเลิก' ? 'void' : ''}"><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="num">${l.qty}</td><td class="num">${money(l.unit)}</td><td class="num">${money(l.amount)}</td></tr>`).join('')}</tbody></table>
    <p class="num">รวม ${money(b.total)} · ส่วนลด ${money(b.discount)} · <b>สุทธิ ${money(b.net)}</b></p>
    ${b.order_type ? `<p><span class="badge pending">${esc(orderText(b))}</span></p>` : ''}
    ${b.returns && b.returns.length ? `<p class="muted">มีบิลคืน: ${b.returns.map(esc).join(', ')}</p>` : ''}
    <div class="actions">
      <button data-close>ปิด</button>
      <button id="bd-print">พิมพ์</button>
      ${canChange && b.pay_method === PAY_CREDIT && !b.settled_date ? '<button class="primary" id="bd-settle">รับชำระบิลค้าง</button>' : ''}
      ${canChange ? '<button id="bd-edit">แก้ไขบิล</button>' : ''}
      ${canChange && b.type === 'ขาย' ? '<button id="bd-return">ทำบิลคืนอะไหล่</button>' : ''}
      ${canChange ? '<button class="danger" id="bd-void">ยกเลิกบิล</button>' : ''}
    </div>`);
  $('#bd-print', box).onclick = () => { closeModal(); printBill(b); };
  if ($('#bd-edit', box)) $('#bd-edit', box).onclick = () => editBill(b);
  if ($('#bd-settle', box)) $('#bd-settle', box).onclick = () => settleBill(b);
  if ($('#bd-void', box)) $('#bd-void', box).onclick = () => voidBill(b);
  if ($('#bd-return', box)) $('#bd-return', box).onclick = () => returnBill(b);
}

// รับชำระบิลค้างจ่าย: ลงวันที่รับเงินจริง (ใบคุมวันนี้) บิลเดิมยังเป็นค้างจ่ายในใบคุมวันที่ขาย
function settleBill(b) {
  const box = openModal(`<h2>รับชำระบิลค้าง ${esc(b.bill_no)}</h2>
    <p>ลูกค้า <b>${esc(b.customer_name)}</b> · บิลวันที่ ${thDate(b.date)} · ยอด <b>${money(b.net)}</b> บาท</p>
    <div class="order-opts" id="st-pay"><span>รับเป็น</span>
      <label><input type="radio" name="st-pay" value="เงินสด" checked> เงินสด</label>
      <label><input type="radio" name="st-pay" value="โอนเงิน"> โอนเงิน</label>
      <input class="payref" placeholder="รายละเอียดการโอน" disabled></div>
    <div class="actions"><button data-close>ยกเลิก</button><button class="primary" id="st-save">บันทึกรับเงิน ${money(b.net)} บาท</button></div>`);
  $('#st-pay', box).onchange = () => payPicker($('#st-pay', box));
  $('#st-save', box).onclick = async () => {
    const p = payPicker($('#st-pay', box));
    if (p.pay_method === 'โอนเงิน' && !p.pay_ref && !confirm('ยังไม่ได้ใส่รายละเอียดการโอน จะบันทึกต่อเลยไหม?')) return;
    const date = todayISO();
    await enqueue({ type: 'settleBill', bill_no: b.bill_no, date, method: p.pay_method, ref: p.pay_ref, by: S.user.username });
    Object.assign(b, { settled_date: date, settled_method: p.pay_method, settled_ref: p.pay_ref, settled_by: S.user.username, _sync: 'pending' });
    await DB.put('bills', b);
    closeModal();
    toast('รับชำระบิล ' + b.bill_no + ' แล้ว');
    if (!$('#tab-control').classList.contains('hidden')) loadControl();
  };
}

function editBill(b) {
  const lines = (b.lines || []).filter((l) => l.status !== 'ยกเลิก').map((l) => {
    const p = findPart(l.code);
    return { code: l.code, name: l.name, model: l.model, qty: Math.abs(l.qty), unit: l.unit, listPrice: p ? num(p.price) : l.unit };
  });
  const box = openModal(`<h2>แก้ไขบิล ${esc(b.bill_no)}</h2>
    <div class="row"><label>ลูกค้า<input id="ed-cus" list="customer-list" value="${esc((customerByName(b.customer_name) && customerLabel(customerByName(b.customer_name))) || b.customer_name)}"></label>
    <input id="ed-scan" placeholder="ยิงบาร์โค้ดเพื่อเพิ่มรายการ"></div>
    <table class="grid small" style="margin-top:8px"><thead><tr><th>รหัส</th><th>ชื่อ</th><th class="num">จำนวน</th><th class="num">หน่วยละ</th><th class="num">จำนวนเงิน</th><th></th></tr></thead><tbody id="ed-body"></tbody></table>
    <div class="row" style="justify-content:flex-end;margin-top:8px"><label>ส่วนลด<input id="ed-disc" type="number" step="0.01" value="${b.discount}"></label><b id="ed-net"></b></div>
    ${b.type === 'คืน' ? '' : `<div class="order-opts" id="ed-order"><span>หมายเหตุ</span>
      <label><input type="checkbox" id="ed-deposit" ${b.order_type === DEPOSIT ? 'checked' : ''}> ${ORDER_LABEL[DEPOSIT]}</label></div>${payPickerHtml('ed-pay', b)}`}
    <div class="actions"><button data-close>ยกเลิก</button><button class="primary" id="ed-save">บันทึกการแก้ไข (ต้องมีรหัสแอดมิน)</button></div>`);
  const render = () => {
    $('#ed-body', box).innerHTML = lines.map((l, i) => `<tr><td>${esc(l.code)}</td><td>${esc(l.name)}</td>
      <td class="num"><input type="number" min="1" data-i="${i}" data-f="qty" value="${l.qty}"></td>
      <td class="num"><input type="number" step="0.01" data-i="${i}" data-f="unit" value="${l.unit}"></td>
      <td class="num">${money(l.qty * l.unit)}</td><td><button class="ghost" data-del="${i}">✕</button></td></tr>`).join('');
    const tot = lines.reduce((s, l) => s + l.qty * l.unit, 0);
    $('#ed-net', box).textContent = 'สุทธิ ' + money(tot - num($('#ed-disc', box).value));
  };
  $('#ed-body', box).onchange = (e) => { const i = e.target.dataset.i; if (i != null) { lines[i][e.target.dataset.f] = num(e.target.value); setTimeout(render); } };
  $('#ed-body', box).onclick = (e) => { const d = e.target.dataset.del; if (d != null) { lines.splice(num(d), 1); render(); } };
  $('#ed-disc', box).oninput = render;
  if ($('#ed-pay', box)) $('#ed-pay', box).onchange = () => payPicker($('#ed-pay', box));
  $('#ed-scan', box).onkeydown = (e) => {
    if (e.key !== 'Enter') return;
    const p = findPart(e.target.value);
    if (!p) return toast('ไม่พบรหัส ' + e.target.value);
    const ex = lines.find((l) => codeKey(l.code) === p.k);
    if (ex) ex.qty++; else lines.push({ code: p.code, name: p.name, model: p.model, qty: 1, unit: num(p.price), listPrice: num(p.price) });
    e.target.value = ''; render();
  };
  render();
  $('#ed-save', box).onclick = async () => {
    if (!lines.length) return toast('บิลต้องมีอย่างน้อย 1 รายการ ถ้าจะลบทั้งบิลให้ใช้ "ยกเลิกบิล"');
    const custText = $('#ed-cus', box).value.trim();
    const custName = customerNameFor(custText);
    const discount = round2(num($('#ed-disc', box).value));
    const dep = $('#ed-deposit', box);
    // ติ๊กมัดจำ = มัดจำรอของ, ไม่ติ๊ก = ปกติ (บิลเก่าที่เป็นสั่งด่วน/สั่งรายสัปดาห์ ถ้าไม่แตะ ให้คงไว้ตามเดิม)
    const order = !dep || (!dep.checked && b.order_type && b.order_type !== DEPOSIT) ? { order_type: b.order_type || '', paid: b.paid }
      : { order_type: dep.checked ? DEPOSIT : '', paid: '' };
    const pay = $('#ed-pay', box) ? payPicker($('#ed-pay', box)) : { pay_method: b.pay_method || '', pay_ref: b.pay_ref || '' };
    const ap = await askApproval('อนุมัติการแก้ไขบิล ' + b.bill_no, 'การแก้ไขจะถูกบันทึกประวัติไว้ทั้งก่อนและหลังแก้');
    if (!ap) return;
    const cust = customerByName(custText);
    const sign = b.type === 'คืน' ? -1 : 1;
    const newLines = lines.map((l, i) => ({ line: i + 1, code: l.code, name: l.name, model: l.model, qty: sign * l.qty, unit: l.unit, amount: round2(sign * l.qty * l.unit), status: 'ปกติ' }));
    const total = round2(newLines.reduce((s, l) => s + l.amount, 0));
    await enqueue({ type: 'editBill', bill_no: b.bill_no, lines: lines.map((l) => ({ code: l.code, name: l.name, model: l.model, qty: l.qty, unit: l.unit })),
      discount, customer_name: custName, customer_id: cust ? cust.id : '', order_type: order.order_type, paid: order.paid, pay_method: pay.pay_method, pay_ref: pay.pay_ref, approval: ap.approval, reason: ap.reason });
    Object.assign(b, { lines: newLines, total, discount, net: round2(total - discount), customer_name: custName, customer_id: cust ? cust.id : '',
      order_type: order.order_type, paid: order.paid, pay_method: pay.pay_method, pay_ref: pay.pay_ref,
      approved_by: ap.approval.approver, note: 'แก้ไข: ' + ap.reason, _sync: 'pending' });
    await DB.put('bills', b);
    toast('บันทึกการแก้ไขแล้ว');
    afterBillChange(b);
  };
}

async function voidBill(b) {
  const ap = await askApproval('ยกเลิกบิล ' + b.bill_no, `ยอด ${money(b.net)} บาท บิลจะไม่ถูกลบ แต่จะขึ้นสถานะยกเลิกและไม่นับในรายงาน`);
  if (!ap) return;
  await enqueue({ type: 'voidBill', bill_no: b.bill_no, approval: ap.approval, reason: ap.reason });
  b.status = 'ยกเลิก'; b.note = 'ยกเลิก: ' + ap.reason; b._sync = 'pending';
  (b.lines || []).forEach((l) => { l.status = 'ยกเลิก'; });
  await DB.put('bills', b);
  toast('ยกเลิกบิล ' + b.bill_no + ' แล้ว');
  afterBillChange(b);
}

function returnBill(b) {
  const lines = (b.lines || []).filter((l) => l.status !== 'ยกเลิก');
  const box = openModal(`<h2>ทำบิลคืนอะไหล่ จากบิล ${esc(b.bill_no)}</h2>
    <p class="muted">ใส่จำนวนที่ลูกค้านำมาคืน</p>
    <table class="grid small"><thead><tr><th>รหัส</th><th>ชื่อ</th><th class="num">ซื้อไป</th><th class="num">หน่วยละ</th><th class="num">คืน</th></tr></thead>
    <tbody>${lines.map((l, i) => `<tr><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td class="num">${l.qty}</td><td class="num">${money(l.unit)}</td>
      <td class="num"><input type="number" min="0" max="${l.qty}" value="0" data-r="${i}"></td></tr>`).join('')}</tbody></table>
    <div class="actions"><button data-close>ยกเลิก</button><button class="primary" id="rt-save">บันทึกบิลคืน (ต้องมีรหัสแอดมิน)</button></div>`);
  $('#rt-save', box).onclick = async () => {
    const ret = [];
    for (const inp of $$('[data-r]', box)) {
      const q = num(inp.value), l = lines[num(inp.dataset.r)];
      if (q < 0 || q > l.qty) return toast('จำนวนคืนเกินที่ซื้อ: ' + l.code);
      if (q > 0) ret.push({ code: l.code, name: l.name, model: l.model, qty: q, unit: l.unit });
    }
    if (!ret.length) return toast('ยังไม่ได้ใส่จำนวนคืน');
    const ap = await askApproval('อนุมัติบิลคืนอะไหล่', `คืน ${ret.length} รายการ จากบิล ${esc(b.bill_no)}`);
    if (!ap) return;
    let no;
    try { no = await takeNumber(b.branch); } catch (e) { return toast(e.message, 6000); }
    const rlines = ret.map((l, i) => ({ line: i + 1, code: l.code, name: l.name, model: l.model, qty: -l.qty, unit: l.unit, amount: round2(-l.qty * l.unit), status: 'ปกติ' }));
    const total = round2(rlines.reduce((s, l) => s + l.amount, 0));
    const discount = num(b.total) ? round2(num(b.discount) * total / num(b.total)) : 0;
    const rb = { bill_no: no, branch: b.branch, date: todayISO(), time: nowTime(), type: 'คืน', ref_bill: b.bill_no, customer_id: b.customer_id,
      customer_name: b.customer_name, total, discount, net: round2(total - discount), status: 'ปกติ', created_by: S.user.username, created_at: nowISO(),
      approved_by: ap.approval.approver, note: 'คืนจากบิล ' + b.bill_no + ': ' + ap.reason, lines: rlines };
    await DB.put('bills', Object.assign({ _sync: 'pending' }, rb));
    await enqueue({ type: 'returnBill', bill: { bill_no: no, ref_bill: b.bill_no, date: rb.date, time: rb.time, created_at: rb.created_at, created_by: rb.created_by, lines: ret }, approval: ap.approval, reason: ap.reason });
    toast('บันทึกบิลคืน ' + no + ' แล้ว');
    S.printBill = rb; renderPrintPreview(rb); printBill(rb);
    afterBillChange(rb);
  };
}

function afterBillChange() {
  closeModal();
  if (!$('#tab-report').classList.contains('hidden')) loadReport();
}

/* ================================================================ ใบประเมินราคาซ่อม */
// บันทึกแยกจากบิลขาย ไม่นับเป็นยอดขาย แก้ไขได้โดยไม่ต้องใช้รหัสแอดมิน
// เปิดบิลขายจากใบประเมินได้ โดยเลือกรายการอะไหล่ที่จะเอาไปขาย

const EST_OPEN = 'ประเมิน', EST_BILLED = 'เปิดบิลแล้ว', EST_VOID = 'ยกเลิก', EST_CLOSED = 'ปิดงาน', EST_WHOLESALE = 'ส่งขายส่ง';
const EST_FIELDS = { date: 'est-date', phone: 'est-phone', address: 'est-address', engine_no: 'est-engine', plate: 'est-plate', color: 'est-color', mechanic: 'est-mechanic', note: 'est-note' };

function newEstimate() {
  S.est = { est_no: '', branch: S.branch, status: EST_OPEN, lines: [] };
  for (const [k, id] of Object.entries(EST_FIELDS)) $('#' + id).value = '';
  $('#est-date').value = todayISO();
  $('#est-customer').value = '';
  renderEstimate();
  ensureNumbers(S.branch, 'est');
}

function loadEstimateForm(e) {
  S.est = JSON.parse(JSON.stringify(e));
  for (const [k, id] of Object.entries(EST_FIELDS)) $('#' + id).value = e[k] == null ? '' : e[k];
  const c = customerByName(e.customer_name);
  $('#est-customer').value = c ? customerLabel(c) : e.customer_name || '';
  renderEstimate();
  switchTab('estimate');
}

function estimateFromForm() {
  const e = S.est;
  for (const [k, id] of Object.entries(EST_FIELDS)) e[k] = $('#' + id).value.trim();
  const text = $('#est-customer').value.trim();
  const c = customerByName(text);
  e.customer_id = c ? c.id : '';
  e.customer_name = customerNameFor(text);
  e.lines = e.lines.map((l, i) => Object.assign(l, { line: i + 1, qty: num(l.qty), unit: num(l.unit), amount: round2(num(l.qty) * num(l.unit)) }));
  e.total = round2(e.lines.reduce((s, l) => s + l.amount, 0));
  return e;
}

function renderEstimate() {
  const e = S.est, locked = e.status !== EST_OPEN;
  $('#est-no').textContent = e.est_no || '(ใบใหม่ ยังไม่บันทึก)';
  $('#est-status').innerHTML = e.status === EST_BILLED ? `<span class="badge">เปิดบิลแล้ว ${esc(e.bill_no)}</span>`
    : e.status === EST_VOID ? '<span class="badge void">ยกเลิก</span>'
    : e.status === EST_CLOSED ? '<span class="badge void">ปิดงาน (ไม่ซ่อม / ไม่ซื้อ)</span>'
    : e.status === EST_WHOLESALE ? '<span class="badge">ส่งขายส่งแล้ว</span>'
    : e._sync === 'pending' ? '<span class="badge pending">รอส่ง</span>' : e._sync === 'error' ? `<span class="badge error">ส่งไม่สำเร็จ: ${esc(e._error)}</span>` : '';
  $('#est-table tbody').innerHTML = e.lines.map((l, i) => `<tr>
    <td>${i + 1}</td><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td>
    <td class="num"><input type="number" min="1" step="1" data-i="${i}" data-f="qty" value="${l.qty}" ${locked ? 'disabled' : ''}></td>
    <td class="num"><input type="number" min="0" step="0.01" data-i="${i}" data-f="unit" value="${l.unit}" ${locked ? 'disabled' : ''}></td>
    <td class="num">${money(num(l.qty) * num(l.unit))}</td>
    <td>${locked ? '' : `<button class="ghost" data-del="${i}" title="ลบรายการนี้">✕</button>`}</td></tr>`).join('')
    || '<tr><td colspan="8" class="muted">ยังไม่มีรายการ ยิงบาร์โค้ดหรือค้นหาชื่ออะไหล่ได้เลย</td></tr>';
  $('#est-total').textContent = money(e.lines.reduce((s, l) => s + num(l.qty) * num(l.unit), 0));
  $$('#tab-estimate .est-head input, #est-scan, #est-note, #est-search, #est-free, #est-save, #est-void').forEach((el) => { el.disabled = locked; });
  $('#est-to-bill').disabled = locked || !e.lines.length;
  $('#est-to-ws').disabled = locked || !e.lines.some((l) => l.code);
  $('#est-close').disabled = locked || !e.est_no;
  $('#est-reopen').classList.toggle('hidden', !(e.status === EST_CLOSED || e.status === EST_WHOLESALE));
  $('#est-save-print').textContent = locked ? 'พิมพ์ / PDF' : 'บันทึกและพิมพ์ / PDF';
}

// 1. ไม่ซ่อม / ไม่ซื้อ: ปิดใบประเมิน (เก็บไว้ดูย้อนหลัง ไม่นับเป็นงานค้าง)
async function closeEstimate() {
  if (!S.est.est_no) return toast('บันทึกใบประเมินก่อน');
  const why = prompt('ปิดงานใบ ' + S.est.est_no + ' (ลูกค้าไม่ซ่อม / ไม่ซื้อ)\nเหตุผล (ไม่ใส่ก็ได้)', '');
  if (why == null) return;
  const e = estimateFromForm();
  e.status = EST_CLOSED;
  if (why.trim()) e.note = [e.note, 'ปิดงาน: ' + why.trim()].filter(Boolean).join(' · ');
  await putEstimate(e);
  S.est = e;
  $('#est-note').value = e.note || '';
  renderEstimate();
  toast('ปิดงานใบประเมินแล้ว');
}

// 3. ขายส่ง: ดาวน์โหลด Excel แบบเดียวกับไฟล์ตัวอย่างนำเข้าออเดอร์ของระบบขายส่ง (รหัสอะไหล่ / จำนวน / ชื่อ)
// แล้วไปกด "นำเข้า Excel" ที่หน้ารับออเดอร์ของระบบขายส่ง รายการที่ไม่มีรหัส (เช่น ค่าแรง) จะไม่ถูกส่งไป
async function estimateToWholesale() {
  const e = estimateFromForm();
  const lines = e.lines.filter((l) => l.code);
  if (!lines.length) return toast('ไม่มีรายการที่มีรหัสอะไหล่');
  if (!window.XLSX) return toast('ตัวสร้าง Excel ยังโหลดไม่เสร็จ ลองใหม่อีกครั้ง');
  const skipped = e.lines.length - lines.length;
  if (!confirm(`ส่งใบ ${e.est_no || '(ใบใหม่)'} ไปขายส่ง ${lines.length} รายการ${skipped ? ` (รายการไม่มีรหัส ${skipped} รายการจะไม่ถูกส่ง)` : ''}\nระบบจะบันทึกใบนี้เป็น "ส่งขายส่ง" และดาวน์โหลดไฟล์ Excel`)) return;
  e.status = EST_WHOLESALE;
  if (!e.est_no) { try { e.est_no = await takeNumber(S.branch, 'est'); } catch (err) { return toast(err.message, 6000); } e.created_at = new Date().toISOString(); }
  await putEstimate(e);
  S.est = e;
  renderEstimate();
  const rows = [['รหัสอะไหล่', 'จำนวน', 'ชื่ออะไหล่ (ไม่ต้องใส่ก็ได้)'], ...lines.map((l) => [l.code, num(l.qty), l.name])];
  const ws = XLSX.utils.aoa_to_sheet(rows); ws['!cols'] = [{ wch: 20 }, { wch: 8 }, { wch: 40 }];
  const info = XLSX.utils.aoa_to_sheet([['ใบประเมิน', e.est_no], ['วันที่', thDate(e.date)], ['ลูกค้า', e.customer_name], ['เบอร์โทร', e.phone || ''], ['สาขา', branchInfo(e.branch).name], ['ยอดตามใบประเมิน', e.total]]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'รายการสั่ง');
  XLSX.utils.book_append_sheet(wb, info, 'ข้อมูลใบประเมิน');
  XLSX.writeFile(wb, `ส่งขายส่ง_${e.est_no}_${String(e.customer_name || '').replace(/[\\/:*?"<>|]/g, '')}.xlsx`);
  toast('ดาวน์โหลดไฟล์แล้ว นำไปกด "นำเข้า Excel" ที่หน้ารับออเดอร์ของระบบขายส่ง', 7000);
}

async function reopenEstimate() {
  if (!confirm('เปิดใบ ' + S.est.est_no + ' กลับมาเป็นใบประเมินที่ยังไม่ปิดงาน?')) return;
  const e = estimateFromForm();
  e.status = EST_OPEN;
  await putEstimate(e);
  S.est = e;
  renderEstimate();
}

function addEstimateLine(p, qty = 1) {
  if (S.est.status !== EST_OPEN) return;
  const ex = p.code && S.est.lines.find((l) => codeKey(l.code) === codeKey(p.code));
  if (ex) ex.qty = num(ex.qty) + qty;
  else S.est.lines.push({ code: p.code || '', name: p.name, model: p.model || '', qty, unit: num(p.price) });
  renderEstimate();
}

function estimateScan(code) {
  code = code.trim();
  if (!code) return;
  let qty = 1;
  const m = code.match(/^(\d+)\*(.+)$/);
  if (m) { qty = num(m[1]); code = m[2]; }
  const p = findPart(code);
  if (p) return addEstimateLine(p, qty);
  openPartSearch(code, (pp) => addEstimateLine(pp, qty), { quickAdd: false });
}

// รายการที่ไม่มีรหัส เช่น ค่าแรง ใส่ได้เฉพาะในใบประเมิน (ไม่เพิ่มเข้าฐานข้อมูลอะไหล่)
function estimateFreeLine() {
  const box = openModal(`<h2>เพิ่มรายการพิเศษ</h2>
    <form class="form-grid" id="ef-form">
      <label>รหัส (ถ้ามี)<input name="code"></label>
      <label>รายการ *<input name="name" required autofocus placeholder="เช่น ค่าแรงเปลี่ยนชุดสี"></label>
      <label>จำนวน<input name="qty" type="number" min="1" value="1"></label>
      <label>ราคา/หน่วย *<input name="price" type="number" step="0.01" min="0" required></label>
      <div class="actions"><button type="button" data-close>ยกเลิก</button><button class="primary" type="submit">เพิ่ม</button></div>
    </form>`);
  $('#ef-form', box).onsubmit = (ev) => {
    ev.preventDefault();
    const f = Object.fromEntries(new FormData(ev.target));
    closeModal();
    addEstimateLine({ code: f.code.trim(), name: f.name.trim(), model: '', price: num(f.price) }, Math.max(1, num(f.qty)));
  };
}

async function saveEstimate(print) {
  const e = estimateFromForm();
  if (!e.lines.length) return toast('ยังไม่มีรายการอะไหล่');
  if (!e.customer_name) return toast('ใส่ชื่อลูกค้า');
  if (!e.est_no) {
    try { e.est_no = await takeNumber(S.branch, 'est'); } catch (err) { return toast(err.message, 6000); }
    e.branch = S.branch; e.created_by = S.user.username; e.created_at = nowISO();
  }
  await putEstimate(e);
  toast('บันทึกใบประเมิน ' + e.est_no + ' แล้ว');
  renderEstimate();
  if (print) printEstimate(e);
}

async function putEstimate(e) {
  const send = Object.assign({}, e); delete send._sync; delete send._error;
  e._sync = 'pending';
  await DB.put('estimates', JSON.parse(JSON.stringify(e)));
  await enqueue({ type: 'saveEstimate', estimate: send });
}

async function markEstimateBilled(estNo, billNo) {
  const e = (await DB.get('estimates', estNo)) || (S.est && S.est.est_no === estNo ? S.est : null);
  if (!e) return;
  Object.assign(e, { status: EST_BILLED, bill_no: billNo });
  await putEstimate(e);
  if (S.est && S.est.est_no === estNo) Object.assign(S.est, { status: EST_BILLED, bill_no: billNo });
}

async function voidEstimate() {
  if (!S.est.est_no) return newEstimate();
  if (!confirm('ยกเลิกใบประเมิน ' + S.est.est_no + ' ?')) return;
  const e = estimateFromForm();
  e.status = EST_VOID;
  await putEstimate(e);
  renderEstimate();
}

// เลือกรายการที่จะเอาไปเปิดบิลขาย (ติ๊กออกได้ แก้จำนวนได้)
async function estimateToBill() {
  const e = estimateFromForm();
  if (!e.lines.length) return;
  await saveEstimate(false);   // เก็บรายการล่าสุดก่อน แล้วค่อยเลือกไปเปิดบิล
  if (!S.est.est_no || S.est._sync !== 'pending') return;
  const box = openModal(`<h2>เปิดบิลขายจากใบประเมิน ${esc(e.est_no)}</h2>
    <p class="small muted">ติ๊กเลือกรายการที่ลูกค้าตกลงซ่อม ราคาจะใช้ราคาขายปัจจุบันในระบบ</p>
    <table class="grid small"><thead><tr><th><input type="checkbox" id="eb-all" checked></th><th>รหัส</th><th>ชื่ออะไหล่</th><th class="num">จำนวน</th><th class="num">ราคา/หน่วย</th></tr></thead>
    <tbody id="eb-body">${e.lines.map((l, i) => {
      const p = l.code && findPart(l.code);
      return `<tr><td><input type="checkbox" data-i="${i}" checked></td><td>${esc(l.code)}</td><td>${esc(l.name)}${p ? '' : ' <span class="badge">ไม่มีในระบบ</span>'}</td>
        <td class="num"><input type="number" min="1" data-q="${i}" value="${l.qty}" style="width:70px"></td><td class="num">${money(p ? p.price : l.unit)}</td></tr>`;
    }).join('')}</tbody></table>
    <p class="small muted" id="eb-sum"></p>
    <div class="actions"><button data-close>ยกเลิก</button><button class="primary" id="eb-go">ไปหน้าขาย</button></div>`);
  const picked = () => $$('#eb-body [data-i]', box).filter((c) => c.checked).map((c) => num(c.dataset.i));
  const sum = () => { $('#eb-sum', box).textContent = `เลือก ${picked().length} จาก ${e.lines.length} รายการ`; };
  $('#eb-all', box).onchange = (ev) => { $$('#eb-body [data-i]', box).forEach((c) => { c.checked = ev.target.checked; }); sum(); };
  $('#eb-body', box).onchange = sum; sum();
  $('#eb-go', box).onclick = async () => {
    const idx = picked();
    if (!idx.length) return toast('เลือกอย่างน้อย 1 รายการ');
    if (S.cart.lines.length && !confirm('หน้าขายมีรายการค้างอยู่ จะล้างแล้วใส่รายการจากใบประเมินแทน?')) return;
    const qtys = Object.fromEntries(idx.map((i) => [i, Math.max(1, num($(`[data-q="${i}"]`, box).value))]));
    const qty = (i) => qtys[i];
    closeModal();
    switchTab('sale');
    await newSale();
    S.cart.fromEst = e.est_no;
    S.cart.lines = idx.map((i) => {
      const l = e.lines[i], p = l.code && findPart(l.code);
      return p ? { code: p.code, name: p.name, model: p.model, qty: qty(i), unit: num(p.price), listPrice: num(p.price) }
        : { code: l.code || 'ไม่มีรหัส', name: l.name, model: l.model, qty: qty(i), unit: num(l.unit), listPrice: num(l.unit) };
    });
    const c = customerByName(e.customer_name);
    $('#sale-customer').value = c ? customerLabel(c) : e.customer_name;
    $('#scan-msg').textContent = `รายการจากใบประเมิน ${e.est_no} ตรวจแล้วกดบันทึกได้เลย`;
    renderCart();
  };
}

// หัวเอกสาร: สำนักงานใหญ่ หรือ "สาขา 1 (แม่สอด)"
function branchTitle(b) { return b.code === 'HQ' || !b.branch_no ? b.name : `สาขา ${b.branch_no} (${b.name})`; }

function estimateHtml(e) {
  const br = branchInfo(e.branch);
  const rows = Math.max(15, e.lines.length);
  let body = '';
  for (let i = 0; i < rows; i++) {
    const l = e.lines[i];
    body += l ? `<tr><td class="c">${i + 1}</td><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="c">${l.qty}</td><td class="num">${money(l.unit)}</td><td class="num">${money(num(l.qty) * num(l.unit))}</td></tr>`
      : `<tr class="empty"><td class="c">${i + 1}</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>`;
  }
  const total = e.lines.reduce((s, l) => s + num(l.qty) * num(l.unit), 0);
  const info = (label, val, cls = '') => `<div class="ed-f ${cls}"><span>${label}</span><b>${esc(val) || '&nbsp;'}</b></div>`;
  return `<div class="est-doc">
    <header class="ed-head">
      <div class="ed-shop">
        <div class="ed-name">${esc(S.settings.shop_name || 'หจก.ประสบชัยกลการ')}</div>
        <div class="ed-branch">${esc(branchTitle(br))}</div>
        <div>${esc(br.address || '')}</div>
        <div>โทร ${esc(br.phone || '')} · ติดต่อฝ่ายบริการ</div>
      </div>
      <div class="ed-titlebox">
        <div class="ed-title">ใบประเมินราคาค่าซ่อม</div>
        <div class="ed-meta"><span>เลขที่</span><b>${esc(e.est_no)}</b></div>
        <div class="ed-meta"><span>วันที่</span><b>${thDate(e.date)}</b></div>
      </div>
    </header>
    <section class="ed-info">
      ${info('ชื่อ-สกุล', e.customer_name, 'w2')}${info('โทร', e.phone)}
      ${info('ที่อยู่', e.address, 'w3')}
      ${info('หมายเลขเครื่อง', e.engine_no)}${info('ทะเบียน', e.plate)}${info('สี', e.color)}
    </section>
    <table class="ed-table">
      <thead><tr><th style="width:7%">ลำดับ</th><th style="width:19%">รหัสอะไหล่</th><th>ชื่ออะไหล่ / รายการ</th><th style="width:13%">รุ่น</th><th style="width:8%">จำนวน</th><th style="width:12%">ราคา/หน่วย</th><th style="width:13%">รวมเงิน</th></tr></thead>
      <tbody>${body}</tbody>
    </table>
    <div class="ed-sum">
      <div class="ed-note">${e.note ? `<b>หมายเหตุ:</b> ${esc(e.note)}` : ''}</div>
      <div class="ed-total"><span>รวมราคาค่าซ่อมโดยประมาณ</span><b>${money(total)}</b><small>(${esc(bahtText(total))})</small></div>
    </div>
    <p class="ed-disclaimer">หมายเหตุ: เอกสารนี้ใช้ในการประเมินค่าซ่อมเบื้องต้น ราคาอะไหล่อาจจะมีการเปลี่ยนแปลง</p>
    <div class="ed-sign">
      <div><div class="line"></div>ผู้ประเมิน${e.mechanic ? ` (ช่าง${esc(e.mechanic)})` : ''}</div>
      <div><div class="line"></div>ลูกค้า</div>
    </div>
  </div>`;
}

// พิมพ์ หรือเลือกเครื่องพิมพ์ "บันทึกเป็น PDF" เพื่อส่งให้ลูกค้า (ชื่อไฟล์ = เลขที่ใบ + ชื่อลูกค้า)
function printEstimate(e) { printDoc(estimateHtml(e), `ใบประเมินราคา ${e.est_no} ${e.customer_name || ''}`.trim()); }

async function estimateList() {
  const box = openModal(`<h2>ค้นหาใบประเมินราคา</h2>
    <div class="row"><label>ตั้งแต่<input type="date" id="el-from"></label><label>ถึง<input type="date" id="el-to"></label>
      <input id="el-q" placeholder="ชื่อ เบอร์ ทะเบียน หรือเลขที่" style="flex:1"><button class="primary" id="el-go">ค้นหา</button></div>
    <p class="small muted" id="el-msg"></p>
    <div class="search-results"><table class="grid small"><thead><tr><th>เลขที่</th><th>วันที่</th><th>ลูกค้า</th><th>ทะเบียน</th><th class="num">รวม</th><th>สถานะ</th></tr></thead><tbody id="el-body"></tbody></table></div>
    <div class="actions"><button data-close>ปิด</button></div>`);
  const d = new Date(); d.setDate(d.getDate() - 30);
  $('#el-from', box).value = `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`;
  $('#el-to', box).value = todayISO();
  let found = [];
  const run = async () => {
    const from = $('#el-from', box).value, to = $('#el-to', box).value, q = $('#el-q', box).value.trim().toLowerCase();
    const local = (await DB.all('estimates')).filter((e) => e.branch === S.branch);
    let rows = local, src = 'ข้อมูลในเครื่อง';
    try {
      const server = await api('estimates', { branch: S.branch, from, to, q });
      const pending = local.filter((e) => e._sync && e._sync !== 'synced');
      const byNo = new Map(server.map((e) => [e.est_no, Object.assign(e, { _sync: 'synced' })]));
      pending.forEach((e) => byNo.set(e.est_no, e));
      rows = [...byNo.values()];
      await DB.putMany('estimates', server.filter((e) => !pending.some((p) => p.est_no === e.est_no)));
      src = 'ข้อมูลจากระบบกลาง';
    } catch (err) { if (!(err instanceof NetError)) toast(err.message); }
    found = rows.filter((e) => e.date >= from && e.date <= to
      && (!q || [e.est_no, e.customer_name, e.phone, e.plate, e.engine_no].join(' ').toLowerCase().includes(q)))
      .sort((a, b) => a.est_no < b.est_no ? 1 : -1);
    $('#el-msg', box).textContent = `${src} · ${found.length} ใบ`;
    $('#el-body', box).innerHTML = found.map((e, i) => `<tr class="clickable" data-pick="${i}"><td>${esc(e.est_no)}</td><td>${thDate(e.date)}</td><td>${esc(e.customer_name)}</td>
      <td>${esc(e.plate)}</td><td class="num">${money(e.total)}</td><td>${esc(e.status)}${e.bill_no ? ' ' + esc(e.bill_no) : ''}${e._sync === 'pending' ? ' <span class="badge pending">รอส่ง</span>' : ''}</td></tr>`).join('')
      || '<tr><td colspan="6" class="muted">ไม่พบ</td></tr>';
  };
  $('#el-go', box).onclick = run;
  $('#el-q', box).onkeydown = (ev) => { if (ev.key === 'Enter') run(); };
  $('#el-body', box).onclick = (ev) => { const tr = ev.target.closest('[data-pick]'); if (tr) { closeModal(); loadEstimateForm(found[num(tr.dataset.pick)]); } };
  run();
}

/* ================================================================ รายงาน */

// บิลจากระบบกลาง + บิลในเครื่องที่ยังไม่ได้ส่ง (ใช้ทั้งหน้ารายงานและหน้าใบคุม)
async function fetchBills(branch, from, to) {
  let bills = [], source = '';
  try {
    bills = await api('bills', { branch, from, to });
    source = 'ข้อมูลจากระบบกลาง';
  } catch (e) {
    if (!(e instanceof NetError)) throw e;
    source = 'ออฟไลน์: แสดงเฉพาะบิลที่อยู่ในเครื่องนี้';
  }
  const local = (await DB.all('bills')).filter((b) => b.date >= from && b.date <= to && (branch === '*' || b.branch === branch));
  const map = new Map(bills.map((b) => [b.bill_no, b]));
  local.forEach((b) => { if (b._sync !== 'synced' || !map.has(b.bill_no)) map.set(b.bill_no, b); });
  return { source, bills: Array.from(map.values()).sort((a, b) => (a.date + a.bill_no < b.date + b.bill_no ? -1 : 1)) };
}

async function loadReport() {
  const from = $('#rep-from').value || todayISO(), to = $('#rep-to').value || from;
  const branch = isAdmin() ? $('#rep-branch').value : S.user.branch;
  $('#rep-msg').textContent = 'กำลังโหลด…';
  let bills, source;
  try { ({ bills, source } = await fetchBills(branch, from, to)); } catch (e) { $('#rep-msg').textContent = e.message; return; }
  S.report = { from, to, branch, bills };
  $('#rep-msg').textContent = source + ` · ${bills.length} บิล`;
  renderReport();
}

function reportRows(bills, includeVoid) {
  const rows = [];
  bills.forEach((b) => {
    if (!includeVoid && b.status === 'ยกเลิก') return;
    (b.lines || []).forEach((l, i) => {
      if (!includeVoid && l.status === 'ยกเลิก') return;
      rows.push({ b, l, first: i === 0 });
    });
  });
  return rows;
}

function renderReport() {
  const R = S.report; if (!R) return;
  const live = R.bills.filter((b) => b.status !== 'ยกเลิก');
  const sum = (f) => round2(live.reduce((s, b) => s + num(b[f]), 0));
  const sales = live.filter((b) => b.type !== 'คืน'), rets = live.filter((b) => b.type === 'คืน');
  let profitHtml = '';
  if (isAdmin()) {
    let cost = 0, missing = 0;
    live.forEach((b) => (b.lines || []).forEach((l) => { if (l.cost === '' || l.cost == null) missing++; else cost += num(l.cost) * num(l.qty); }));
    const netEx = sum('net') / 1.07;
    profitHtml = `<div><span>ทุนรวม (ไม่รวมภาษี)</span><b>${money(cost)}</b></div><div><span>ขายสุทธิไม่รวมภาษี</span><b>${money(netEx)}</b></div>
      <div><span>กำไร (ไม่รวมภาษี)</span><b>${money(netEx - cost)}</b></div>${missing ? `<div><span>บรรทัดที่ไม่มีทุน</span><b>${missing}</b></div>` : ''}`;
  }
  $('#rep-summary').innerHTML = `<div class="summary">
    <div><span>บิลขาย</span><b>${sales.length}</b></div><div><span>บิลคืน</span><b>${rets.length}</b></div>
    <div><span>ยกเลิก</span><b>${R.bills.length - live.length}</b></div>
    <div><span>รวมเงิน</span><b>${money(sum('total'))}</b></div><div><span>ส่วนลด</span><b>${money(sum('discount'))}</b></div>
    <div><span>เหลือ (สุทธิ)</span><b>${money(sum('net'))}</b></div>${profitHtml}</div>`;
  $('#rep-table tbody').innerHTML = reportRows(R.bills, true).map(({ b, l, first }) => {
    const cls = b.status === 'ยกเลิก' || l.status === 'ยกเลิก' ? 'void' : b.type === 'คืน' ? 'ret' : '';
    const st = b.status === 'ยกเลิก' ? 'ยกเลิก' : b.type === 'คืน' ? 'คืน' : '';
    const sync = b._sync === 'pending' ? ' <span class="badge pending">รอส่ง</span>' : b._sync === 'error' ? ' <span class="badge error">ไม่สำเร็จ</span>' : '';
    return `<tr class="clickable ${cls}" data-bill="${esc(b.bill_no)}"><td>${first ? esc(b.bill_no) : ''}</td><td>${first ? thDate(b.date) : ''}</td>
      <td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="num">${l.qty}</td><td class="num">${money(l.unit)}</td><td class="num">${money(l.amount)}</td>
      <td>${first ? esc(b.customer_name) : ''}</td><td class="num">${first ? money(b.total) : ''}</td><td class="num">${first ? money(b.discount) : ''}</td>
      <td class="num">${first ? money(b.net) : ''}</td><td>${first ? st + sync : ''}</td><td>${first ? esc([payText(b), orderText(b, false)].filter(Boolean).join(' · ')) : ''}</td></tr>`;
  }).join('') || '<tr><td colspan="14" class="muted">ไม่มีบิล</td></tr>';
}

/* ---------------- สะสมแต้มลูกค้า (ไฟล์อัปโหลดเข้าแอปสะสมแต้ม แบบเดียวกับ add-point-template) */

const POINT_MIN = 100;   // ยอดต่อวันต่ำกว่านี้ไม่นับ
const POINT_RATE = 100;  // ทุก 100 บาทของยอดนับแต้มในวันนั้น = 1 แต้ม (เศษปัดทิ้ง)
const dayPoints = (amount) => Math.max(0, Math.floor(round2(amount) / POINT_RATE));
function pointBranches() { return String((S.settings && S.settings.points_branches) || 'HQ,BT').split(',').map((x) => x.trim()).filter(Boolean); }
function pointStart(branch) { return branch === 'HQ' ? String((S.settings && S.settings.points_start) || '2026-10-08') : ''; }
function addDays(iso, n) { const d = new Date(iso + 'T00:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function ptsNoPoint(l) { if (/^ค่า/.test(String(l.name || '').trim())) return true; const p = findPart(l.code); return !!(p && p.noPoint); }
function ptsWhy(l) { if (/^ค่า/.test(String(l.name || '').trim())) return 'ค่าแรง/ค่าบริการ'; const p = findPart(l.code); return p && p.noPoint ? 'อะไหล่ไม่นับแต้ม' : ''; }

// ยอดนับแต้มของบิล = ยอดแต่ละรายการหลังเฉลี่ยส่วนลดท้ายบิล ไม่รวมค่าแรง/อะไหล่ไม่นับแต้ม (บิลคืนติดลบ)
function billFactor(b) { const total = num(b.total); return total ? num(b.net) / total : 1; }
function billPointAmount(b) {
  const f = billFactor(b);
  return (b.lines || []).reduce((s, l) => s + (ptsNoPoint(l) ? 0 : num(l.amount) * f), 0);
}

// แต้มคิดรายวันเสมอ (ไม่รวมข้ามวัน) ช่วงวันที่ที่เลือกใช้แค่ดึงหลายวันมาดูพร้อมกัน, cus = id ลูกค้า (ว่าง = สมาชิกทุกคน)
async function computePoints(branch, from, to, cus = '') {
  const start = pointStart(branch);
  if (start && from < start) from = start;
  if (from > to) return { rows: [], days: [], notes: [], from, to };
  const [{ bills }, customers] = await Promise.all([fetchBills(branch, from, to), api('customers', { branch })]);
  const byId = new Map(customers.map((c) => [c.id, c])), byName = new Map(customers.map((c) => [c.display, c]));
  const dateOf = new Map(bills.map((b) => [b.bill_no, b.date]));
  const day = new Map(), notes = [];
  for (const b of bills) {
    if (b.status !== 'ปกติ') continue;
    const c = byId.get(b.customer_id) || byName.get(b.customer_name);
    if (!isMember(c) || (cus && c.id !== cus)) continue;
    const phone = mobileOf(c.phone);
    if (!phone) { notes.push(`${b.bill_no} ${c.display}: เป็นสมาชิกแต่ไม่มีเบอร์มือถือ 10 หลัก`); continue; }
    let d = b.date;
    // บิลคืนหักจากวันของบิลขายเดิม (ถ้าบิลเดิมอยู่ในช่วงวันที่เลือก)
    if (b.type === 'คืน') {
      if (dateOf.has(b.ref_bill)) d = dateOf.get(b.ref_bill);
      else notes.push(`${b.bill_no} คืนจากบิล ${b.ref_bill} ซึ่งอยู่นอกช่วงวันที่เลือก หักจากวันที่คืน ${thDate(b.date)} แทน ถ้าวันนั้นส่งแต้มไปแล้ว ต้องหักในแอปสะสมแต้มเอง`);
    }
    const k = phone + '|' + d;
    const r = day.get(k) || { date: d, phone, name: c.display, cid: c.id, amount: 0, bills: [], list: [] };
    r.amount += billPointAmount(b); r.bills.push(b.bill_no); r.list.push(b);
    day.set(k, r);
  }
  const days = Array.from(day.values()).map((r) => Object.assign(r, { amount: round2(r.amount), ok: round2(r.amount) >= POINT_MIN, pts: round2(r.amount) >= POINT_MIN ? dayPoints(r.amount) : 0 }))
    .sort((a, b) => (a.date + a.name < b.date + b.name ? -1 : 1));
  const rows = days.filter((r) => r.ok);
  return { rows, days, dropped: days.length - rows.length, notes, from, to };
}

// รวมยอดหลายวันต่อคน (เฉพาะวันที่ถึง 100 บาท)
function pointsByCustomer(rows) {
  const m = new Map();
  rows.forEach((r) => { const x = m.get(r.phone) || { name: r.name, phone: r.phone, days: 0, amount: 0, pts: 0 }; x.days++; x.amount = round2(x.amount + r.amount); x.pts += r.pts; m.set(r.phone, x); });
  return Array.from(m.values()).sort((a, b) => b.amount - a.amount);
}

function thDateFull(iso) { const [y, m, d] = String(iso).slice(0, 10).split('-'); return `${num(d)} ${TH_MONTHS[num(m) - 1]} ${num(y) + 543}`; }
function ptsRange() { const from = $('#pts-from').value || todayISO(), to = $('#pts-to').value || from; return from <= to ? { from, to } : { from: to, to: from }; }
function ptsKey() { const r = ptsRange(); return [$('#pts-branch').value, r.from, r.to, $('#pts-cus').value].join('|'); }

async function loadPointMembers() {
  const sel = $('#pts-cus'), branch = $('#pts-branch').value;
  sel.innerHTML = '<option value="">ลูกค้าสมาชิกทั้งหมด</option>';
  try {
    const list = (await api('customers', { branch })).filter(isMember).sort((a, b) => String(a.display).localeCompare(String(b.display), 'th'));
    sel.innerHTML += list.map((c) => `<option value="${esc(c.id)}">${esc(c.display)}${mobileOf(c.phone) ? ' · ' + esc(mobileOf(c.phone)) : ''}</option>`).join('');
  } catch (e) { /* ออฟไลน์: เลือกได้แค่ทั้งหมด */ }
}

async function loadPoints() {
  const branch = $('#pts-branch').value, w = ptsRange(), key = ptsKey();
  $('#pts-msg').textContent = 'กำลังคำนวณ…'; $('#pts-table tbody').innerHTML = ''; $('#pts-sum tbody').innerHTML = ''; $('#pts-notes').innerHTML = '';
  try {
    const r = await computePoints(branch, w.from, w.to, $('#pts-cus').value);
    S.points = Object.assign({ branch, key, range: w }, r);
    const sum = r.rows.reduce((s, x) => s + x.amount, 0);
    const label = w.from === w.to ? thDateLong(w.from) : `${thDateLong(w.from)} - ${thDateLong(w.to)}`;
    $('#pts-msg').textContent = `${branchInfo(branch).name} · ${label}` + (r.from !== w.from ? ` (เริ่มนับ ${thDateLong(r.from)})` : '')
      + ` · ${r.rows.length} แถว ยอดรวม ${money(sum)} บาท · ${r.rows.reduce((s, x) => s + x.pts, 0).toLocaleString()} แต้ม` + (r.dropped ? ` · ไม่ถึง ${POINT_MIN} บาท ${r.dropped} วัน (ไม่นับ)` : '');
    const by = pointsByCustomer(r.rows);
    $('#pts-sum tbody').innerHTML = by.map((x) => `<tr><td>${esc(x.name)}</td><td>${esc(x.phone)}</td><td class="num">${x.days}</td><td class="num">${money(x.amount)}</td><td class="num"><b>${x.pts.toLocaleString()}</b></td></tr>`).join('')
      + (by.length > 1 ? `<tr><td colspan="3"><b>รวมทั้งหมด ${by.length} คน</b></td><td class="num"><b>${money(sum)}</b></td><td class="num"><b>${by.reduce((s, x) => s + x.pts, 0).toLocaleString()}</b></td></tr>` : '')
      || '<tr><td colspan="5" class="muted">ไม่มีลูกค้าสมาชิกที่ได้แต้มในช่วงนี้</td></tr>';
    $('#pts-table tbody').innerHTML = r.rows.map((x) => `<tr><td>${esc(thDateFull(x.date))}</td><td>${esc(x.name)}</td><td>${esc(x.phone)}</td>
      <td class="num">${money(x.amount)}</td><td class="num"><b>${x.pts}</b></td><td class="small">${x.bills.map((n) => `<a href="#" data-bill="${esc(n)}">${esc(n)}</a>`).join(' ')}</td></tr>`).join('')
      || '<tr><td colspan="6" class="muted">ไม่มีลูกค้าสมาชิกที่ได้แต้มในช่วงนี้</td></tr>';
    $('#pts-notes').innerHTML = r.notes.map((n) => `<p class="warn">⚠️ ${esc(n)}</p>`).join('');
  } catch (e) {
    $('#pts-msg').textContent = e instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : 'คำนวณไม่สำเร็จ: ' + e.message;
  }
}

async function ptsReady() {
  if (!S.points || S.points.key !== ptsKey()) await loadPoints();
  if (!window.XLSX) { toast('ตัวสร้าง Excel ยังโหลดไม่เสร็จ'); return null; }
  return S.points && S.points.days ? S.points : null;
}
function ptsFileTag(P) { return `${branchInfo(P.branch).name}_${thDateLong(P.from).replace(/ /g, '')}-${thDateLong(P.to).replace(/ /g, '')}`; }
const PTS_MONEY = '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-';

// ไฟล์สำหรับใส่แอพ: แบบเดียวกับ add-point-template (1 แถว = ลูกค้า 1 คน 1 วัน)
async function exportPoints() {
  const P = await ptsReady();
  if (!P) return;
  if (!P.rows.length) return toast('ไม่มีรายการให้ส่งออก');
  const ws = XLSX.utils.aoa_to_sheet([['Phone Number', 'Billing ID', 'Money Amount']].concat(P.rows.map((x) => [x.phone, thDateFull(x.date), x.amount])));
  P.rows.forEach((x, i) => {
    ws['A' + (i + 2)] = { t: 's', v: x.phone, z: '@' };
    ws['C' + (i + 2)].z = PTS_MONEY;
  });
  ws['!cols'] = [{ wch: 14 }, { wch: 20 }, { wch: 14 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Dates');
  XLSX.writeFile(wb, `ไฟล์สำหรับใส่แอพ_${ptsFileTag(P)}.xlsx`);
}

// ข้อมูลexport: รวมตามลูกค้า + รายวัน (รวมวันที่ไม่ถึง 100) + ทุกรายการในบิล
async function exportPointsData() {
  const P = await ptsReady();
  if (!P) return;
  if (!P.days.length) return toast('ไม่มีรายการให้ส่งออก');
  const money2 = (ws, cols, n) => { for (let i = 2; i <= n + 1; i++) cols.forEach((c) => { if (ws[c + i] && ws[c + i].t === 'n') ws[c + i].z = PTS_MONEY; }); };
  const by = pointsByCustomer(P.rows);
  const s1 = XLSX.utils.aoa_to_sheet([['ลูกค้า', 'เบอร์มือถือ', 'จำนวนวันที่นับ', 'ยอดนับแต้มรวม', 'แต้ม']].concat(by.map((x) => [x.name, x.phone, x.days, x.amount, x.pts])));
  money2(s1, ['D'], by.length); s1['!cols'] = [{ wch: 30 }, { wch: 14 }, { wch: 12 }, { wch: 16 }, { wch: 8 }];
  const s2 = XLSX.utils.aoa_to_sheet([['วันที่', 'ลูกค้า', 'เบอร์มือถือ', 'ยอดนับแต้ม', 'แต้ม', 'นับแต้ม', 'บิล']].concat(P.days.map((x) =>
    [thDateFull(x.date), x.name, x.phone, x.amount, x.pts, x.ok ? 'นับ' : `ไม่นับ (ไม่ถึง ${POINT_MIN})`, x.bills.join(' ')])));
  money2(s2, ['D'], P.days.length); s2['!cols'] = [{ wch: 18 }, { wch: 30 }, { wch: 14 }, { wch: 14 }, { wch: 8 }, { wch: 16 }, { wch: 30 }];
  const lines = [['วันที่นับ', 'เลขบิล', 'วันที่บิล', 'ประเภท', 'ลูกค้า', 'เบอร์มือถือ', 'รหัสอะไหล่', 'ชื่ออะไหล่', 'จำนวน', 'จำนวนเงิน', 'ส่วนลดเฉลี่ย', 'ยอดนับแต้ม', 'หมายเหตุ']];
  P.days.forEach((x) => x.list.forEach((b) => {
    const f = billFactor(b);
    (b.lines || []).forEach((l) => {
      const why = ptsWhy(l), amt = round2(num(l.amount));
      lines.push([thDateFull(x.date), b.bill_no, thDate(b.date), b.type || 'ขาย', x.name, x.phone, l.code, l.name, num(l.qty), amt,
        round2(amt - amt * f), why ? 0 : round2(amt * f), why ? 'ไม่นับ: ' + why : (x.ok ? '' : `วันนี้ไม่ถึง ${POINT_MIN} บาท`)]);
    });
  }));
  const s3 = XLSX.utils.aoa_to_sheet(lines);
  for (let i = 2; i <= lines.length; i++) { s3['F' + i] = { t: 's', v: lines[i - 1][5], z: '@' }; }
  money2(s3, ['J', 'K', 'L'], lines.length - 1);
  s3['!cols'] = [{ wch: 18 }, { wch: 12 }, { wch: 11 }, { wch: 7 }, { wch: 26 }, { wch: 13 }, { wch: 16 }, { wch: 30 }, { wch: 7 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 22 }];
  [[s1, 'B', by.length], [s2, 'C', P.days.length]].forEach(([ws, c, n]) => { for (let i = 2; i <= n + 1; i++) if (ws[c + i]) ws[c + i] = { t: 's', v: String(ws[c + i].v), z: '@' }; });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, s1, 'รวมตามลูกค้า');
  XLSX.utils.book_append_sheet(wb, s2, 'รายวัน');
  XLSX.utils.book_append_sheet(wb, s3, 'รายการในบิล');
  XLSX.writeFile(wb, `ข้อมูลexport_${ptsFileTag(P)}.xlsx`);
}

function bindPoints() {
  const sel = $('#pts-branch');
  sel.innerHTML = pointBranches().map((c) => `<option value="${esc(c)}">${esc(branchInfo(c).name)}</option>`).join('');
  if (!$('#pts-from').value) { $('#pts-from').value = addDays(todayISO(), -1); $('#pts-to').value = addDays(todayISO(), -1); }
  sel.onchange = () => { loadPointMembers(); };
  $('#btn-pts-load').onclick = loadPoints;
  $('#btn-pts-excel').onclick = exportPoints;
  $('#btn-pts-export').onclick = exportPointsData;
  $('#pts-table').onclick = (e) => { const a = e.target.closest('[data-bill]'); if (a) { e.preventDefault(); openBill(a.dataset.bill); } };
}

/* ---------------- แอดมิน > อะไหล่ไม่นับแต้ม */

function renderNoPoint() {
  const q = $('#np-filter').value.trim().toLowerCase();
  const all = S.partsList.filter((p) => p.noPoint).sort((a, b) => (String(a.code) < String(b.code) ? -1 : 1));
  const list = q ? all.filter((p) => [p.code, p.name, p.model].some((v) => String(v || '').toLowerCase().includes(q))) : all;
  $('#np-count').textContent = `(${all.length.toLocaleString()} รายการ${q ? ` · ค้นเจอ ${list.length}` : ''})`;
  $('#np-table tbody').innerHTML = list.map((p) => `<tr><td><input type="checkbox" checked data-np="${esc(p.code)}"></td><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td>${esc(p.model)}</td><td class="num">${money(p.price)}</td></tr>`).join('')
    || '<tr><td colspan="5" class="muted">ไม่มีรายการ</td></tr>';
}

async function setNoPoint(codes, value) {
  const r = await api('setNoPoint', { codes, value }, 'กำลังบันทึก');
  await syncParts();
  renderNoPoint();
  return r;
}

async function addNoPoint(text) {
  const raw = String(text || '').split(/[\s,]+/).map((x) => x.trim()).filter(Boolean);
  if (!raw.length) return toast('ใส่รหัสอะไหล่ก่อน');
  const found = [], missing = [];
  raw.forEach((c) => { const p = findPart(c); if (p) found.push(p.code); else missing.push(c); });
  if (!found.length) return toast('ไม่พบรหัส ' + missing.join(', '));
  try {
    await setNoPoint(found, true);
    $('#np-add-code').value = '';
    toast(`ตั้งไม่นับแต้มแล้ว ${found.length} รายการ` + (missing.length ? ` · ไม่พบรหัส ${missing.join(', ')}` : ''));
  } catch (e) { toast(e instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : e.message); }
}

function bindNoPoint() {
  $('#np-filter').oninput = renderNoPoint;
  $('#btn-np-add').onclick = () => addNoPoint($('#np-add-code').value);
  $('#np-add-code').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); addNoPoint(e.target.value); } };
  $('#btn-np-search').onclick = () => openPartSearch('', (p) => { const v = $('#np-add-code').value.trim(); $('#np-add-code').value = (v ? v + ' ' : '') + p.code; }, { quickAdd: false });
  $('#np-table').onchange = async (e) => {
    const cb = e.target.closest('[data-np]');
    if (!cb || cb.checked) return;
    try { await setNoPoint([cb.dataset.np], false); toast(`${cb.dataset.np} กลับมานับแต้มแล้ว`); }
    catch (err) { cb.checked = true; toast(err instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : err.message); }
  };
}

function exportExcel() {
  const R = S.report;
  if (!R) return toast('กด "แสดง" ก่อน');
  if (!window.XLSX) return toast('ตัวสร้าง Excel ยังโหลดไม่เสร็จ ลองใหม่อีกครั้ง');
  const pending = R.bills.filter((b) => b._sync === 'pending').length;
  if (pending && !confirm(`มี ${pending} บิลที่ยังไม่ได้ส่งขึ้นระบบ จะดาวน์โหลดรวมไว้ด้วยไหม?`)) return;
  const head = ['ใบเสร็จเลขที่', 'วันที่', 'รหัสอะไหล่', 'ชื่ออะไหล่', 'รุ่น', 'จำนวน', 'หน่วยละ', 'จำนวนเงิน', 'ชื่อผู้ซื้อ', 'รวมเงิน', 'ส่วนลด', 'เหลือ', 'สาขา', 'ประเภท', 'อ้างอิงบิล', 'ชำระโดย', 'รายละเอียดการโอน', 'หมายเหตุ', 'ผู้อนุมัติ'];
  const data = [head];
  reportRows(R.bills, false).forEach(({ b, l, first }) => {
    data.push([b.bill_no, thDate(b.date), l.code, l.name, l.model, num(l.qty), round2(num(l.unit)), round2(num(l.amount)),
      first ? b.customer_name : '', first ? round2(num(b.total)) : '', first ? round2(num(b.discount)) : '', first ? round2(num(b.net)) : '',
      branchInfo(b.branch).name, b.type, b.ref_bill || '', b.pay_method || '', b.pay_ref || '', orderText(b), b.approved_by || '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [12, 11, 18, 34, 16, 7, 10, 11, 26, 10, 9, 10, 12, 7, 12, 9, 26, 20, 10].map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'บันทึกขายอะไหล่');
  const bname = R.branch === '*' ? 'ทุกสาขา' : branchInfo(R.branch).name;
  const dname = R.from === R.to ? R.from : R.from + '_ถึง_' + R.to;
  XLSX.writeFile(wb, `ขายอะไหล่_${bname}_${dname}.xlsx`);
}

/* ================================================================ ใบคุมบิลรายวัน (ฝ่ายบัญชี) */

// เลขแบบใหม่ PA1-00001: หัวเลขเดียวกัน = เครื่องเดียวกัน
// เลขแบบเดิม PA6900001: แต่ละเครื่องจองเลขไว้ชุดละ NUMBER_BLOCK เลข ชุดเลขเดียวกัน = เครื่องเดียวกัน
function billSeq(no) { return num(String(no).slice(-5)); }
function billDevice(no) { const m = /(\d{1,2})-\d{5}$/.exec(String(no)); return m ? Number(m[1]) : null; }
function billBlock(no) {
  if (billDevice(no) != null) return String(no).slice(0, -5);
  const n = billSeq(no); return String(no).slice(0, -5) + '|' + Math.floor((n - 1) / NUMBER_BLOCK);
}
function blockLabel(no, last) {
  const head = String(no).slice(0, -5);
  if (billDevice(no) != null) return `${no}${last && last !== no ? ' - ' + last : ''}`;
  const n = billSeq(no), a = Math.floor((n - 1) / NUMBER_BLOCK) * NUMBER_BLOCK + 1;
  return `${head}${pad(a, 5)} - ${head}${pad(a + NUMBER_BLOCK - 1, 5)}`;
}
// เงินที่ได้รับจากบิล: ทุกบิลรับเต็มยอด บิลคืน = จ่ายคืน (ติดลบ) บิลยกเลิก = 0
function billCash(b) {
  if (b.status === 'ยกเลิก') return 0;
  return b.type === 'คืน' ? -num(b.net) : num(b.net);
}
// แยกเงินสด / เงินโอน / ค้างจ่าย (บิลเก่าที่ไม่ได้ระบุ และบิลคืน นับเป็นเงินสด)
function billSplit(b) {
  const c = billCash(b);
  if (b.pay_method === PAY_CREDIT) return { cash: 0, transfer: 0, credit: c };
  return b.pay_method === 'โอนเงิน' ? { cash: 0, transfer: c, credit: 0 } : { cash: c, transfer: 0, credit: 0 };
}

function controlSheetHtml(R) {
  const days = new Map(), add = (k) => { if (!days.has(k)) days.set(k, { bills: [], settled: [] }); return days.get(k); };
  R.bills.forEach((b) => add(b.date + '|' + b.branch).bills.push(b));
  (R.settled || []).forEach((b) => add(b.settled_date + '|' + b.branch).settled.push(b));
  const keys = Array.from(days.keys()).sort();
  if (!keys.length) return '';
  const m = (v) => (v ? money(v) : '');
  return keys.map((k) => {
    const [date, branch] = k.split('|'), { bills, settled } = days.get(k);
    const blocks = new Map();
    bills.slice().sort((a, b) => (a.bill_no < b.bill_no ? -1 : 1)).forEach((b) => { const g = billBlock(b.bill_no); if (!blocks.has(g)) blocks.set(g, []); blocks.get(g).push(b); });
    const T = { net: 0, cash: 0, transfer: 0, credit: 0, nCash: 0, nTransfer: 0, nCredit: 0, n: 0, void: 0, ret: 0, retAmt: 0, gaps: 0, sCash: 0, sTransfer: 0 };
    let i = 0, n = 0;
    const body = Array.from(blocks.values()).map((list) => {
      n++;
      const devNo = billDevice(list[0].bill_no) != null ? billDevice(list[0].bill_no) : n;
      const S2 = { net: 0, cash: 0, transfer: 0, credit: 0 };
      let prev = null, rows = '';
      list.forEach((b) => {
        const seq = billSeq(b.bill_no);
        if (prev != null && seq - prev > 1) {
          const miss = seq - prev - 1; T.gaps += miss;
          rows += `<tr class="cs-gap"><td></td><td colspan="11">⚠ ไม่พบเลข ${esc(String(b.bill_no).slice(0, -5) + pad(prev + 1, 5))}${miss > 1 ? ' ถึง ' + esc(String(b.bill_no).slice(0, -5) + pad(seq - 1, 5)) : ''} (${miss} เลข) ให้ตรวจสอบเครื่องว่าส่งข้อมูลครบหรือไม่</td></tr>`;
        }
        prev = seq;
        const dead = b.status === 'ยกเลิก', ret = b.type === 'คืน', sp = billSplit(b);
        if (dead) T.void++;
        else {
          T.n++; if (ret) { T.ret++; T.retAmt += num(b.net); }
          S2.net += ret ? -num(b.net) : num(b.net); S2.cash += sp.cash; S2.transfer += sp.transfer; S2.credit += sp.credit;
          if (!ret) { if (sp.credit) T.nCredit++; else if (sp.transfer) T.nTransfer++; else T.nCash++; }
        }
        const kind = dead ? 'ยกเลิก' : ret ? 'คืน' + (b.ref_bill ? ' (' + esc(b.ref_bill) + ')' : '') : [b.pay_method === PAY_CREDIT ? 'ขายค้างจ่าย' : '', b.order_type ? esc(orderText(b)) : ''].filter(Boolean).join(' · ') || 'ขายสด';
        rows += `<tr class="${dead ? 'cs-void' : ''}"><td class="c">${++i}</td><td>${esc(b.bill_no)}</td><td>${esc(b.time || '')}</td><td>${esc(b.customer_name)}</td><td>${kind}</td>
          <td class="num">${dead ? '-' : money(ret ? -num(b.net) : b.net)}</td><td class="num">${dead ? '-' : m(sp.cash)}</td>
          <td class="num">${dead ? '-' : m(sp.transfer)}${b.pay_method === 'โอนเงิน' && b.pay_ref ? `<div class="cs-ref">${esc(b.pay_ref)}</div>` : ''}</td>
          <td class="num">${dead ? '-' : m(sp.credit)}${sp.credit && b.settled_date ? `<div class="cs-ref">ชำระแล้ว ${thDate(b.settled_date)}</div>` : ''}</td>
          <td>${esc(b.created_by || '')}</td><td>${esc(b.approved_by || '')}</td><td></td></tr>`;
      });
      T.net += S2.net; T.cash += S2.cash; T.transfer += S2.transfer; T.credit += S2.credit;
      const users = Array.from(new Set(list.map((b) => b.created_by).filter(Boolean))).join(', ');
      return `<tr class="cs-dev"><td colspan="12">เครื่องที่ ${devNo} · ${billDevice(list[0].bill_no) != null ? 'เลขที่' : 'ชุดเลข'} ${esc(blockLabel(list[0].bill_no, list[list.length - 1].bill_no))} · ${list.length} บิล${users ? ' · ผู้ขาย ' + esc(users) : ''}</td></tr>${rows}
        <tr class="cs-sub"><td colspan="5">รวมเครื่องที่ ${devNo}</td><td class="num">${money(S2.net)}</td><td class="num">${money(S2.cash)}</td><td class="num">${money(S2.transfer)}</td><td class="num">${money(S2.credit)}</td><td colspan="3"></td></tr>`;
    }).join('');
    // รับชำระบิลค้างจ่ายในวันนี้ (บิลของวันก่อน หรือวันนี้ก็ได้)
    let settledHtml = '';
    if (settled.length) {
      settledHtml = `<div class="cs-subtitle">รับชำระบิลค้างจ่าย (${settled.length} บิล)</div>
        <table class="cs-table"><thead><tr><th>เลขที่บิล</th><th>วันที่ขาย</th><th>ลูกค้า</th><th class="num">เงินสด</th><th class="num">เงินโอน</th><th>ผู้รับเงิน</th><th>ตรวจ</th></tr></thead><tbody>
        ${settled.map((b) => { const v = num(b.net), tr = b.settled_method === 'โอนเงิน'; if (tr) T.sTransfer += v; else T.sCash += v;
          return `<tr><td>${esc(b.bill_no)}</td><td>${thDate(b.date)}</td><td>${esc(b.customer_name)}</td><td class="num">${tr ? '' : money(v)}</td>
            <td class="num">${tr ? money(v) : ''}${tr && b.settled_ref ? `<div class="cs-ref">${esc(b.settled_ref)}</div>` : ''}</td><td>${esc(b.settled_by || '')}</td><td></td></tr>`; }).join('')}
        </tbody><tfoot><tr class="cs-total"><td colspan="3">รวมรับชำระ</td><td class="num">${money(T.sCash)}</td><td class="num">${money(T.sTransfer)}</td><td colspan="2"></td></tr></tfoot></table>`;
    }
    const br = branchInfo(branch);
    const cashAll = T.cash + T.sCash, transferAll = T.transfer + T.sTransfer;
    return `<div class="cs-doc">
      <div class="cs-head"><div><div class="cs-title">ใบคุมบิลขายอะไหล่ประจำวัน</div><div>${esc(S.settings.shop_name || 'หจก.ประสบชัยกลการ')} · ${esc(branchTitle(br))}</div></div>
        <div class="cs-date">วันที่ <b>${thDate(date)}</b></div></div>
      ${bills.length ? `<table class="cs-table"><thead><tr><th>#</th><th>เลขที่บิล</th><th>เวลา</th><th>ลูกค้า</th><th>ประเภท</th><th class="num">ยอดบิล</th><th class="num">เงินสด</th><th class="num">เงินโอน</th><th class="num">ค้างจ่าย</th><th>ผู้ขาย</th><th>ผู้อนุมัติ</th><th>ตรวจ</th></tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr class="cs-total"><td colspan="5">รวมทั้งวัน</td><td class="num">${money(T.net)}</td><td class="num">${money(T.cash)}</td><td class="num">${money(T.transfer)}</td><td class="num">${money(T.credit)}</td><td colspan="3"></td></tr></tfoot></table>` : '<p class="muted">ไม่มีบิลขายในวันนี้</p>'}
      ${settledHtml}
      <div class="cs-eod">
        <div class="cs-eod-title">สรุปยอดสิ้นวัน</div>
        <table>
          <tr><td colspan="2">บิลขาย ${T.n - T.ret} ใบ · บิลคืน ${T.ret} ใบ · ยกเลิก ${T.void} ใบ${T.gaps ? ` · <b class="cs-warn">เลขที่หายไป ${T.gaps} เลข</b>` : ''}</td></tr>
          <tr><td>ขายรับเงินสด (${T.nCash} บิล)${T.retAmt ? ` หักคืนเงิน ${money(T.retAmt)}` : ''}</td><td class="num">${money(T.cash)}</td></tr>
          <tr><td>ขายรับเงินโอน (${T.nTransfer} บิล)</td><td class="num">${money(T.transfer)}</td></tr>
          ${T.nCredit ? `<tr class="cs-muted"><td>ขายค้างจ่าย (${T.nCredit} บิล) ยังไม่ได้รับเงิน</td><td class="num">${money(T.credit)}</td></tr>` : ''}
          ${settled.length ? `<tr><td>รับชำระบิลค้าง เงินสด / เงินโอน</td><td class="num">${money(T.sCash)} / ${money(T.sTransfer)}</td></tr>` : ''}
          <tr class="cs-eod-sum"><td>รวมรับเงินทั้งวัน (เงินสด ${money(cashAll)} + โอน ${money(transferAll)})</td><td class="num">${money(cashAll + transferAll)}</td></tr>
          <tr class="cs-eod-big"><td>เงินสดที่ต้องส่งฝ่ายบัญชี</td><td class="num">${money(cashAll)}</td></tr>
        </table>
      </div>
      ${(R.credits || []).some((b) => b.date <= date) ? creditListHtml(R.credits.filter((b) => b.date <= date), date) : ''}
      <div class="cs-count">นับเงินสดได้จริง ...................................... บาท &nbsp;&nbsp; ขาด/เกิน ...................................... บาท &nbsp;&nbsp; ตรวจยอดโอนกับบัญชีธนาคารแล้ว ☐</div>
      <div class="cs-sign"><div>ลงชื่อ ......................................................<br>ผู้ส่งเงิน</div><div>ลงชื่อ ......................................................<br>ฝ่ายบัญชีผู้รับเงิน</div></div>
    </div>`;
  }).join('');
}

// บิลค้างจ่ายที่ยังไม่ได้รับเงิน (ณ วันที่พิมพ์)
function daysBetween(a, b) { return Math.round((new Date(b + 'T00:00:00') - new Date(a + 'T00:00:00')) / 864e5); }
function creditListHtml(list, asOf) {
  const sum = list.reduce((s, b) => s + num(b.net), 0);
  return `<div class="cs-subtitle">บิลค้างจ่ายที่ยังไม่ได้รับเงิน ${list.length} บิล รวม ${money(sum)} บาท</div>
    <table class="cs-table"><thead><tr><th>เลขที่บิล</th><th>วันที่ขาย</th><th class="num">ค้างมา (วัน)</th><th>ลูกค้า</th><th class="num">ยอด</th><th>ผู้ขาย</th><th>ผู้อนุมัติ</th></tr></thead><tbody>
    ${list.map((b) => `<tr class="clickable" data-bill="${esc(b.bill_no)}"><td>${esc(b.bill_no)}</td><td>${thDate(b.date)}</td><td class="num">${Math.max(0, daysBetween(b.date, asOf))}</td><td>${esc(b.customer_name)}</td>
      <td class="num">${money(b.net)}</td><td>${esc(b.created_by || '')}</td><td>${esc(b.approved_by || '')}</td></tr>`).join('')}</tbody></table>`;
}

async function loadControl() {
  try { await loadControlInner(); } catch (e) { $('#ctl-msg').textContent = 'โหลดใบคุมไม่ได้: ' + e.message; }
}

async function loadControlInner() {
  if (!$('#ctl-branch').options.length && S.branches.length) fillBranchSelects();
  const date = $('#ctl-date').value || todayISO();
  const branch = isAdmin() ? $('#ctl-branch').value || S.branch : S.user.branch;
  $('#ctl-msg').textContent = 'กำลังโหลด…';
  let r;
  try { r = await fetchBills(branch, date, date); } catch (e) { $('#ctl-msg').textContent = e.message; return; }
  // รับชำระบิลค้างวันนี้ + บิลค้างจ่ายคงค้าง (ออฟไลน์ใช้ข้อมูลในเครื่อง)
  const local = (await DB.all('bills')).filter((b) => b.branch === branch && b.pay_method === PAY_CREDIT && b.status === 'ปกติ');
  let settled = [], credits = [];
  try { settled = await api('settled', { branch, date }); credits = await api('credits', { branch }); } catch (e) { /* ออฟไลน์ */ }
  const merge = (arr, keep) => { const mp = new Map(arr.map((b) => [b.bill_no, b])); local.forEach((b) => { if (b._sync !== 'synced' || mp.has(b.bill_no)) mp.set(b.bill_no, b); }); return Array.from(mp.values()).filter(keep); };
  settled = merge(settled, (b) => b.settled_date === date);
  credits = merge(credits, (b) => !b.settled_date).sort((a, b) => (a.bill_no < b.bill_no ? -1 : 1));
  S.control = { from: date, to: date, branch, bills: r.bills, settled, credits };
  $('#ctl-msg').textContent = r.source + ` · ${r.bills.length} บิล`;
  $('#ctl-preview').innerHTML = controlSheetHtml(S.control)
    || (credits.length ? `<div class="cs-doc"><p class="muted">ไม่มีบิลในวันนี้</p>${creditListHtml(credits, date)}</div>` : '<p class="muted">ไม่มีบิลในวันนี้</p>');
}

async function printControlSheet() {
  if (!S.control) await loadControl();
  if (!$('#ctl-preview .cs-doc')) return toast('ไม่มีบิลในวันนี้');
  const html = $('#ctl-preview').innerHTML;
  $('#print-area').innerHTML = html;
  $('#print-area').className = 'print-only';
  setTimeout(() => window.print(), 50);
}

/* ================================================================ แอดมิน */

async function loadAudit() {
  try {
    const rows = await api('audit', { limit: 300 });
    $('#audit-table tbody').innerHTML = rows.map((r) => {
      let d = r.detail;
      try { const o = JSON.parse(r.detail); d = o.reason ? 'เหตุผล: ' + o.reason : ''; if (o.total != null) d += ` ยอด ${money(o.total)}`; if (o.added != null) d = `เพิ่ม ${o.added} แก้ ${o.updated}`; if (o.username) d = o.username; } catch (e) { /* แสดงตามเดิม */ }
      return `<tr ${r.bill_no ? `class="clickable" data-bill="${esc(r.bill_no)}"` : ''}><td>${esc(String(r.ts).replace('T', ' '))}</td><td>${esc(r.user)}</td><td>${esc(branchInfo(r.branch).name)}</td>
        <td>${esc(r.action)}</td><td>${esc(r.bill_no)}</td><td>${esc(r.approved_by)}</td><td>${esc(d)}</td></tr>`;
    }).join('') || '<tr><td colspan="7" class="muted">ยังไม่มีประวัติ</td></tr>';
  } catch (e) { toast(e.message); }
}

function loadDiscountAdmin() {
  $('#branch-disc-body').innerHTML = S.branches.map((b) => `<tr data-code="${esc(b.code)}"><td>${esc(b.name)}</td>
    <td><input name="branch_no" value="${esc(b.branch_no)}" style="width:60px" placeholder="-"></td>
    <td><input name="address" value="${esc(b.address)}" style="min-width:320px"></td>
    <td><input name="phone" value="${esc(b.phone)}" style="width:120px"></td>
    <td class="num"><input type="number" step="0.01" min="0" max="100" name="discount_pct" value="${num(b.discount_pct)}" style="width:80px"></td></tr>`).join('');
  $('#branch-disc-msg').textContent = '';
  const odd = S.partsList.filter((p) => p.maxDisc !== '' && p.maxDisc != null).sort((a, b) => a.code < b.code ? -1 : 1);
  $('#special-parts-body').innerHTML = odd.map((p) => `<tr><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td class="num">${money(p.price)}</td>
    <td class="num">${num(p.maxDisc) ? num(p.maxDisc) + '%' : 'ลดไม่ได้'}</td><td><button type="button" data-edit-part="${esc(p.code)}">แก้ไข</button></td></tr>`).join('')
    || '<tr><td colspan="5" class="muted">ยังไม่มี อะไหล่ทุกตัวลดตาม % ของสาขา</td></tr>';
}

async function saveBranchDiscounts(ev) {
  ev.preventDefault();
  const branches = [];
  for (const tr of $$('#branch-disc-body tr')) {
    const g = (n) => $(`[name=${n}]`, tr).value;
    const v = num(g('discount_pct'));
    if (g('discount_pct') === '' || v < 0 || v > 100) return toast('ส่วนลดต้องเป็น 0 ถึง 100%');
    branches.push({ code: tr.dataset.code, discount_pct: v, branch_no: g('branch_no').trim(), address: g('address').trim(), phone: g('phone').trim() });
  }
  try {
    const r = await api('saveBranches', { branches });
    S.branches = r.branches;
    const sess = await DB.kv.get('session');
    if (sess) { sess.branches = r.branches; await DB.kv.set('session', sess); }
    loadDiscountAdmin();
    $('#branch-disc-msg').textContent = 'บันทึกแล้ว เครื่องอื่นจะได้ค่าใหม่ตอนเปิดแอปครั้งถัดไป';
    renderCart();
  } catch (e) { toast(e instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อนจึงจะแก้ข้อมูลสาขาได้' : e.message, 6000); }
}

function renderPendingParts() {
  const rows = S.partsList.filter((p) => p.status === PART_PENDING);
  const keep = new Set($$('#pending-parts-body [data-pp]:checked').map((c) => c.dataset.pp));
  $('#pending-parts-card').classList.toggle('hidden', !rows.length);
  $('#pending-parts-count').textContent = rows.length ? `(${rows.length} รายการ)` : '';
  $('#pending-parts-body').innerHTML = rows.map((p) => `<tr><td><input type="checkbox" data-pp="${esc(p.code)}" ${keep.has(p.code) ? 'checked' : ''}></td>
    <td>${esc(p.code)}</td><td>${esc(p.name)}</td><td>${esc(p.model)}</td>
    <td class="num">${money(p.price)}</td><td>${esc(p.addedBy)}</td><td>${p.noPoint ? '<span class="danger-text">ไม่นับ</span>' : 'นับ'}</td>
    <td><button type="button" data-fix="${esc(p.code)}">แก้ไข</button></td></tr>`).join('');
  ppSelChanged();
}

function ppSelected() { return $$('#pending-parts-body [data-pp]:checked').map((c) => findPart(c.dataset.pp)).filter(Boolean); }
function ppSelChanged() {
  const all = $$('#pending-parts-body [data-pp]'), n = all.filter((c) => c.checked).length;
  $('#pp-selected').textContent = n ? `เลือกไว้ ${n} รายการ` : 'ยังไม่ได้เลือก';
  $('#pp-all').checked = !!all.length && n === all.length;
  ['#btn-pp-approve', '#btn-pp-nopoint', '#btn-pp-point', '#btn-pp-delete'].forEach((id) => { $(id).disabled = !n; });
}

// ยืนยันอะไหล่รอตรวจ: ส่งขึ้นระบบ แล้วแก้สถานะในเครื่องทันที (ไม่ต้องรอโหลดข้อมูลอะไหล่ใหม่)
async function approveParts(parts, btn) {
  const label = btn.textContent;
  btn.disabled = true; btn.textContent = 'กำลังยืนยัน…';
  try {
    await api('saveParts', { parts: parts.map((p) => ({ code: p.code, name: p.name, model: p.model, cost: p.cost, price_ex: '', price: p.price, max_disc: '', approve: true })) }, 'กำลังยืนยัน');
    parts.forEach((p) => { p.status = ''; p.maxDisc = ''; });
    await DB.putMany('parts', parts);
    toast(parts.length === 1 ? 'ยืนยันอะไหล่ ' + parts[0].code + ' แล้ว' : `ยืนยันอะไหล่ ${parts.length} รายการแล้ว`);
    renderPendingParts(); renderCart();
    syncParts().then(renderPendingParts);
  } catch (err) {
    toast(err instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : 'ยืนยันไม่สำเร็จ: ' + err.message, 6000);
  } finally { btn.textContent = label; ppSelChanged(); }
}

// ลบอะไหล่รอตรวจที่ไม่อนุมัติ (บิลที่ขายไปแล้วไม่กระทบ)
async function deletePendingParts(parts, btn) {
  if (!confirm(`ลบอะไหล่ ${parts.length} รายการนี้ออกจากระบบ?\n${parts.slice(0, 10).map((p) => p.code + ' ' + p.name).join('\n')}${parts.length > 10 ? '\n…' : ''}`)) return;
  btn.disabled = true;
  try {
    const r = await api('deleteParts', { codes: parts.map((p) => p.code) }, 'กำลังลบ');
    await removePartsLocal(r.deleted);
    toast(`ลบอะไหล่ ${r.deleted.length} รายการแล้ว` + (r.skipped.length ? ` · ${r.skipped.join(', ')} ยืนยันไปแล้ว ลบไม่ได้` : ''), 5000);
    renderPendingParts(); renderCart();
    syncParts().then(renderPendingParts);
  } catch (err) {
    toast(err instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : 'ลบไม่สำเร็จ: ' + err.message, 6000);
  } finally { ppSelChanged(); }
}

async function removePartsLocal(codes) {
  for (const c of codes || []) await DB.del('parts', codeKey(c));
  await loadPartsFromDb();
}

async function onPendingPartClick(e) {
  if (e.target.dataset.pp !== undefined) return ppSelChanged();
  const b = e.target.closest('button[data-fix]');
  const p = b && findPart(b.dataset.fix);
  if (p) fillPartForm(p);
}

function bindPendingParts() {
  $('#pending-parts-body').onclick = onPendingPartClick;
  $('#pp-all').onchange = (e) => { $$('#pending-parts-body [data-pp]').forEach((c) => { c.checked = e.target.checked; }); ppSelChanged(); };
  $('#btn-pp-approve').onclick = (e) => { const s = ppSelected(); if (s.length) approveParts(s, e.target); };
  $('#btn-pp-delete').onclick = (e) => { const s = ppSelected(); if (s.length) deletePendingParts(s, e.target); };
  const pts = (val) => async () => {
    const s = ppSelected();
    if (!s.length) return;
    try { await setNoPoint(s.map((p) => p.code), val); renderPendingParts(); toast(`ตั้ง${val ? 'ไม่นับแต้ม' : 'นับแต้ม'} ${s.length} รายการแล้ว`); }
    catch (err) { toast(err instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : err.message); }
  };
  $('#btn-pp-nopoint').onclick = pts(true);
  $('#btn-pp-point').onclick = pts(false);
}

function fillPartForm(p) {
  if (!p) return;
  const f = $('#part-form');
  f.code.value = p.code; f.name.value = p.name; f.model.value = p.model; f.price.value = p.price;
  f.cost.value = p.cost === '' || p.cost == null ? '' : p.cost;
  f.max_disc.value = p.maxDisc === '' || p.maxDisc == null ? '' : p.maxDisc;
  f.no_point.checked = !!p.noPoint;
  f.scrollIntoView({ behavior: 'smooth' }); f.name.focus();
}

async function loadUsers() {
  try {
    const rows = await api('users');
    S._users = rows;
    $('#users-table tbody').innerHTML = rows.map((u, i) => `<tr><td>${esc(u.username)}</td><td>${esc(u.name)}</td><td>${u.role === 'admin' ? 'แอดมิน' : 'พนักงาน'}</td>
      <td>${esc(branchInfo(u.branch).name)}</td><td>${u.active ? 'ใช้งาน' : '<span class="badge void">ปิด</span>'}</td><td><button data-user="${i}">แก้ไข</button></td></tr>`).join('');
  } catch (e) { toast(e.message); }
}

async function saveUserForm(ev) {
  ev.preventDefault();
  const f = Object.fromEntries(new FormData(ev.target));
  const u = { username: f.username.trim(), name: f.name.trim(), role: f.role, branch: f.branch, pin: f.pin, approve: f.approve, active: !!f.active };
  try {
    await api('saveUser', { user: u });
    toast('บันทึกผู้ใช้แล้ว');
    ev.target.reset();
    loadUsers();
  } catch (e) { toast(e.message); }
}

async function savePartForm(ev) {
  ev.preventDefault();
  const f = Object.fromEntries(new FormData(ev.target));
  if (f.max_disc !== '' && (num(f.max_disc) < 0 || num(f.max_disc) > 100)) return toast('ลดได้สูงสุดต้องเป็น 0 ถึง 100%');
  f.approve = true;   // แอดมินบันทึกผ่านฟอร์ม = ตรวจแล้ว
  f.no_point = !!f.no_point;
  try {
    const r = await api('saveParts', { parts: [f] });
    toast(r.added ? 'เพิ่มอะไหล่แล้ว' : 'แก้ไขอะไหล่แล้ว');
    ev.target.reset();
    await syncParts(); renderPendingParts();
  } catch (e) { toast(e.message); }
}

async function importPartsFile(file) {
  const msg = $('#parts-import-msg');
  if (!window.XLSX) return toast('ตัวอ่าน Excel ยังโหลดไม่เสร็จ');
  msg.textContent = 'กำลังอ่านไฟล์…';
  const wb = XLSX.read(await file.arrayBuffer());
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  const hi = aoa.findIndex((r) => r.some((c) => String(c).trim() === 'รหัสอะไหล่'));
  if (hi < 0) { msg.textContent = 'ไม่พบหัวคอลัมน์ "รหัสอะไหล่"'; return; }
  const head = aoa[hi].map((c) => String(c).trim());
  const col = (names) => head.findIndex((h) => names.includes(h));
  const ix = { code: col(['รหัสอะไหล่']), name: col(['ชื่ออะไหล่']), model: col(['รุ่น']), cost: col(['ทุนไม่รวมภาษี', 'ทุนไม่รวมภาษี/ชิ้น', 'ทุน']),
    price_ex: col(['ราคาขายไม่รวมภาษี', 'ขายปลีกไม่รวมภาษี']), price: col(['ราคาขายรวมภาษี', 'ราคาขาย']) };
  if (ix.price < 0) { msg.textContent = 'ไม่พบคอลัมน์ "ราคาขายรวมภาษี"'; return; }
  const items = aoa.slice(hi + 1).filter((r) => String(r[ix.code]).trim()).map((r) => ({
    code: String(r[ix.code]).trim(), name: ix.name >= 0 ? r[ix.name] : '', model: ix.model >= 0 ? r[ix.model] : '',
    cost: ix.cost >= 0 ? r[ix.cost] : '', price_ex: ix.price_ex >= 0 ? r[ix.price_ex] : '', price: r[ix.price] }));
  if (!confirm(`นำเข้า ${items.length.toLocaleString()} รายการ?`)) { msg.textContent = ''; return; }
  let added = 0, updated = 0;
  try {
    for (let i = 0; i < items.length; i += 2000) {
      msg.textContent = `กำลังส่ง ${i.toLocaleString()} / ${items.length.toLocaleString()}…`;
      const r = await api('saveParts', { parts: items.slice(i, i + 2000) });
      added += r.added; updated += r.updated;
    }
    msg.textContent = `เสร็จแล้ว: เพิ่ม ${added.toLocaleString()} แก้ไข ${updated.toLocaleString()} รายการ`;
    await syncParts();
  } catch (e) { msg.textContent = 'ผิดพลาด: ' + e.message; }
}

/* ================================================================ ใบสั่งซื้ออะไหล่ (สั่งจากร้านอื่น) */
// รายการในใบสั่งซื้อพิมพ์เองทั้งหมด แยกจากรายการอะไหล่ที่ขาย ระบบจำชื่อที่เคยพิมพ์ไว้ให้เลือก
// ต้องต่อเน็ตตอนบันทึก เพราะเลขที่ใบสั่งซื้อออกที่ระบบกลาง (เลขในสาขาเรียงต่อกันเสมอ)

const PO_CREDIT = 'เครดิต', PO_CASH = 'ชำระเลย';
const PO_MIN_ROWS = 5;
const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
const TH_MON = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
function thDateLong(iso) { if (!iso) return ''; const [y, m, d] = String(iso).slice(0, 10).split('-'); return `${num(d)} ${TH_MON[num(m) - 1]} ${num(y) + 543}`; }
function monthName(ym) { const [y, m] = ym.split('-'); return `${TH_MONTHS[num(m) - 1]} ${num(y) + 543}`; }
function monthRange(ym) { const [y, m] = ym.split('-').map(Number); return { from: `${ym}-01`, to: `${ym}-${pad(new Date(y, m, 0).getDate(), 2)}` }; }
const poTermText = (t) => (t === PO_CASH ? 'ชำระเลย' : 'เครดิต (รวมจ่ายสิ้นเดือน)');

async function loadSuppliers(branch) {
  try { S.suppliers = await api('suppliers', { branch }); S.suppliersBranch = branch; } catch (e) { S.suppliers = S.suppliers || []; toast(e.message); }
  $('#po-sup-list').innerHTML = S.suppliers.filter((s) => s.branch === branch).map((s) => `<option value="${esc(s.name)}">`).join('');
}
function supplierByName(name) { return (S.suppliers || []).find((s) => s.branch === S.branch && s.name === String(name).trim()); }

async function openPOTab() {
  if (!$('#po-month').value) $('#po-month').value = todayISO().slice(0, 7);
  if (!$('#po-branch').options.length) fillBranchSelects();
  if (S.suppliersBranch !== S.branch) await loadSuppliers(S.branch);
  if (!S.po || S.po.branch !== S.branch) newPO();
  if (!S.poList || S.poList.branch !== (isAdmin() ? $('#po-branch').value || S.branch : S.user.branch)) loadPOs();
}

function newPO() {
  S.po = { branch: S.branch, po_no: '', lines: [] };
  $('#po-no').textContent = 'ใบใหม่';
  $('#po-date').value = todayISO();
  $('#po-supplier').value = ''; $('#po-phone').value = ''; $('#po-note').value = '';
  $(`#po-term [value="${PO_CREDIT}"]`).checked = true;
  renderPOLines();
}

function editPO(p) {
  S.po = { branch: p.branch, po_no: p.po_no, lines: p.lines.map((l) => ({ name: l.name, qty: l.qty, unit: l.unit })) };
  $('#po-no').textContent = p.po_no + ' (แก้ไข)';
  $('#po-date').value = p.date;
  $('#po-supplier').value = p.supplier_name; $('#po-phone').value = p.phone || ''; $('#po-note').value = p.note || '';
  $(`#po-term [value="${p.term === PO_CASH ? PO_CASH : PO_CREDIT}"]`).checked = true;
  renderPOLines();
  $('#tab-po').scrollIntoView({ behavior: 'smooth' });
}

function poLineAmount(l) { return l.unit === '' || l.unit == null || isNaN(Number(l.unit)) ? null : round2(num(l.qty) * num(l.unit)); }
function renderPOLines() {
  const L = S.po.lines;
  while (L.length < PO_MIN_ROWS) L.push({ name: '', qty: '', unit: '' });
  $('#po-lines tbody').innerHTML = L.map((l, i) => {
    const a = poLineAmount(l);
    return `<tr data-i="${i}"><td class="c">${i + 1}</td>
      <td><input data-k="name" list="po-item-list" value="${esc(l.name)}" placeholder="พิมพ์ชื่ออะไหล่ / รายละเอียด"></td>
      <td><input data-k="qty" type="number" min="0" step="any" class="num" value="${esc(l.qty)}"></td>
      <td><input data-k="unit" type="number" min="0" step="any" class="num" value="${esc(l.unit)}"></td>
      <td class="num po-amt">${a == null || !l.name ? '' : money(a)}</td>
      <td><button type="button" class="link" data-del="${i}" title="ลบบรรทัด">✕</button></td></tr>`;
  }).join('');
  poTotal();
}
function poTotal() {
  const t = S.po.lines.filter((l) => l.name).reduce((s, l) => s + (poLineAmount(l) || 0), 0);
  $('#po-total').textContent = money(t);
  return round2(t);
}
function onPOLineInput(ev) {
  const tr = ev.target.closest('tr[data-i]'); if (!tr) return;
  const l = S.po.lines[num(tr.dataset.i)], k = ev.target.dataset.k;
  l[k] = ev.target.value;
  if (k === 'name' && l.name && !l.qty) { l.qty = 1; $('[data-k="qty"]', tr).value = 1; }
  const a = poLineAmount(l);
  $('.po-amt', tr).textContent = a == null || !l.name ? '' : money(a);
  poTotal();
}

function poForm() {
  const term = $('#po-term input:checked').value;
  const lines = S.po.lines.filter((l) => String(l.name).trim()).map((l) => ({ name: String(l.name).trim(), qty: num(l.qty) || 1, unit: l.unit === '' || l.unit == null ? '' : num(l.unit) }));
  return { po_no: S.po.po_no, branch: S.po.branch, date: $('#po-date').value, supplier_name: $('#po-supplier').value.trim(), phone: $('#po-phone').value.trim(), term, note: $('#po-note').value.trim(), lines };
}

async function savePO(print) {
  const f = poForm();
  if (!f.supplier_name) return toast('กรุณาใส่ชื่อร้าน');
  if (!f.lines.length) return toast('ยังไม่มีรายการ');
  if (f.lines.some((l) => l.unit === '') && !confirm('บางรายการยังไม่ได้ใส่ราคา บันทึกไปก่อนเลยไหม?')) return;
  const sup = supplierByName(f.supplier_name);
  if (sup) f.supplier_id = sup.id;
  else if (!confirm(`"${f.supplier_name}" ยังไม่มีในรายชื่อร้าน ระบบจะเพิ่มให้ (เงื่อนไข: ${poTermText(f.term)}) ตกลงไหม?`)) return;
  $('#btn-po-save').disabled = $('#btn-po-save-print').disabled = true;
  try {
    const p = await api('savePO', { po: f });
    toast(`บันทึกใบสั่งซื้อ ${p.po_no} แล้ว`);
    if (!sup) await loadSuppliers(S.branch);
    rememberPOItems([p]);
    if (print) printPO(p);
    newPO();
    if (S.poList && S.poList.branch === p.branch) loadPOs();
  } catch (e) {
    toast(e instanceof NetError ? 'บันทึกใบสั่งซื้อต้องต่ออินเทอร์เน็ต: ' + e.message : e.message, 6000);
  } finally {
    $('#btn-po-save').disabled = $('#btn-po-save-print').disabled = false;
  }
}

function rememberPOItems(list) {
  S.poItems = S.poItems || new Set();
  list.forEach((p) => p.lines.forEach((l) => l.name && S.poItems.add(l.name)));
  $('#po-item-list').innerHTML = Array.from(S.poItems).sort().slice(0, 2000).map((n) => `<option value="${esc(n)}">`).join('');
}

async function loadPOs() {
  const branch = isAdmin() ? $('#po-branch').value || S.branch : S.user.branch;
  const ym = $('#po-month').value || todayISO().slice(0, 7);
  const { from, to } = monthRange(ym);
  $('#po-msg').textContent = 'กำลังโหลด…';
  try {
    const [list, unpaid] = await Promise.all([api('pos', { branch, from, to }), api('pos', { branch, unpaid: true, to })]);
    S.poList = { branch, ym, from, to, list, unpaid };
    rememberPOItems(list);
    if (!S.poItemsLoaded) { S.poItemsLoaded = true; api('pos', { branch, from: isoAddDays(from, -180), to }).then(rememberPOItems).catch(() => {}); }
    renderPOList();
  } catch (e) { $('#po-msg').textContent = e.message; }
}
function isoAddDays(iso, n) { const d = new Date(iso + 'T00:00:00'); d.setDate(d.getDate() + n); return `${d.getFullYear()}-${pad(d.getMonth() + 1, 2)}-${pad(d.getDate(), 2)}`; }

function poStatus(p) {
  if (p.status === 'ยกเลิก') return '<span class="muted">ยกเลิก</span>';
  return p.paid_date ? `<span class="ok">ชำระแล้ว ${thDate(p.paid_date)}</span>` : '<b class="warn">ยังไม่ชำระ</b>';
}
function renderPOList() {
  const R = S.poList, unpaidOnly = $('#po-filter').value === 'unpaid';
  const rows = unpaidOnly ? R.unpaid : R.list;
  const live = R.list.filter((p) => p.status !== 'ยกเลิก');
  const sum = (a) => a.reduce((s, p) => s + num(p.total), 0);
  $('#po-msg').textContent = `${branchInfo(R.branch).name} · ${monthName(R.ym)} · ${live.length} ใบ ยอดรวม ${money(sum(live))} บาท · ยังไม่ชำระทั้งหมด ${R.unpaid.length} ใบ ${money(sum(R.unpaid))} บาท`;
  $('#po-table tbody').innerHTML = rows.map((p) => `<tr class="${p.status === 'ยกเลิก' ? 'void' : ''}">
      <td>${esc(p.po_no)}</td><td>${thDate(p.date)}</td><td>${esc(p.supplier_name)}</td><td>${p.term === PO_CASH ? 'ชำระเลย' : 'เครดิต'}</td>
      <td class="num">${money(p.total)}</td><td>${poStatus(p)}</td><td>${esc(p.created_by || '')}</td>
      <td class="nowrap"><button type="button" data-po-print="${esc(p.po_no)}">พิมพ์</button>
        ${p.status === 'ยกเลิก' ? '' : `${p.paid_date && p.term === PO_CREDIT ? '' : `<button type="button" data-po-edit="${esc(p.po_no)}">แก้ไข</button>`}
        <button type="button" data-po-pay="${esc(p.po_no)}">${p.paid_date ? 'ยกเลิกจ่าย' : 'จ่ายแล้ว'}</button>
        <button type="button" data-po-void="${esc(p.po_no)}">ยกเลิก</button>`}</td></tr>`).join('')
    || `<tr><td colspan="8" class="muted">${unpaidOnly ? 'ไม่มีใบสั่งซื้อที่ค้างชำระ' : 'ยังไม่มีใบสั่งซื้อในเดือนนี้'}</td></tr>`;
}
function findPOInList(no) { const R = S.poList || {}; return (R.list || []).concat(R.unpaid || []).find((p) => p.po_no === no); }

async function payPOs(nos, paid) {
  const date = paid ? prompt('วันที่จ่ายเงิน (ปปปป-ดด-วว)', todayISO()) : '';
  if (paid && !/^\d{4}-\d{2}-\d{2}$/.test(String(date || ''))) return;
  if (!paid && !confirm('ยกเลิกการบันทึกว่าจ่ายเงินแล้ว?')) return;
  try { await api('payPO', { po_nos: nos, date }); toast(paid ? `บันทึกจ่ายแล้ว ${nos.length} ใบ` : 'ยกเลิกการบันทึกจ่ายแล้ว'); await loadPOs(); if (!$('#po-preview').classList.contains('hidden')) showPOSummary(); }
  catch (e) { toast(e.message, 6000); }
}
async function voidPO(p) {
  let approval, reason = '';
  if (!isAdmin()) { const ap = await askApproval('ยกเลิกใบสั่งซื้อ ' + p.po_no, `${esc(p.supplier_name)} ยอด ${money(p.total)} บาท`); if (!ap) return; approval = ap.approval; reason = ap.reason; }
  else { reason = prompt('เหตุผลที่ยกเลิกใบสั่งซื้อ ' + p.po_no); if (reason == null) return; }
  try { await api('voidPO', { po_no: p.po_no, approval, reason }); toast('ยกเลิกใบสั่งซื้อแล้ว'); loadPOs(); } catch (e) { toast(e.message, 6000); }
}

/* ---------------- ใบสั่งซื้อ (พิมพ์ A4 แนวนอน: ต้นฉบับ + สำเนา) */

function poCopyHtml(p, copy) {
  const br = branchInfo(p.branch);
  const lines = p.lines || [];
  const rows = lines.map((l, i) => `<tr><td class="c">${i + 1}</td><td>${esc(l.name)}</td><td class="num">${num(l.qty)}</td>
      <td class="num">${l.unit === '' || l.unit == null ? '' : money(l.unit)}</td><td class="num">${l.amount === '' || l.amount == null ? '' : money(l.amount)}</td></tr>`).join('');
  const box = (on) => `<span class="po-box">${on ? '✓' : ''}</span>`;
  return `<div class="po-copy${p.status === 'ยกเลิก' ? ' po-void' : ''}">
    <div class="po-top">
      <img class="po-logo" src="${LOGO}" alt="">
      <div class="po-co"><div class="po-co-name">${esc(S.settings.shop_name || 'หจก.ประสบชัยกลการ')}</div>
        <div class="po-co-sub">สาขา${esc(br.name)}${br.address ? ' · ' + esc(br.address) : ''}${br.phone ? ' · โทร ' + esc(br.phone) : ''}</div></div>
      <div class="po-tag">${copy}</div>
    </div>
    <div class="po-titlebar"><div class="po-title">ใบสั่งซื้ออะไหล่</div>
      <div class="po-meta"><div><span>เลขที่</span><b>${esc(p.po_no)}</b></div><div><span>วันที่</span><b>${thDateLong(p.date)}</b></div></div></div>
    <div class="po-info">
      <div><span>ร้าน</span><b>${esc(p.supplier_name)}</b></div>
      <div><span>โทร</span>${esc(p.phone || '-')}</div>
      <div class="po-terms">${box(p.term !== PO_CASH)} เครดิต (รวมจ่ายสิ้นเดือน) &nbsp; ${box(p.term === PO_CASH)} ชำระเลย</div>
    </div>
    <table class="po-table"><thead><tr><th style="width:8%">ลำดับ</th><th>รายการ</th><th style="width:11%" class="num">จำนวน</th><th style="width:17%" class="num">ราคาต่อหน่วย</th><th style="width:18%" class="num">จำนวนเงิน</th></tr></thead>
      <tbody>${rows}</tbody></table>
    <div class="po-sum"><div class="po-words">${num(p.total) ? '(' + esc(bahtText(p.total)) + ')' : ''}</div><div class="po-total"><span>รวมเป็นเงิน</span><b>${money(p.total)}</b></div></div>
    ${p.note ? `<div class="po-note">หมายเหตุ: ${esc(p.note)}</div>` : ''}
    <div class="po-sign">
      <div><div class="po-line"></div>ผู้สั่งซื้อ${p.created_by ? ' (' + esc(p.created_by) + ')' : ''}</div>
      <div><div class="po-line"></div>ผู้อนุมัติ</div>
      <div><div class="po-line"></div>ผู้รับของ</div>
    </div>
    ${p.status === 'ยกเลิก' ? '<div class="po-stamp">ยกเลิก</div>' : ''}
  </div>`;
}
function poDocHtml(p) { return `<div class="po-sheet">${poCopyHtml(p, 'ต้นฉบับ')}${poCopyHtml(p, 'สำเนา')}</div>`; }

// พิมพ์ หรือเลือกเครื่องพิมพ์ "บันทึกเป็น PDF" ชื่อไฟล์จะตามชื่อเอกสาร
function printDoc(html, title, tight) {
  $('#print-area').innerHTML = html;
  $('#print-area').className = 'print-only' + (tight ? ' pa-tight' : '');
  const old = document.title;
  document.title = title;
  setTimeout(() => { window.print(); setTimeout(() => { document.title = old; }, 500); }, 80);
}
function printPO(p) { setPrintPage('@page { size: A4 landscape; margin: 0; }'); printDoc(poDocHtml(p), `ใบสั่งซื้อ ${p.po_no} ${p.supplier_name}`, true); }

/* ---------------- ใบสรุปเบิกเงินสิ้นเดือน */

// ใบสั่งซื้อที่ต้องเบิกเงิน: แบบเครดิตของเดือนนี้ + ทุกใบที่ยังไม่ชำระ (รวมเดือนก่อน) จัดกลุ่มตามร้าน
function poSummaryData() {
  const R = S.poList;
  const map = new Map();
  R.list.filter((p) => p.status !== 'ยกเลิก' && p.term !== PO_CASH).concat(R.unpaid).forEach((p) => map.set(p.po_no, p));
  const groups = new Map();
  Array.from(map.values()).sort((a, b) => (a.supplier_name + a.po_no < b.supplier_name + b.po_no ? -1 : 1)).forEach((p) => {
    if (!groups.has(p.supplier_name)) groups.set(p.supplier_name, []);
    groups.get(p.supplier_name).push(p);
  });
  const cash = R.list.filter((p) => p.status !== 'ยกเลิก' && p.term === PO_CASH);
  return { groups, cash };
}
function poItemsText(p) { const n = p.lines.map((l) => l.name + (num(l.qty) !== 1 ? ' x' + num(l.qty) : '')); return n.slice(0, 3).join(', ') + (n.length > 3 ? ` และอีก ${n.length - 3} รายการ` : ''); }

function poSummaryHtml(forPrint) {
  const R = S.poList, { groups, cash } = poSummaryData(), br = branchInfo(R.branch);
  let all = 0, unpaid = 0, i = 0;
  const body = Array.from(groups.entries()).map(([name, list]) => {
    const t = list.reduce((s, p) => s + num(p.total), 0), u = list.filter((p) => !p.paid_date);
    all += t; unpaid += u.reduce((s, p) => s + num(p.total), 0);
    return `<tr class="cs-dev"><td colspan="6">${esc(name)} · ${list.length} ใบ${!forPrint && u.length ? ` <button type="button" class="po-noprint" data-po-paygroup="${esc(u.map((p) => p.po_no).join(','))}">บันทึกจ่ายร้านนี้ (${u.length} ใบ)</button>` : ''}</td></tr>
      ${list.map((p) => `<tr><td class="c">${++i}</td><td>${esc(p.po_no)}</td><td>${thDate(p.date)}</td><td>${esc(poItemsText(p))}</td>
        <td class="num">${money(p.total)}</td><td>${p.paid_date ? 'ชำระแล้ว ' + thDate(p.paid_date) : (p.date < R.from ? 'ค้างจากเดือนก่อน' : 'รอจ่าย')}</td></tr>`).join('')}
      <tr class="cs-sub"><td colspan="4">รวม ${esc(name)}</td><td class="num">${money(t)}</td><td></td></tr>`;
  }).join('');
  const cashSum = cash.reduce((s, p) => s + num(p.total), 0);
  return `<div class="cs-doc po-sum-doc">
    <div class="cs-head"><div><img class="po-logo-s" src="${LOGO}" alt=""><span class="cs-title">ใบสรุปยอดใบสั่งซื้อเพื่อเบิกเงิน</span><div class="cs-muted">${esc(S.settings.shop_name || 'หจก.ประสบชัยกลการ')} · สาขา${esc(br.name)}</div></div>
      <div class="cs-date">ประจำเดือน <b>${monthName(R.ym)}</b></div></div>
    ${body ? `<table class="cs-table"><thead><tr><th style="width:6%">#</th><th style="width:14%">เลขที่</th><th style="width:12%">วันที่</th><th>รายการ</th><th style="width:14%" class="num">ยอดเงิน</th><th style="width:17%">สถานะ</th></tr></thead>
      <tbody>${body}</tbody>
      <tfoot><tr class="cs-total"><td colspan="4">รวมทุกร้าน</td><td class="num">${money(all)}</td><td></td></tr></tfoot></table>` : '<p class="muted">ไม่มีใบสั่งซื้อแบบเครดิตหรือที่ค้างชำระ</p>'}
    <div class="cs-sum po-sum-box">
      <div class="big"><span>ยอดที่ขอเบิกเงิน (ยังไม่ชำระ)</span><b>${money(unpaid)}</b></div>
      <div><span>ชำระแล้วในรายการข้างบน</span><b>${money(all - unpaid)}</b></div>
      <div><span>ใบสั่งซื้อชำระเลยเดือนนี้ (${cash.length} ใบ)</span><b>${money(cashSum)}</b></div>
    </div>
    <div class="cs-count">(${esc(bahtText(unpaid))})</div>
    <div class="cs-sign"><div>ลงชื่อ ...................................<br>ผู้จัดทำ</div><div>ลงชื่อ ...................................<br>ผู้อนุมัติ</div><div>ลงชื่อ ...................................<br>ฝ่ายการเงิน (ผู้จ่ายเงิน)</div></div>
  </div>`;
}
async function showPOSummary() {
  if (!S.poList) await loadPOs();
  if (!S.poList) return;
  $('#po-preview').classList.remove('hidden');
  $('#po-preview').innerHTML = `<div class="row po-noprint" style="margin-bottom:8px"><b class="grow">ใบสรุปเบิกเงิน ${monthName(S.poList.ym)}</b>
    <button type="button" id="btn-po-sum-print" class="primary">พิมพ์ / บันทึก PDF</button><button type="button" id="btn-po-sum-close">ปิด</button></div>` + poSummaryHtml(false);
  $('#po-preview').scrollIntoView({ behavior: 'smooth' });
}
function printPOSummary() {
  const R = S.poList;
  printDoc(poSummaryHtml(true), `สรุปใบสั่งซื้อ ${branchInfo(R.branch).name} ${monthName(R.ym)}`);
}

function exportPOExcel() {
  const R = S.poList;
  if (!R) return toast('กด "แสดง" ก่อน');
  if (!window.XLSX) return toast('ตัวสร้าง Excel ยังโหลดไม่เสร็จ ลองใหม่อีกครั้ง');
  const { groups } = poSummaryData();
  const sum = [['ใบสรุปยอดใบสั่งซื้อเพื่อเบิกเงิน', '', '', '', '', '', ''], [`สาขา${branchInfo(R.branch).name}`, `ประจำเดือน ${monthName(R.ym)}`, '', '', '', '', ''], [],
    ['ร้าน', 'เลขที่', 'วันที่', 'การชำระ', 'รายการ', 'ยอดเงิน', 'สถานะ']];
  let total = 0, unpaid = 0;
  groups.forEach((list, name) => {
    list.forEach((p) => { total += num(p.total); if (!p.paid_date) unpaid += num(p.total); sum.push([name, p.po_no, thDate(p.date), p.term, poItemsText(p), round2(num(p.total)), p.paid_date ? 'ชำระแล้ว ' + thDate(p.paid_date) : 'ยังไม่ชำระ']); });
    sum.push(['', '', '', '', 'รวม ' + name, round2(list.reduce((s, p) => s + num(p.total), 0)), '']);
  });
  sum.push([], ['', '', '', '', 'รวมทุกร้าน', round2(total), ''], ['', '', '', '', 'ยอดที่ขอเบิก (ยังไม่ชำระ)', round2(unpaid), '']);
  const det = [['เลขที่', 'วันที่', 'ร้าน', 'โทร', 'การชำระ', 'ลำดับ', 'รายการ', 'จำนวน', 'ราคาต่อหน่วย', 'จำนวนเงิน', 'ยอดทั้งใบ', 'สถานะ', 'วันที่ชำระ', 'ผู้ออกใบ', 'หมายเหตุ']];
  R.list.forEach((p) => p.lines.forEach((l, i) => det.push([p.po_no, thDate(p.date), p.supplier_name, p.phone || '', p.term, i + 1, l.name, num(l.qty), l.unit === '' ? '' : round2(num(l.unit)),
    l.amount === '' ? '' : round2(num(l.amount)), i === 0 ? round2(num(p.total)) : '', p.status === 'ยกเลิก' ? 'ยกเลิก' : p.paid_date ? 'ชำระแล้ว' : 'ยังไม่ชำระ', thDate(p.paid_date), p.created_by || '', p.note || ''])));
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet(sum); ws1['!cols'] = [22, 12, 11, 10, 46, 12, 18].map((w) => ({ wch: w }));
  const ws2 = XLSX.utils.aoa_to_sheet(det); ws2['!cols'] = [12, 11, 20, 13, 9, 6, 40, 7, 12, 12, 12, 11, 11, 10, 20].map((w) => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, ws1, 'สรุปเบิกเงิน');
  XLSX.utils.book_append_sheet(wb, ws2, 'ใบสั่งซื้อทั้งเดือน');
  XLSX.writeFile(wb, `ใบสั่งซื้อ_${branchInfo(R.branch).name}_${R.ym}.xlsx`);
}

/* ---------------- รายชื่อร้าน */

function showSuppliers(focusNew) {
  const list = (S.suppliers || []).filter((s) => s.branch === S.branch).sort((a, b) => (a.id < b.id ? -1 : 1));
  const box = openModal(`<h2>รายชื่อร้านที่สั่งซื้อ · สาขา${esc(branchInfo(S.branch).name)}</h2>
    <table class="grid small"><thead><tr><th>ชื่อร้าน</th><th>เบอร์โทร</th><th>ให้เครดิต</th><th></th></tr></thead><tbody>
    ${list.map((s) => `<tr data-id="${esc(s.id)}"><td><input name="name" class="sup-in" value="${esc(s.name)}"></td><td><input name="phone" class="sup-in" value="${esc(s.phone || '')}"></td>
      <td class="c"><input type="checkbox" name="credit" ${s.credit ? 'checked' : ''}></td><td><button type="button" data-sup-save>บันทึก</button></td></tr>`).join('')}
    <tr data-id="" class="sup-new"><td><input name="name" class="sup-in" placeholder="พิมพ์ชื่อร้านใหม่ตรงนี้"></td><td><input name="phone" class="sup-in" placeholder="เบอร์โทร"></td><td class="c"><input type="checkbox" name="credit" checked></td><td><button type="button" class="primary" data-sup-save>เพิ่ม</button></td></tr>
    </tbody></table>
    <p class="muted">ร้านที่ให้เครดิต: เวลาเลือกร้านในใบสั่งซื้อ ระบบจะตั้งเป็น "เครดิต (รวมจ่ายสิ้นเดือน)" ให้เอง</p>
    <div class="actions"><button type="button" id="sup-close">ปิด</button></div>`);
  $('#sup-close', box).onclick = closeModal;
  if (focusNew === true) setTimeout(() => $('tr[data-id=""] [name=name]', box).focus(), 50);
  box.onclick = async (ev) => {
    const b = ev.target.closest('[data-sup-save]'); if (!b) return;
    const tr = b.closest('tr');
    const s = { id: tr.dataset.id, branch: S.branch, name: $('[name=name]', tr).value.trim(), phone: $('[name=phone]', tr).value.trim(), credit: $('[name=credit]', tr).checked };
    if (!s.name) return toast('กรุณาใส่ชื่อร้าน');
    try { await api('saveSupplier', { supplier: s }); toast('บันทึกแล้ว'); await loadSuppliers(S.branch); if (!s.id && focusNew === true) { closeModal(); $('#po-supplier').value = s.name; $('#po-supplier').dispatchEvent(new Event('change')); return; } showSuppliers(focusNew); } catch (e) { toast(e.message, 6000); }
  };
}

function bindPO() {
  $('#po-lines tbody').addEventListener('input', onPOLineInput);
  $('#po-lines tbody').addEventListener('click', (ev) => {
    const d = ev.target.closest('[data-del]'); if (!d) return;
    S.po.lines.splice(num(d.dataset.del), 1); renderPOLines();
  });
  $('#btn-po-line').onclick = () => { S.po.lines.push({ name: '', qty: '', unit: '' }); renderPOLines(); $$('#po-lines [data-k="name"]').pop().focus(); };
  $('#po-supplier').addEventListener('change', () => {
    const s = supplierByName($('#po-supplier').value); if (!s) return;
    if (s.phone) $('#po-phone').value = s.phone;
    $(`#po-term [value="${s.credit ? PO_CREDIT : PO_CASH}"]`).checked = true;
  });
  $('#btn-po-new').onclick = () => { if (!S.po.lines.some((l) => l.name) || confirm('ล้างหน้าจอใบสั่งซื้อนี้?')) newPO(); };
  $('#btn-po-save').onclick = () => savePO(false);
  $('#btn-po-save-print').onclick = () => savePO(true);
  $('#btn-po-load').onclick = loadPOs;
  $('#po-month').onchange = loadPOs;
  $('#po-branch').onchange = loadPOs;
  $('#po-filter').onchange = () => S.poList && renderPOList();
  $('#btn-po-summary').onclick = showPOSummary;
  $('#btn-po-excel').onclick = exportPOExcel;
  $('#btn-po-suppliers').onclick = showSuppliers;
  $('#btn-po-addsup').onclick = () => showSuppliers(true);
  $('#po-table').onclick = (ev) => {
    const b = ev.target.closest('button'); if (!b) return;
    const p = findPOInList(b.dataset.poPrint || b.dataset.poEdit || b.dataset.poPay || b.dataset.poVoid); if (!p) return;
    if (b.dataset.poPrint) printPO(p);
    else if (b.dataset.poEdit) editPO(p);
    else if (b.dataset.poPay) payPOs([p.po_no], !p.paid_date);
    else if (b.dataset.poVoid) voidPO(p);
  };
  $('#po-preview').onclick = (ev) => {
    if (ev.target.id === 'btn-po-sum-print') printPOSummary();
    else if (ev.target.id === 'btn-po-sum-close') { $('#po-preview').classList.add('hidden'); $('#po-preview').innerHTML = ''; }
    else if (ev.target.dataset.poPaygroup) payPOs(ev.target.dataset.poPaygroup.split(','), true);
  };
}

/* ================================================================ ผูกปุ่ม */

function switchTab(name) {
  $$('.tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  $$('.tab').forEach((t) => t.classList.toggle('hidden', t.id !== 'tab-' + name));
  if (name === 'sale') $('#scan-input').focus();
  if (name === 'report' && !S.report) { $('#rep-from').value = $('#rep-from').value || todayISO(); $('#rep-to').value = $('#rep-to').value || todayISO(); loadReport(); }
  if (name === 'print' && S.printBill) renderPrintPreview(S.printBill);
  if (name === 'admin') loadAudit();
  if (name === 'control') { if (!$('#ctl-date').value) $('#ctl-date').value = todayISO(); loadControl(); }
  if (name === 'po') openPOTab();
  if (name === 'points' && $('#pts-cus').options.length < 2) loadPointMembers();
  if (name === 'estimate') { if (!S.est || S.est.branch !== S.branch) newEstimate(); $('#est-scan').focus(); }
}

function bindStatic() {
  $('#setup-save').onclick = async () => {
    const url = $('#setup-url').value.trim();
    if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) { $('#setup-err').textContent = 'ลิงก์ต้องขึ้นต้น https://script.google.com/ และลงท้าย /exec'; return; }
    S.apiUrl = url; await DB.kv.set('apiUrl', url); showLogin();
  };
  $('#login-reset-url').onclick = (e) => { e.preventDefault(); $('#setup-url').value = S.apiUrl; show('scr-setup'); };
  $('#login-form').onsubmit = async (e) => {
    e.preventDefault();
    $('#login-err').textContent = '';
    try {
      const d = await doLogin($('#login-user').value, $('#login-pin').value);
      $('#login-pin').value = '';
      const sess = { token: d.token, user: d.user, branches: d.branches, settings: d.settings, approvers: d.approvers };
      await DB.kv.set('session', sess);
      await startSession(sess);
    } catch (err) { $('#login-err').textContent = err.message; }
  };
  $('#btn-logout').onclick = async () => {
    if (S.pending && !confirm(`ยังมี ${S.pending} รายการรอส่ง (ข้อมูลไม่หาย จะส่งเมื่อมีคนเข้าระบบในเครื่องนี้) ออกจากระบบ?`)) return;
    await DB.kv.set('session', null);
    clearInterval(S._timer);
    S.token = ''; S.user = null; S.report = null;
    showLogin();
  };
  $('#sync-chip').onclick = showQueue;
  $('#branch-select').onchange = async (e) => {
    S.branch = e.target.value;
    await DB.kv.set('branch:' + S.user.username, S.branch);
    await loadCustomers(); syncCustomers();
    newSale();
    if (!$('#tab-estimate').classList.contains('hidden')) newEstimate();
  };
  $$('.tabs button').forEach((b) => b.onclick = () => switchTab(b.dataset.tab));
  $$('.subtabs button').forEach((b) => b.onclick = () => {
    $$('.subtabs button').forEach((x) => x.classList.toggle('active', x === b));
    $$('.sub').forEach((s) => s.classList.toggle('hidden', s.id !== 'sub-' + b.dataset.sub));
    if (b.dataset.sub === 'users') loadUsers();
    if (b.dataset.sub === 'discounts') loadDiscountAdmin();
    if (b.dataset.sub === 'parts') renderPendingParts();
    if (b.dataset.sub === 'nopoint') renderNoPoint();
    if (b.dataset.sub === 'audit') loadAudit();
  });

  // หน้าขาย
  $('#scan-input').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); scan(e.target.value); e.target.value = ''; } };
  $('#btn-search-part').onclick = () => openPartSearch('', (p) => addToCart(p));
  $('#sale-table').onchange = (e) => {
    const i = e.target.dataset.i;
    if (i == null) return;
    S.cart.lines[i][e.target.dataset.f] = num(e.target.value);
    setTimeout(renderCart);   // รอให้ช่องที่กำลังแก้หลุดโฟกัสก่อนวาดตารางใหม่
  };
  $('#sale-table').onclick = (e) => { const d = e.target.dataset.del; if (d != null) { S.cart.lines.splice(num(d), 1); renderCart(); } };
  $('#sale-discount').oninput = () => { S.cart.discBy = 'baht'; renderCart(); };
  $('#sale-discount-pct').oninput = () => { S.cart.discBy = 'pct'; renderCart(); };
  $('#sale-customer').onchange = (e) => { const c = customerByName(e.target.value); if (c) e.target.value = customerLabel(c); };
  $('#sale-customer').onfocus = (e) => e.target.select();
  $('#btn-clear').onclick = () => { if (!S.cart.lines.length || confirm('ล้างรายการบนหน้าจอ?')) newSale(); };
  $('#btn-save').onclick = () => saveSale(false);
  $('#btn-save-print').onclick = () => saveSale(true);
  document.addEventListener('keydown', (e) => {
    if ($('#scr-app').classList.contains('hidden')) return;
    if (e.key === 'F9') { e.preventDefault(); if (!$('#tab-sale').classList.contains('hidden') && $('#modal').classList.contains('hidden')) saveSale(true); }
    if (e.key === 'F2') { e.preventDefault(); if ($('#modal').classList.contains('hidden')) openPartSearch('', (p) => addToCart(p)); }
    if (e.key === 'Escape' && !$('#modal').classList.contains('hidden') && !$('#ap-form')) closeModal();
  });
  $('#modal').addEventListener('click', (e) => { if (e.target.matches('[data-close]')) closeModal(); });

  // พิมพ์ใบเสร็จ
  const loadPrint = async () => {
    $('#print-msg').textContent = '';
    try { const b = await loadBill($('#print-no').value); S.printBill = b; renderPrintPreview(b); }
    catch (e) { $('#print-msg').textContent = e.message; }
  };
  $('#btn-print-load').onclick = loadPrint;
  $('#print-no').onkeydown = (e) => { if (e.key === 'Enter') loadPrint(); };
  $('#btn-print-go').onclick = () => { if (S.printBill) printBill(S.printBill); else toast('ยังไม่ได้เปิดบิล'); };
  $('#print-preview').onclick = () => { if (S.printBill) openBill(S.printBill.bill_no); };

  // รายงาน
  $('#btn-rep-load').onclick = loadReport;
  $('#sale-pay').onchange = () => { const p = payPicker($('#sale-pay')); if (p.pay_method === 'โอนเงิน') $('#sale-payref').focus(); };
  $('#btn-rep-control').onclick = () => {
    $('#ctl-date').value = $('#rep-from').value || todayISO();
    if (isAdmin() && $('#rep-branch').value !== '*') $('#ctl-branch').value = $('#rep-branch').value;
    switchTab('control'); loadControl();
  };
  $('#btn-ctl-load').onclick = loadControl;
  $('#ctl-preview').onclick = (e) => { const tr = e.target.closest('[data-bill]'); if (tr) openBill(tr.dataset.bill); };
  $('#btn-ctl-print').onclick = printControlSheet;
  bindPO();
  $('#sale-dev').onclick = (ev) => { ev.preventDefault(); changeDevice(); };
  $('#ctl-date').onchange = loadControl;
  $('#ctl-branch').onchange = loadControl;
  $('#btn-rep-excel').onclick = exportExcel;
  bindPoints();
  bindNoPoint();
  $('#rep-table').onclick = (e) => { const tr = e.target.closest('[data-bill]'); if (tr) openBill(tr.dataset.bill); };

  // ลูกค้า
  $('#cus-search').oninput = renderCustomers;
  $('#btn-cus-new').onclick = () => customerForm();
  $('#cus-table').onclick = (e) => { const id = e.target.dataset.cus; if (id) customerForm(S.customers.find((c) => c.id === id)); };

  // แอดมิน
  $('#btn-audit-load').onclick = loadAudit;
  $('#audit-table').onclick = (e) => { const tr = e.target.closest('[data-bill]'); if (tr) openBill(tr.dataset.bill); };
  $('#user-form').onsubmit = saveUserForm;
  $('#branch-disc-form').onsubmit = saveBranchDiscounts;
  // ใบประเมินราคา
  $('#est-scan').onkeydown = (e) => { if (e.key === 'Enter') { e.preventDefault(); estimateScan(e.target.value); e.target.value = ''; } };
  $('#est-search').onclick = () => openPartSearch('', (p) => addEstimateLine(p), { quickAdd: false });
  $('#est-free').onclick = estimateFreeLine;
  $('#est-table').onchange = (e) => { const i = e.target.dataset.i; if (i == null) return; S.est.lines[i][e.target.dataset.f] = num(e.target.value); setTimeout(renderEstimate); };
  $('#est-table').onclick = (e) => { const d = e.target.dataset.del; if (d != null) { S.est.lines.splice(num(d), 1); renderEstimate(); } };
  $('#est-customer').onchange = (e) => {
    const c = customerByName(e.target.value); if (!c) return;
    e.target.value = customerLabel(c);
    if (!$('#est-phone').value && c.phone) $('#est-phone').value = c.phone;
    if (!$('#est-address').value && c.address) $('#est-address').value = c.address;
  };
  $('#est-new').onclick = () => { if (S.est && !S.est.est_no && S.est.lines.length && !confirm('ใบนี้ยังไม่ได้บันทึก จะเริ่มใบใหม่?')) return; newEstimate(); };
  $('#est-list-btn').onclick = estimateList;
  $('#est-save').onclick = () => saveEstimate(false);
  $('#est-save-print').onclick = () => (S.est.status === EST_OPEN ? saveEstimate(true) : printEstimate(S.est));
  $('#est-void').onclick = voidEstimate;
  $('#est-to-bill').onclick = estimateToBill;
  $('#est-close').onclick = closeEstimate;
  $('#est-to-ws').onclick = estimateToWholesale;
  $('#est-reopen').onclick = reopenEstimate;
  bindPendingParts();
  $('#special-parts-body').onclick = (e) => {
    const code = e.target.dataset.editPart; if (!code) return;
    $$('.subtabs button').find((b) => b.dataset.sub === 'parts').click(); fillPartForm(findPart(code));
  };
  $('#users-table').onclick = (e) => {
    const i = e.target.dataset.user; if (i == null) return;
    const u = S._users[i], f = $('#user-form');
    f.username.value = u.username; f.name.value = u.name; f.role.value = u.role; f.branch.value = u.branch; f.active.checked = u.active;
    f.pin.value = ''; f.approve.value = ''; f.username.focus();
  };
  $('#part-form').onsubmit = savePartForm;
  $('#part-form').code.onchange = (e) => fillPartForm(findPart(e.target.value));
  $('#parts-file').onchange = (e) => { if (e.target.files[0]) importPartsFile(e.target.files[0]); e.target.value = ''; };
}

boot().catch((e) => { document.body.innerHTML = '<p style="padding:20px;color:#b3261e">เปิดแอปไม่สำเร็จ: ' + esc(e.message) + '</p>'; });
