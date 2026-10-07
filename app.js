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
      const r = indexedDB.open('parts-app', 1);
      r.onupgradeneeded = () => {
        const d = r.result;
        d.createObjectStore('kv');
        d.createObjectStore('parts', { keyPath: 'k' });
        d.createObjectStore('customers', { keyPath: 'id' });
        d.createObjectStore('bills', { keyPath: 'bill_no' });
        d.createObjectStore('queue', { keyPath: 'op_id' });
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
async function api(action, body = {}) {
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

async function boot() {
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
    } catch (e) { if (!(e instanceof NetError)) toast(e.message); }
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
    maxDisc: r[6] == null ? '' : r[6], status: r[7] || '', addedBy: r[8] || '' };
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
      if (d.rows.length) {
        const items = d.rows.map(partFromRow);
        let maxU = since;
        items.forEach((p) => { if (p.u > maxU) maxU = p.u; });
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
      rows.forEach((c) => { if (!pendingIds.has(c.id)) s.put(c); });
    });
    await loadCustomers();
  } catch (e) { if (!(e instanceof NetError)) toast(e.message); }
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
function customerFull(c) { return c ? [c.title, c.first, c.last].filter(Boolean).join(' ').trim() : ''; }

function renderCustomers() {
  const q = ($('#cus-search').value || '').toLowerCase();
  const rows = S.customers.filter((c) => !q || JSON.stringify(c).toLowerCase().includes(q)).slice(0, 300);
  $('#cus-table tbody').innerHTML = rows.map((c) => `<tr><td>${esc(customerCode(c))}</td>
    <td>${esc(c.display)} ${c._pending ? '<span class="badge pending">รอส่ง</span>' : ''}</td><td>${esc(customerFull(c))}</td>
    <td>${esc(c.phone)}</td><td>${esc(c.address)}</td><td><button data-cus="${esc(c.id)}">แก้ไข</button></td></tr>`).join('');
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
function numberPrefix(branch) {
  const pre = S.settings && 'bill_prefix' in S.settings ? String(S.settings.bill_prefix).trim().toUpperCase() : 'P';
  return pre + branchInfo(branch).letter;
}
async function numberBlocks(branch) {
  const head = numberPrefix(branch);
  return ((await DB.kv.get('numbers:' + branch)) || []).filter((b) => b.prefix.indexOf(head) === 0 && /^\d{2}$/.test(b.prefix.slice(head.length)));
}
function remaining(blocks) { return blocks.reduce((s, b) => s + (b.end - b.next + 1), 0); }

async function ensureNumbers(branch) {
  const blocks = await numberBlocks(branch);
  if (remaining(blocks) >= NUMBER_LOW) return;
  try {
    const r = await api('reserve', { branch, count: NUMBER_BLOCK });
    blocks.push({ prefix: r.prefix, next: r.start, end: r.end });
    await DB.kv.set('numbers:' + branch, blocks);
  } catch (e) { if (!(e instanceof NetError)) toast('จองเลขบิลไม่สำเร็จ: ' + e.message); }
}

async function peekNumber(branch) {
  const b = (await numberBlocks(branch)).find((x) => x.next <= x.end);
  return b ? b.prefix + pad(b.next, 5) : null;
}

async function takeNumber(branch) {
  let blocks = await numberBlocks(branch);
  if (!blocks.some((x) => x.next <= x.end)) { await ensureNumbers(branch); blocks = await numberBlocks(branch); }
  const b = blocks.find((x) => x.next <= x.end);
  if (!b) throw new Error('เลขที่บิลสำรองในเครื่องหมดแล้ว กรุณาต่ออินเทอร์เน็ตสักครู่');
  const no = b.prefix + pad(b.next, 5);
  b.next++;
  await DB.kv.set('numbers:' + branch, blocks.filter((x) => x.next <= x.end));
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
  if (S.syncing || !S.token) return;
  S.syncing = true; renderSync();
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
    await refreshPending();
  }
}

async function markSynced(op, ok, error) {
  const no = op.bill ? op.bill.bill_no : op.bill_no;
  if (no) {
    const b = await DB.get('bills', no);
    if (b) { b._sync = ok ? 'synced' : 'error'; b._error = error || ''; await DB.put('bills', b); }
  }
  if (op.type === 'saveCustomer') {
    const c = await DB.get('customers', op.customer.id);
    if (c && ok) { delete c._pending; await DB.put('customers', c); await loadCustomers(); }
  }
}

function showQueue() {
  DB.all('queue').then((q) => {
    q.sort((a, b) => a.seq - b.seq);
    const label = { createBill: 'บิลขาย', editBill: 'แก้ไขบิล', voidBill: 'ยกเลิกบิล', returnBill: 'บิลคืน', saveCustomer: 'ลูกค้า', addPart: 'อะไหล่ด่วน' };
    const box = openModal(`<h2>ข้อมูลที่ยังไม่ได้ส่งขึ้นระบบ</h2>
      <p class="muted">${S.online ? 'ออนไลน์' : 'ออฟไลน์อยู่ ระบบจะส่งให้อัตโนมัติเมื่อมีอินเทอร์เน็ต'}</p>
      ${q.length ? `<table class="grid small"><thead><tr><th>รายการ</th><th>บิล/ลูกค้า</th><th>สถานะ</th><th></th></tr></thead><tbody>
      ${q.map((o) => `<tr><td>${label[o.type] || o.type}</td><td>${esc(o.bill ? o.bill.bill_no : o.bill_no || (o.customer && o.customer.display) || (o.part && o.part.code))}</td>
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
  $('#scan-msg').textContent = '';
  renderCart();
  await ensureNumbers(S.branch);
  $('#sale-no').textContent = (await peekNumber(S.branch)) || 'รอจองเลข (ต่อเน็ต)';
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

function openPartSearch(q, onPick) {
  const box = openModal(`<h2>ค้นหาอะไหล่</h2>
    <input id="ps-q" placeholder="พิมพ์ชื่อ รุ่น หรือรหัสบางส่วน" style="width:100%" value="${esc(q || '')}" autofocus>
    <div class="search-results"><table class="grid small"><thead><tr><th>รหัส</th><th>ชื่ออะไหล่</th><th>รุ่น</th><th class="num">ราคา</th></tr></thead><tbody id="ps-body"></tbody></table></div>
    <div class="actions"><button type="button" id="ps-add" title="รหัสไม่มีในระบบ ลูกค้ารอ">➕ เพิ่มอะไหล่ด่วน</button><button data-close>ปิด</button></div>`);
  let found = [];
  $('#ps-add', box).onclick = () => { const v = $('#ps-q', box).value.trim(); closeModal(); quickAddPart(/\s/.test(v) ? '' : v, v, onPick); };
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
  const needAp = priceChanged || overDisc;
  let ap = null;
  if (needAp) {
    const why = [priceChanged ? 'มีการแก้ราคาขาย' : '', overDisc ? `ส่วนลด ${money(discount)} บาท เกินที่ลดได้ ${money(allowed)} บาท` : ''].filter(Boolean).join(' และ ');
    ap = await askApproval('ต้องให้แอดมินอนุมัติ', esc(why), false);
    if (!ap) return;
  }
  let no;
  try { no = await takeNumber(S.branch); } catch (e) { return toast(e.message, 6000); }
  const bill = {
    bill_no: no, branch: S.branch, date: $('#sale-date').value || todayISO(), time: nowTime(), type: 'ขาย', ref_bill: '',
    customer_id: cust ? cust.id : '', customer_name: custName, total, discount, net: round2(total - discount), status: 'ปกติ',
    created_by: S.user.username, created_at: nowISO(), note: '',
    lines: lines.map((l, i) => ({ line: i + 1, code: l.code, name: l.name, model: l.model, qty: num(l.qty), unit: num(l.unit), amount: round2(num(l.qty) * num(l.unit)) }))
  };
  await DB.put('bills', Object.assign({ _sync: 'pending' }, bill));
  await enqueue({ type: 'createBill', bill, approval: ap ? ap.approval : undefined });
  toast('บันทึกบิล ' + no + ' แล้ว');
  S.printBill = bill;
  renderPrintPreview(bill);
  if (print) printBill(bill);
  newSale();
}

/* ================================================================ ใบเสร็จ */

function receiptHtml(bill, copy) {
  const c = S.customers.find((x) => x.id === bill.customer_id);
  const isRet = bill.type === 'คืน';
  const title = isRet ? 'ใบคืนอะไหล่' : (S.settings.receipt_title || 'ใบเบิกอะไหล่');
  const lines = bill.lines || [];
  const rows = Math.max(15, lines.length);
  let body = '';
  for (let i = 0; i < rows; i++) {
    const l = lines[i];
    body += l ? `<tr><td class="num">${l.qty}</td><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="num">${money(l.unit)}</td><td class="num">${money(l.amount)}</td><td>${i === 0 && !isRet ? 'จ่ายเงินวันที่' : ''}</td></tr>`
      : '<tr><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>';
  }
  const br = branchInfo(bill.branch);
  return `<div class="receipt ${lines.length > 15 ? 'long' : ''}">
    <div class="r-copy">${copy}</div>
    <div class="r-title">${esc(title)}</div>
    <div class="small" style="text-align:center">${esc(S.settings.shop_name || '')} สาขา${esc(br.name)}${bill.status === 'ยกเลิก' ? ' — <b>ยกเลิกแล้ว</b>' : ''}</div>
    <div class="r-head">
      <div>ชื่อ <b>${esc(bill.customer_name)}</b></div><div>เลขที่ <b>${esc(bill.bill_no)}</b></div>
      <div>ชื่อลูกค้า ${esc(customerFull(c))} ${c && c.address ? ' ' + esc(c.address) : ''}</div><div>วันที่ ${thDate(bill.date)}</div>
      ${isRet ? `<div>อ้างอิงบิล ${esc(bill.ref_bill)}</div><div></div>` : ''}
    </div>
    <table><thead><tr><th style="width:9%">จำนวน</th><th style="width:17%">รหัสอะไหล่</th><th>ชื่ออะไหล่</th><th style="width:14%">รุ่น</th><th style="width:10%">หน่วยละ</th><th style="width:12%">จำนวนเงิน</th><th style="width:13%">หมายเหตุ</th></tr></thead>
    <tbody>${body}</tbody></table>
    <div class="r-foot">
      <div>(${esc(bahtText(bill.net))})</div><div>รวม</div><div class="num">${money(bill.total)}</div>
      <div>${esc(bill.customer_name)}</div><div>ส่วนลด</div><div class="num">${money(bill.discount)}</div>
      <div>ลงชื่อ.......................................................ผู้รับสินค้า</div><div><b>เป็นเงิน</b></div><div class="num"><b>${money(bill.net)}</b></div>
    </div>
    <div class="r-sign">ลงชื่อ...............................................ผู้จัดอะไหล่ (${esc(bill.created_by || '')})&nbsp;&nbsp;&nbsp;ลงชื่อ..................................ผู้ตรวจสอบ</div>
    ${copy === 'สำหรับลูกค้า' ? `<div class="r-note">${esc(S.settings.receipt_note || '')}<br>${esc(S.settings.receipt_footer || '')}</div>` : ''}
  </div>`;
}

function billDocHtml(bill) { return receiptHtml(bill, 'สำหรับร้านเก็บ') + receiptHtml(bill, 'สำหรับลูกค้า'); }
function renderPrintPreview(bill) { $('#print-preview').innerHTML = bill ? billDocHtml(bill) : ''; if (bill) $('#print-no').value = bill.bill_no; }
function printBill(bill) {
  $('#print-area').innerHTML = billDocHtml(bill);
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
  try { b = await loadBill(no); } catch (e) { return toast(e.message); }
  if (!isAdmin() && b.branch !== S.user.branch) return toast('ไม่มีสิทธิ์ดูบิลสาขาอื่น');
  const syncBadge = b._sync === 'pending' ? '<span class="badge pending">รอส่ง</span>' : b._sync === 'error' ? `<span class="badge error">ส่งไม่สำเร็จ: ${esc(b._error)}</span>` : '';
  const canChange = b.status === 'ปกติ';
  const box = openModal(`<h2>บิล ${esc(b.bill_no)} ${b.type === 'คืน' ? '<span class="badge ret">บิลคืน</span>' : ''}
      ${b.status === 'ยกเลิก' ? '<span class="badge void">ยกเลิก</span>' : ''} ${syncBadge}</h2>
    <p>สาขา ${esc(branchInfo(b.branch).name)} · วันที่ ${thDate(b.date)} ${esc(b.time || '')} · ลูกค้า <b>${esc(b.customer_name)}</b> · ผู้ขาย ${esc(b.created_by)}
      ${b.ref_bill ? ' · อ้างอิง ' + esc(b.ref_bill) : ''}${b.approved_by ? ' · อนุมัติโดย ' + esc(b.approved_by) : ''}</p>
    ${b.note ? `<p class="muted">${esc(b.note)}</p>` : ''}
    <table class="grid small"><thead><tr><th>รหัส</th><th>ชื่อ</th><th>รุ่น</th><th class="num">จำนวน</th><th class="num">หน่วยละ</th><th class="num">จำนวนเงิน</th></tr></thead>
    <tbody>${(b.lines || []).map((l) => `<tr class="${l.status === 'ยกเลิก' ? 'void' : ''}"><td>${esc(l.code)}</td><td>${esc(l.name)}</td><td>${esc(l.model)}</td><td class="num">${l.qty}</td><td class="num">${money(l.unit)}</td><td class="num">${money(l.amount)}</td></tr>`).join('')}</tbody></table>
    <p class="num">รวม ${money(b.total)} · ส่วนลด ${money(b.discount)} · <b>สุทธิ ${money(b.net)}</b></p>
    ${b.returns && b.returns.length ? `<p class="muted">มีบิลคืน: ${b.returns.map(esc).join(', ')}</p>` : ''}
    <div class="actions">
      <button data-close>ปิด</button>
      <button id="bd-print">พิมพ์</button>
      ${canChange ? '<button id="bd-edit">แก้ไขบิล</button>' : ''}
      ${canChange && b.type === 'ขาย' ? '<button id="bd-return">ทำบิลคืนอะไหล่</button>' : ''}
      ${canChange ? '<button class="danger" id="bd-void">ยกเลิกบิล</button>' : ''}
    </div>`);
  $('#bd-print', box).onclick = () => { closeModal(); printBill(b); };
  if ($('#bd-edit', box)) $('#bd-edit', box).onclick = () => editBill(b);
  if ($('#bd-void', box)) $('#bd-void', box).onclick = () => voidBill(b);
  if ($('#bd-return', box)) $('#bd-return', box).onclick = () => returnBill(b);
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
    const ap = await askApproval('อนุมัติการแก้ไขบิล ' + b.bill_no, 'การแก้ไขจะถูกบันทึกประวัติไว้ทั้งก่อนและหลังแก้');
    if (!ap) return;
    const cust = customerByName(custText);
    const sign = b.type === 'คืน' ? -1 : 1;
    const newLines = lines.map((l, i) => ({ line: i + 1, code: l.code, name: l.name, model: l.model, qty: sign * l.qty, unit: l.unit, amount: round2(sign * l.qty * l.unit), status: 'ปกติ' }));
    const total = round2(newLines.reduce((s, l) => s + l.amount, 0));
    await enqueue({ type: 'editBill', bill_no: b.bill_no, lines: lines.map((l) => ({ code: l.code, name: l.name, model: l.model, qty: l.qty, unit: l.unit })),
      discount, customer_name: custName, customer_id: cust ? cust.id : '', approval: ap.approval, reason: ap.reason });
    Object.assign(b, { lines: newLines, total, discount, net: round2(total - discount), customer_name: custName, customer_id: cust ? cust.id : '',
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

/* ================================================================ รายงาน */

async function loadReport() {
  const from = $('#rep-from').value || todayISO(), to = $('#rep-to').value || from;
  const branch = isAdmin() ? $('#rep-branch').value : S.user.branch;
  $('#rep-msg').textContent = 'กำลังโหลด…';
  let bills = [], source = '';
  try {
    bills = await api('bills', { branch, from, to });
    source = 'ข้อมูลจากระบบกลาง';
  } catch (e) {
    if (!(e instanceof NetError)) { $('#rep-msg').textContent = e.message; return; }
    source = 'ออฟไลน์: แสดงเฉพาะบิลที่อยู่ในเครื่องนี้';
  }
  // รวมบิลในเครื่องที่ยังไม่ได้ส่ง (หรือแก้ไขแล้วยังไม่ได้ส่ง) ให้เห็นด้วย
  const local = (await DB.all('bills')).filter((b) => b.date >= from && b.date <= to && (branch === '*' || b.branch === branch));
  const map = new Map(bills.map((b) => [b.bill_no, b]));
  local.forEach((b) => { if (b._sync !== 'synced' || !map.has(b.bill_no)) map.set(b.bill_no, b); });
  bills = Array.from(map.values()).sort((a, b) => (a.date + a.bill_no < b.date + b.bill_no ? -1 : 1));
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
      <td class="num">${first ? money(b.net) : ''}</td><td>${first ? st + sync : ''}</td></tr>`;
  }).join('') || '<tr><td colspan="13" class="muted">ไม่มีบิล</td></tr>';
}

function exportExcel() {
  const R = S.report;
  if (!R) return toast('กด "แสดง" ก่อน');
  if (!window.XLSX) return toast('ตัวสร้าง Excel ยังโหลดไม่เสร็จ ลองใหม่อีกครั้ง');
  const pending = R.bills.filter((b) => b._sync === 'pending').length;
  if (pending && !confirm(`มี ${pending} บิลที่ยังไม่ได้ส่งขึ้นระบบ จะดาวน์โหลดรวมไว้ด้วยไหม?`)) return;
  const head = ['ใบเสร็จเลขที่', 'วันที่', 'รหัสอะไหล่', 'ชื่ออะไหล่', 'รุ่น', 'จำนวน', 'หน่วยละ', 'จำนวนเงิน', 'ชื่อผู้ซื้อ', 'รวมเงิน', 'ส่วนลด', 'เหลือ', 'สาขา', 'ประเภท', 'อ้างอิงบิล'];
  const data = [head];
  reportRows(R.bills, false).forEach(({ b, l, first }) => {
    data.push([b.bill_no, thDate(b.date), l.code, l.name, l.model, num(l.qty), round2(num(l.unit)), round2(num(l.amount)),
      first ? b.customer_name : '', first ? round2(num(b.total)) : '', first ? round2(num(b.discount)) : '', first ? round2(num(b.net)) : '',
      branchInfo(b.branch).name, b.type, b.ref_bill || '']);
  });
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [12, 11, 18, 34, 16, 7, 10, 11, 26, 10, 9, 10, 12, 7, 12].map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'บันทึกขายอะไหล่');
  const bname = R.branch === '*' ? 'ทุกสาขา' : branchInfo(R.branch).name;
  const dname = R.from === R.to ? R.from : R.from + '_ถึง_' + R.to;
  XLSX.writeFile(wb, `ขายอะไหล่_${bname}_${dname}.xlsx`);
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
  $('#branch-disc-body').innerHTML = S.branches.map((b) => `<tr><td>${esc(b.name)}</td>
    <td class="num"><input type="number" step="0.01" min="0" max="100" name="${esc(b.code)}" value="${num(b.discount_pct)}" style="width:90px"></td></tr>`).join('');
  $('#branch-disc-msg').textContent = '';
  const odd = S.partsList.filter((p) => p.maxDisc !== '' && p.maxDisc != null).sort((a, b) => a.code < b.code ? -1 : 1);
  $('#special-parts-body').innerHTML = odd.map((p) => `<tr><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td class="num">${money(p.price)}</td>
    <td class="num">${num(p.maxDisc) ? num(p.maxDisc) + '%' : 'ลดไม่ได้'}</td><td><button type="button" data-edit-part="${esc(p.code)}">แก้ไข</button></td></tr>`).join('')
    || '<tr><td colspan="5" class="muted">ยังไม่มี อะไหล่ทุกตัวลดตาม % ของสาขา</td></tr>';
}

async function saveBranchDiscounts(ev) {
  ev.preventDefault();
  const discounts = {};
  for (const inp of $$('#branch-disc-body input')) {
    const v = num(inp.value);
    if (inp.value === '' || v < 0 || v > 100) return toast('ส่วนลดต้องเป็น 0 ถึง 100%');
    discounts[inp.name] = v;
  }
  try {
    const r = await api('saveBranchDiscounts', { discounts });
    S.branches = r.branches;
    const sess = await DB.kv.get('session');
    if (sess) { sess.branches = r.branches; await DB.kv.set('session', sess); }
    $('#branch-disc-msg').textContent = 'บันทึกแล้ว เครื่องอื่นจะได้ค่าใหม่ตอนเปิดแอปครั้งถัดไป';
    loadDiscountAdmin(); $('#branch-disc-msg').textContent = 'บันทึกแล้ว เครื่องอื่นจะได้ค่าใหม่ตอนเปิดแอปครั้งถัดไป';
    renderCart();
  } catch (e) { toast(e instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อนจึงจะแก้ส่วนลดได้' : e.message, 6000); }
}

function renderPendingParts() {
  const rows = S.partsList.filter((p) => p.status === PART_PENDING);
  $('#pending-parts-card').classList.toggle('hidden', !rows.length);
  $('#pending-parts-body').innerHTML = rows.map((p) => `<tr><td>${esc(p.code)}</td><td>${esc(p.name)}</td><td>${esc(p.model)}</td>
    <td class="num">${money(p.price)}</td><td>${esc(p.addedBy)}</td>
    <td><button type="button" class="primary" data-approve="${esc(p.code)}">ยืนยัน</button> <button type="button" data-fix="${esc(p.code)}">แก้ไข</button></td></tr>`).join('');
}

async function onPendingPartClick(e) {
  const code = e.target.dataset.approve || e.target.dataset.fix;
  const p = code && findPart(code);
  if (!p) return;
  if (e.target.dataset.fix) { fillPartForm(p); return; }
  try {
    await api('saveParts', { parts: [{ code: p.code, name: p.name, model: p.model, cost: p.cost, price_ex: '', price: p.price, max_disc: '', approve: true }] });
    toast('ยืนยันอะไหล่ ' + p.code + ' แล้ว');
    await syncParts(); renderPendingParts();
  } catch (err) { toast(err instanceof NetError ? 'ต้องต่ออินเทอร์เน็ตก่อน' : err.message); }
}

function fillPartForm(p) {
  if (!p) return;
  const f = $('#part-form');
  f.code.value = p.code; f.name.value = p.name; f.model.value = p.model; f.price.value = p.price;
  f.cost.value = p.cost === '' || p.cost == null ? '' : p.cost;
  f.max_disc.value = p.maxDisc === '' || p.maxDisc == null ? '' : p.maxDisc;
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

/* ================================================================ ผูกปุ่ม */

function switchTab(name) {
  $$('.tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  $$('.tab').forEach((t) => t.classList.toggle('hidden', t.id !== 'tab-' + name));
  if (name === 'sale') $('#scan-input').focus();
  if (name === 'report' && !S.report) { $('#rep-from').value = $('#rep-from').value || todayISO(); $('#rep-to').value = $('#rep-to').value || todayISO(); loadReport(); }
  if (name === 'print' && S.printBill) renderPrintPreview(S.printBill);
  if (name === 'admin') loadAudit();
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
  };
  $$('.tabs button').forEach((b) => b.onclick = () => switchTab(b.dataset.tab));
  $$('.subtabs button').forEach((b) => b.onclick = () => {
    $$('.subtabs button').forEach((x) => x.classList.toggle('active', x === b));
    $$('.sub').forEach((s) => s.classList.toggle('hidden', s.id !== 'sub-' + b.dataset.sub));
    if (b.dataset.sub === 'users') loadUsers();
    if (b.dataset.sub === 'discounts') loadDiscountAdmin();
    if (b.dataset.sub === 'parts') renderPendingParts();
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
  $('#btn-rep-excel').onclick = exportExcel;
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
  $('#pending-parts-body').onclick = onPendingPartClick;
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
