/* Reno App — renovation accounts kept in OneDrive.
 *
 * Data lives in OneDrive › "Reno App":
 *   tracker.json          all payments, vendors, categories and budget (same format as the Mac app)
 *   Attachments/          bills, receipts and quotations (Attachments/Removed keeps deleted ones)
 *   Backups/              one copy of tracker.json per day, taken before the first change of the day
 *
 * Every change is applied to the newest copy from OneDrive and saved with an eTag check,
 * so two phones (or phone + laptop) editing at the same time never overwrite each other.
 */
'use strict';

const CFG = {
  clientId: 'c010a16c-4e83-493b-b778-86ad96b766e9',
  authority: 'https://login.microsoftonline.com/consumers/oauth2/v2.0',
  scopes: 'Files.ReadWrite User.Read offline_access',
  graph: 'https://graph.microsoft.com/v1.0',
  folder: 'Documents/02 ACCOUNTS/Villa Renovation/Reno App',
  folderLabel: 'Villa Renovation › Reno App',
  version: '1.1',
};
// The one address registered with Microsoft. Any other spelling of it (no slash, capitals,
// index.html, the old reno-accounts name) is sent here first, so sign-in always matches.
const CANON = 'https://dijango999.github.io/reno-app/';
const ON_SITE = location.hostname === 'dijango999.github.io';
if (ON_SITE && location.origin + location.pathname !== CANON) location.replace(CANON + location.search);
const REDIRECT = ON_SITE ? CANON : location.origin + location.pathname.replace(/index\.html$/, '');

const PAL = ['#3B82F6','#F97316','#10B981','#8B5CF6','#F43F5E','#06B6D4','#F59E0B','#6366F1','#84CC16','#EC4899','#14B8A6','#A855F7','#0EA5E9','#EAB308','#64748B'];
const KINDS = ['Bill', 'Receipt', 'Quotation', 'Photo', 'Other'];
const KIND_COLOR = { Bill: '#3B82F6', Receipt: '#10B981', Quotation: '#8B5CF6', Photo: '#F97316', Other: '#64748B' };
const MODES = ['UPI', 'Cash', 'Bank Transfer', 'Cheque', 'Card'];
const EST_TYPES = ['Rough', 'Quotation', 'Final'];

/* ---------------- small helpers ---------------- */
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ls = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch {} },
  del(k) { try { localStorage.removeItem(k); } catch {} },
};
const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
  const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16);
})).toUpperCase();
const slug = s => String(s).toLowerCase().trim().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || uuid().toLowerCase();
const isoNow = () => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');   // same format the Mac app writes
const todayStr = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
const key = s => String(s ?? '').trim().toLowerCase();
const inrF = new Intl.NumberFormat('en-IN', { maximumFractionDigits: 0 });
const inr = n => (n < 0 ? '-' : '') + '₹' + inrF.format(Math.abs(Math.round(+n || 0)));
const trimZ = s => s.replace(/\.?0+$/, '');
const lakh = n => { n = +n || 0; const a = Math.abs(n);
  if (a >= 1e7) return (n < 0 ? '-' : '') + '₹' + trimZ((a / 1e7).toFixed(2)) + ' Cr';
  if (a >= 1e5) return (n < 0 ? '-' : '') + '₹' + trimZ((a / 1e5).toFixed(2)) + ' L';
  return inr(n); };
const parseDay = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const shortDate = s => { const d = parseDay(s); return d ? d.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' }) : '–'; };
const monthLabel = ym => { const [y, m] = ym.split('-'); return new Date(+y, +m - 1, 1).toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }); };
const monthShort = ym => { const [y, m] = ym.split('-'); return new Date(+y, +m - 1, 1).toLocaleDateString('en-IN', { month: 'short' }) + ' ' + y.slice(2); };
const timeStr = d => d ? d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '';
const cmpText = (a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
/** Accepts "55,090", "₹55090" and sums like "55090-1690+200". */
function parseAmount(s) {
  const t = String(s ?? '').replace(/[,₹\s]/g, '');
  if (!t || !/^[+-]?\d+(\.\d+)?([+-]\d+(\.\d+)?)*$/.test(t)) return NaN;
  let total = 0; for (const m of t.matchAll(/([+-]?)(\d+(?:\.\d+)?)/g)) total += (m[1] === '-' ? -1 : 1) * parseFloat(m[2]);
  return Math.round(total * 100) / 100;
}
const initials = n => String(n).split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase() || '?';
function hashColor(s) { let h = 0; for (const c of String(s)) h = (h * 31 + c.charCodeAt(0)) | 0; return PAL[Math.abs(h) % PAL.length]; }
const ICONS = [['civil', '🔨'], ['roof', '🏠'], ['door', '🚪'], ['window', '🪟'], ['electric', '⚡'], ['plumb', '💧'], ['floor', '🧱'], ['tile', '🧱'],
  ['bath', '🚿'], ['kitchen', '🍳'], ['paint', '🖌️'], ['furniture', '🛋️'], ['interlock', '🧱'], ['solar', '☀️'], ['garden', '🌿'], ['landscap', '🌿'],
  ['gate', '🚧'], ['pest', '🐜'], ['carpent', '🪚'], ['contingency', '🛡️'], ['misc', '📦']];
const catIcon = name => (ICONS.find(([k]) => key(name).includes(k)) || [0, '🏷️'])[1];
const isImageName = n => /\.(jpe?g|png|gif|webp|heic|heif|bmp)$/i.test(n);
const docIcon = n => isImageName(n) ? '🖼️' : /\.pdf$/i.test(n) ? '📄' : '📎';

/* ---------------- state ---------------- */
const S = {
  data: null, etag: null, account: ls.get('ra.account'), signedIn: false,
  phase: 'loading',          // loading | welcome | needsignin | import | ready
  tab: ls.get('ra.tab') || 'overview', vendor: null,
  sync: { state: 'idle', at: null, msg: '' }, offline: false,
  filters: { paySearch: '', payCat: '', docSearch: '', docKind: '' },
  sort: {
    payments: ls.get('sort.payments') || 'newest', vendors: ls.get('sort.vendors') || 'spentHigh',
    vendorPays: ls.get('sort.vendorPays') || 'newest', categories: ls.get('sort.categories') || 'original',
    documents: ls.get('sort.documents') || 'newest',
  },
};
let tokens = null; try { tokens = JSON.parse(ls.get('ra.tokens') || 'null'); } catch {}

/* ---------------- Microsoft sign-in (PKCE, no secrets) ---------------- */
class AuthError extends Error {}
class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randStr = n => { const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a).slice(0, n); };

/** silent = renew without showing anything, using the Microsoft sign-in this phone already has. */
async function signIn(chooseAccount, silent) {
  const verifier = randStr(64);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = randStr(24);
  ls.set('ra.pkce', JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({
    client_id: CFG.clientId, response_type: 'code', redirect_uri: REDIRECT, scope: CFG.scopes,
    code_challenge: challenge, code_challenge_method: 'S256', state, response_mode: 'query',
  });
  if (S.account && !chooseAccount) p.set('login_hint', S.account); else p.set('prompt', 'select_account');
  if (silent) { p.set('prompt', 'none'); ls.set('ra.silentAt', String(Date.now())); }
  location.href = `${CFG.authority}/authorize?${p}`;
}

async function handleRedirect() {
  const q = new URLSearchParams(location.search);
  if (!q.has('code') && !q.has('error')) return;
  history.replaceState(null, '', REDIRECT);
  if (q.get('error')) {
    const quiet = ['access_denied', 'login_required', 'interaction_required', 'consent_required'].includes(q.get('error'));
    if (!quiet) toast(q.get('error_description') || 'Sign-in did not complete.');
    return;
  }
  let saved = null; try { saved = JSON.parse(ls.get('ra.pkce') || 'null'); } catch {}
  ls.del('ra.pkce');
  if (!saved || saved.state !== q.get('state')) { toast('That sign-in expired. Please try again.'); return; }
  await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: REDIRECT, code_verifier: saved.verifier });
  try {
    const me = await (await g('/me?$select=userPrincipalName,displayName,mail')).json();
    S.account = me.userPrincipalName || me.mail || me.displayName || 'Microsoft account';
    ls.set('ra.account', S.account);
  } catch {}
}

async function tokenRequest(fields) {
  const body = new URLSearchParams({ client_id: CFG.clientId, scope: CFG.scopes, ...fields });
  let r; try { r = await fetch(`${CFG.authority}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }); }
  catch { throw new HttpError(0, 'No internet connection.'); }
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) throw new AuthError(j.error_description || 'Microsoft sign-in failed.');
  tokens = { access: j.access_token, refresh: j.refresh_token || tokens?.refresh, exp: Date.now() + ((j.expires_in || 3600) - 120) * 1000 };
  ls.set('ra.tokens', JSON.stringify(tokens));
}
let refreshing = null;
async function accessToken() {
  if (tokens?.access && tokens.exp > Date.now()) return tokens.access;
  if (!tokens?.refresh) throw new AuthError('Please sign in.');
  refreshing ||= tokenRequest({ grant_type: 'refresh_token', refresh_token: tokens.refresh }).finally(() => { refreshing = null; });
  await refreshing; return tokens.access;
}
function signOut() {
  ls.del('ra.lock'); ls.del('ra.loc'); LOC = null;
  ls.del('ra.account'); S.account = null;
  ls.del('ra.tokens'); ls.del('ra.cache'); ls.del('ra.etag'); tokens = null;
  S.data = null; S.etag = null; S.phase = 'welcome'; closeSheet(); render();
}

/* ---------------- OneDrive (Microsoft Graph) ---------------- */
async function g(path, opts = {}) {
  const url = path.startsWith('http') ? path : CFG.graph + path;
  for (let i = 0; i < 2; i++) {
    const t = await accessToken();
    let r; try { r = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: 'Bearer ' + t } }); }
    catch { throw new HttpError(0, 'No internet connection.'); }
    if (r.status === 401 && i === 0) { tokens.exp = 0; continue; }
    if (!r.ok) { let m = ''; try { m = (await r.json()).error?.message; } catch {} throw new HttpError(r.status, m || `OneDrive error ${r.status}`); }
    return r;
  }
}
/* ---- where the data lives: a OneDrive folder each person chooses ----
 * The choice is saved in OneDrive › Apps › Reno App › location.json, so every phone or
 * laptop signed in to the same account finds the same folder without asking again. */
const POINTER = 'Apps/Reno App/location.json';
const LEGACY_PATH = 'Documents/02 ACCOUNTS/Villa Renovation/Reno App';   // where version 1 kept the data
let LOC = null; try { LOC = JSON.parse(ls.get('ra.loc') || 'null'); } catch {}
const encRel = rel => rel.split('/').filter(Boolean).map(encodeURIComponent).join('/');
const item = rel => { if (!LOC) throw new Error('Choose where to keep your data first.'); return `/drives/${LOC.driveId}/items/${LOC.id}:/${encRel(rel)}`; };
const locFull = (l = LOC) => l ? 'OneDrive › ' + (l.path ? l.path.split('/').join(' › ') : 'My files') : 'not chosen';
const locShort = (l = LOC) => l ? (l.path ? l.path.split('/').slice(-2).join(' › ') : 'My files') : '';
function folderPath(meta) {
  const parent = decodeURIComponent((meta.parentReference?.path || '').replace(/^\/drive(s\/[^/]+)?\/root:?\/?/, ''));
  return meta.root ? '' : [parent, meta.name].filter(Boolean).join('/');
}
function saveLoc(l) {
  const changed = !LOC || LOC.id !== l.id;
  LOC = { driveId: l.driveId, id: l.id, path: l.path || '' }; ls.set('ra.loc', JSON.stringify(LOC));
  if (changed) { ls.del('ra.cache'); ls.del('ra.etag'); S.data = null; S.etag = null; }
}
async function readPointer() {
  try {
    const m = await (await g(`/me/drive/root:/${encRel(POINTER)}`)).json();
    const t = await (await fetch(m['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).text();
    const p = JSON.parse(t); return p && p.id && p.driveId ? p : null;
  } catch (e) { if (e.status === 404 || e instanceof SyntaxError) return null; throw e; }
}
async function writePointer(l) {
  await g(`/me/drive/root:/${encRel(POINTER)}:/content`, { method: 'PUT', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ driveId: l.driveId, id: l.id, path: l.path, savedAt: isoNow(), by: S.account || '' }, null, 2) });
}
/** Finds the chosen folder. false = nothing chosen yet (ask the person). */
async function resolveLocation() {
  const p = await readPointer();
  if (p) {
    try { const m = await (await g(`/drives/${p.driveId}/items/${p.id}?$select=id,name,root,parentReference,deleted`)).json();
      saveLoc({ driveId: p.driveId, id: p.id, path: folderPath(m) }); return true; }
    catch (e) { if (e.status !== 404) throw e; }        // folder was deleted → choose again
  }
  if (LOC) { await writePointer(LOC).catch(() => {}); return true; }
  try {
    const f = await (await g(`/me/drive/root:/${encRel(LEGACY_PATH)}?$select=id,name,parentReference`)).json();
    const l = { driveId: f.parentReference.driveId, id: f.id, path: LEGACY_PATH };
    saveLoc(l); await writePointer(l).catch(() => {}); return true;
  } catch (e) { if (e.status !== 404) throw e; }
  return false;
}

/** Newest tracker.json from OneDrive, or null if there isn't one yet. */
async function loadRemote() {
  let meta;
  try { meta = await (await g(item('tracker.json'))).json(); }
  catch (e) { if (e.status === 404) return null; throw e; }
  const url = meta['@microsoft.graph.downloadUrl'];
  let text;
  if (url) { const r = await fetch(url, { cache: 'no-store' }); if (!r.ok) throw new HttpError(r.status, 'Could not download your data.'); text = await r.text(); }
  else text = await (await g(item('tracker.json') + ':/content', { cache: 'no-store' })).text();
  let data; try { data = JSON.parse(text); } catch { throw new Error('tracker.json in OneDrive could not be read. Nothing was changed.'); }
  return { data: normalize(data), etag: meta.eTag, text };
}
function normalize(d) {
  d.categories ||= []; d.vendors ||= []; d.payments ||= []; d.items ||= []; d.totalBudget = +d.totalBudget || 0;
  return d;
}
/** Key-sorted, pretty JSON (same layout as the Mac app). */
function stableJSON(v) {
  const sortObj = x => Array.isArray(x) ? x.map(sortObj) : (x && typeof x === 'object')
    ? Object.keys(x).sort().reduce((o, k) => (o[k] = sortObj(x[k]), o), {}) : x;
  return JSON.stringify(sortObj(v), null, 2);
}
async function putTracker(data, etag, createOnly) {
  const headers = { 'Content-Type': 'application/json' };
  if (createOnly) headers['If-None-Match'] = '*'; else if (etag) headers['If-Match'] = etag;
  const r = await g(item('tracker.json') + ':/content', { method: 'PUT', headers, body: stableJSON(data) });
  return (await r.json()).eTag;
}

async function dailyBackup(text) {
  const day = todayStr();
  if (ls.get('ra.backup') === day) return;
  try {
    await g(item(`Backups/tracker-${day}.json`) + ':/content?@microsoft.graph.conflictBehavior=fail',
      { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: text });
  } catch (e) { if (e.status !== 409) { console.warn('backup', e); return; } }
  ls.set('ra.backup', day);
}

let saving = Promise.resolve();
/** Applies `change` to the newest data in OneDrive and saves it. Retries if someone else saved first. */
function mutate(change) {
  const run = async () => {
    if (!navigator.onLine) throw new Error('You are offline. Connect to the internet to save changes.');
    setSync('busy', 'Saving…');
    for (let attempt = 0; attempt < 5; attempt++) {
      const remote = await loadRemote();
      if (!remote) throw new Error('Your data folder is missing in OneDrive. Nothing was saved.');
      await dailyBackup(remote.text);
      const d = remote.data;
      change(d);
      d.savedAt = isoNow();
      try {
        const etag = await putTracker(d, remote.etag);
        applyData(d, etag);
        setSync('ok');
        return true;
      } catch (e) { if (e.status === 412 || e.status === 409) continue; throw e; }
    }
    throw new Error('OneDrive is busy. Please try again.');
  };
  const p = saving.then(run, run);
  saving = p.catch(() => {});
  return p.catch(e => { setSync('bad', e.message); toast(e.message); return false; });
}
function applyData(d, etag) {
  S.data = d; S.etag = etag;
  ls.set('ra.cache', JSON.stringify(d)); ls.set('ra.etag', etag || '');
  if (S.phase !== 'ready') S.phase = 'ready';
  render();
}
async function pull(force) {
  if (!tokens) return;
  try {
    if (!force && S.etag) {
      const meta = await (await g(item('tracker.json') + '?$select=eTag')).json();
      if (meta.eTag === S.etag) { setSync('ok'); return; }
    }
    setSync('busy', 'Syncing…');
    const remote = await loadRemote();
    if (!remote) { S.phase = 'import'; render(); return; }
    applyData(remote.data, remote.etag);
    setSync('ok');
  } catch (e) {
    if (e instanceof AuthError) { if (trySilent()) return; S.phase = 'needsignin'; render(); return; }
    if (e.status === 404) { S.phase = 'import'; render(); return; }
    setSync('bad', e.status === 0 ? 'Offline — showing the last synced copy' : e.message);
  }
}
function setSync(state, msg = '') {
  S.sync = { state, msg, at: state === 'ok' ? new Date() : S.sync.at };
  const el = $('#syncline'); if (el) el.innerHTML = syncLine();
}

/* ---- attachments ---- */
async function shrinkImage(file) {
  if (!/^image\/(jpe?g|png|webp|heic|heif)$/i.test(file.type) || file.size < 1.5e6) return file;
  try {
    const bmp = await createImageBitmap(file);
    const scale = Math.min(1, 2400 / Math.max(bmp.width, bmp.height));
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
    c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
    const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.85));
    return blob && blob.size < file.size ? new File([blob], file.name.replace(/\.\w+$/, '') + '.jpg', { type: 'image/jpeg' }) : file;
  } catch { return file; }
}
async function uploadBlob(rel, blob, onProgress) {
  if (blob.size <= 4 * 1024 * 1024) {
    await g(item(rel) + ':/content?@microsoft.graph.conflictBehavior=replace',
      { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob });
    onProgress?.(1); return;
  }
  const s = await (await g(item(rel) + ':/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) })).json();
  const chunk = 320 * 1024 * 16;
  for (let start = 0; start < blob.size; start += chunk) {
    const end = Math.min(start + chunk, blob.size);
    const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!r.ok) throw new HttpError(r.status, 'Upload failed. Please try again.');
    onProgress?.(end / blob.size);
  }
}
function attachmentFileName(file, owner, label, kind) {
  const ext = (/\.(\w{1,5})$/.exec(file.name)?.[1] || (file.type.split('/')[1] || 'bin')).toLowerCase();
  const clean = s => String(s || '').replace(/[\\/:*?"<>|#%{}~&]+/g, ' ').replace(/\s+/g, ' ').trim();
  const code = Math.random().toString(16).slice(2, 6).toUpperCase();
  return [todayStr(), clean(owner), clean(label), kind, code].filter(Boolean).join(' ').slice(0, 150) + '.' + ext;
}
async function uploadAttachment(file, owner, label, kind) {
  const f = await shrinkImage(file);
  const name = attachmentFileName(f, owner, label, kind);
  await uploadBlob('Attachments/' + name, f);
  return { id: uuid(), file: name, name: file.name || name, kind, addedAt: isoNow() };
}
async function moveToRemoved(a) {
  try {
    try { await g(item('Attachments') + ':/children', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Removed', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) }); }
    catch (e) { if (e.status !== 409) throw e; }
    const removed = await (await g(item('Attachments/Removed') + '?$select=id')).json();
    await g(item('Attachments/' + a.file), { method: 'PATCH', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ parentReference: { id: removed.id } }) });
  } catch (e) { console.warn('move to Removed', e); }
}


/* ---------------- OneDrive folder browser ---------------- */
function folderBrowser(box, { onChosen, current }) {
  const st = { stack: [{ name: 'My files', driveId: null, id: null, path: '' }], items: [], loading: true, err: '' };
  const here = () => st.stack[st.stack.length - 1];
  async function load() {
    st.loading = true; st.err = ''; draw();
    try {
      const h = here();
      const url = h.id ? `/drives/${h.driveId}/items/${h.id}/children` : '/me/drive/root/children';
      let next = url + '?$select=id,name,folder,file,parentReference,root&$top=200', all = [];
      while (next) { const j = await (await g(next)).json(); all = all.concat(j.value || []); next = j['@odata.nextLink']; }
      if (!h.id) { try { const r = await (await g('/me/drive/root?$select=id,parentReference')).json(); h.id = r.id; h.driveId = r.parentReference?.driveId; } catch {} }
      st.items = all.filter(x => x.folder || x.name === 'tracker.json').sort((a, b) => (!!b.folder - !!a.folder) || cmpText(a.name, b.name));
      if (!h.driveId && all[0]) h.driveId = all[0].parentReference?.driveId;
    } catch (e) { st.err = e.message; }
    st.loading = false; draw();
  }
  function draw() {
    const h = here(), hasData = st.items.some(x => x.name === 'tracker.json'), atRoot = st.stack.length === 1;
    box.innerHTML = `<div class="list" style="text-align:left">
      <div class="gh" style="text-transform:none;letter-spacing:0;align-items:center;gap:8px">
        ${atRoot ? '' : `<button class="linkbtn" data-b="up" style="padding:0">‹ Back</button>`}
        <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">📁 ${esc(st.stack.map(x => x.name).join(' › '))}</span></div>
      ${st.loading ? '<div class="empty"><span class="spin"></span></div>' : st.err ? `<div class="empty" style="color:var(--bad)">${esc(st.err)}</div>` :
        (st.items.map((x, i) => x.folder
          ? `<button class="row" data-b="open" data-i="${i}"><span class="ic" style="--c:#3B82F6">📁</span><span class="main"><span class="t">${esc(x.name)}</span><span class="s">${x.folder.childCount} item${x.folder.childCount === 1 ? '' : 's'}</span></span><span class="faint">›</span></button>`
          : `<div class="row" style="cursor:default"><span class="ic" style="--c:#10B981">✅</span><span class="main"><span class="t">Reno App data found here</span><span class="s">tracker.json</span></span></div>`).join('')
          || '<div class="empty">No folders here.</div>')}
    </div>
    <div style="display:flex;gap:8px;margin-top:10px">
      <button class="btn sec" data-b="new" style="flex:1">＋ New folder</button>
      <button class="btn" data-b="use" style="flex:1.4" ${atRoot || st.loading ? 'disabled' : ''}>${hasData ? 'Use this data' : 'Use this folder'}</button>
    </div>
    <div class="small faint" style="margin-top:6px">${atRoot ? 'Open a folder, or tap “New folder” to make one (e.g. “Reno App”).' : `Data will be kept in <b>${esc(locFull({ path: here().path }))}</b>.`}</div>`;
  }
  box.addEventListener('click', async e => {
    const b = e.target.closest('[data-b]'); if (!b || st.loading) return;
    if (b.dataset.b === 'up') { st.stack.pop(); load(); }
    else if (b.dataset.b === 'open') { const x = st.items[+b.dataset.i]; st.stack.push({ name: x.name, id: x.id, driveId: x.parentReference?.driveId || here().driveId, path: [here().path, x.name].filter(Boolean).join('/') }); load(); }
    else if (b.dataset.b === 'new') {
      const name = (prompt('Name for the new folder:', 'Reno App') || '').trim().replace(/[\\/:*?"<>|]/g, ' ');
      if (!name) return;
      try {
        const h = here(), url = h.id ? `/drives/${h.driveId}/items/${h.id}/children` : '/me/drive/root/children';
        const f = await (await g(url, { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }) })).json();
        st.stack.push({ name: f.name, id: f.id, driveId: f.parentReference?.driveId || h.driveId, path: [h.path, f.name].filter(Boolean).join('/') }); load();
      } catch (err) { toast(err.message); }
    } else if (b.dataset.b === 'use') { const h = here(); onChosen({ driveId: h.driveId, id: h.id, path: h.path }); }
  });
  load();
}

/** Switches to a folder. If it is empty and there is data now, offers to copy it there. */
async function chooseFolder(f) {
  try {
    if (LOC && f.id === LOC.id) { closeSheet(); return; }
    let hasData = false;
    try { await g(`/drives/${f.driveId}/items/${f.id}:/tracker.json?$select=id`); hasData = true; } catch (e) { if (e.status !== 404) throw e; }
    const prev = LOC, current = S.data && prev ? await loadRemote().catch(() => null) : null;
    if (!hasData && current) {
      if (!confirm(`Copy your data (${current.data.payments.length} payments and all bills) to “${locFull(f)}” and use it from there?\n\nThe old folder is left as it is, as a safe copy.`)) return;
      toast('Copying your data…');
      try {   // bills: OneDrive copies the whole Attachments folder by itself
        const att = await (await g(item('Attachments') + '?$select=id')).json();
        await g(`/drives/${prev.driveId}/items/${att.id}/copy`, { method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ parentReference: { driveId: f.driveId, id: f.id } }) });
      } catch (e) { if (e.status !== 404) throw e; }
      saveLoc(f);
      await putTracker(current.data, null, true);
    } else if (hasData && current && !confirm(`“${locFull(f)}” already has Reno App data. Switch to that data? (Your current folder is not changed.)`)) return;
    else saveLoc(f);
    await writePointer(LOC);
    closeSheet(); S.phase = 'loading'; render();
    await pull(true);
    if (S.phase === 'ready') toast('Now using ' + locShort());
  } catch (e) { toast(e.message); }
}
function changeFolderSheet() {
  openSheet('Where to keep your data', `<p class="small muted" style="margin:0 0 10px">Now: <b>${esc(locFull())}</b></p><div id="sheetBrowser"></div>`, { noSave: true });
  folderBrowser($('#sheetBrowser'), { onChosen: f => chooseFolder(f) });
}

/* ---------------- figures ---------------- */
const D = () => S.data;
const budgeted = c => (+c.quoted || 0) * (1 + (+c.contingency || 0));
const totalSpent = () => D().payments.reduce((s, p) => s + (+p.amount || 0), 0);
const totalBudgeted = () => D().categories.reduce((s, c) => s + budgeted(c), 0);
const spentIn = name => D().payments.filter(p => p.category === name).reduce((s, p) => s + (+p.amount || 0), 0);
function catColor(name) {
  const cats = D().categories, i = cats.findIndex(c => c.name === name);
  if (i < 0) return '#94A3B8';
  return PAL[(((cats[i].colorIndex ?? i) % PAL.length) + PAL.length) % PAL.length];
}
const vendorColor = name => hashColor(name);
function vendorStats() {
  const map = new Map();
  for (const v of D().vendors) map.set(v.name, { name: v.name, trade: v.trade || '', phone: v.phone || '', total: 0, count: 0, first: null, last: null, byCat: {} });
  for (const p of D().payments) {
    if (!p.vendor) continue;
    let s = map.get(p.vendor);
    if (!s) { s = { name: p.vendor, trade: '', phone: '', total: 0, count: 0, first: null, last: null, byCat: {} }; map.set(p.vendor, s); }
    s.total += +p.amount || 0; s.count++;
    if (!s.first || p.date < s.first) s.first = p.date;
    if (!s.last || p.date > s.last) s.last = p.date;
    s.byCat[p.category] = (s.byCat[p.category] || 0) + (+p.amount || 0);
  }
  return [...map.values()];
}
function allDocuments() {
  const out = [];
  for (const p of D().payments) for (const a of p.attachments || [])
    out.push({ a, owner: p.vendor || p.category, detail: `${p.item || p.category} · ${inr(p.amount)}`, date: p.date, ref: { type: 'payment', id: p.id } });
  for (const v of D().vendors) for (const a of v.attachments || [])
    out.push({ a, owner: v.name, detail: 'Vendor document', date: (a.addedAt || '').slice(0, 10), ref: { type: 'vendor', id: v.id } });
  for (const c of D().categories) for (const a of c.attachments || [])
    out.push({ a, owner: c.name, detail: 'Category quotation', date: (a.addedAt || '').slice(0, 10), ref: { type: 'category', id: c.id } });
  return out;
}

/* ---------------- sorting ---------------- */
const SORTS = {
  payments: [['newest', 'Date: newest first'], ['oldest', 'Date: oldest first'], ['amountHigh', 'Amount: high to low'], ['amountLow', 'Amount: low to high'],
    ['vendor', 'Vendor: A to Z'], ['category', 'Category: A to Z'], ['added', 'Recently added']],
  vendors: [['spentHigh', 'Total paid: high to low'], ['spentLow', 'Total paid: low to high'], ['name', 'Name: A to Z'], ['nameDesc', 'Name: Z to A'],
    ['count', 'Most payments'], ['latest', 'Latest payment']],
  categories: [['original', 'My order'], ['name', 'Name: A to Z'], ['spentHigh', 'Spent: high to low'], ['budgetHigh', 'Budget: high to low'],
    ['usedHigh', '% used: highest first'], ['leftHigh', 'Left to spend: most first']],
  documents: [['newest', 'Date: newest first'], ['oldest', 'Date: oldest first'], ['added', 'Recently added'], ['name', 'File name: A to Z'],
    ['owner', 'Vendor / category: A to Z'], ['kind', 'Type']],
};
SORTS.vendorPays = SORTS.payments;
function sortPayments(list, how) {
  const byDate = (a, b) => cmpText(b.date + (b.createdAt || ''), a.date + (a.createdAt || ''));
  const l = [...list];
  switch (how) {
    case 'oldest': return l.sort((a, b) => -byDate(a, b));
    case 'amountHigh': return l.sort((a, b) => (b.amount - a.amount) || byDate(a, b));
    case 'amountLow': return l.sort((a, b) => (a.amount - b.amount) || byDate(a, b));
    case 'vendor': return l.sort((a, b) => cmpText(a.vendor, b.vendor) || byDate(a, b));
    case 'category': return l.sort((a, b) => cmpText(a.category, b.category) || byDate(a, b));
    case 'added': return l.sort((a, b) => cmpText(b.createdAt || '', a.createdAt || ''));
    default: return l.sort(byDate);
  }
}
function sortVendors(list, how) {
  const l = [...list];
  switch (how) {
    case 'spentLow': return l.sort((a, b) => (a.total - b.total) || cmpText(a.name, b.name));
    case 'name': return l.sort((a, b) => cmpText(a.name, b.name));
    case 'nameDesc': return l.sort((a, b) => cmpText(b.name, a.name));
    case 'count': return l.sort((a, b) => (b.count - a.count) || (b.total - a.total));
    case 'latest': return l.sort((a, b) => cmpText(b.last || '', a.last || ''));
    default: return l.sort((a, b) => (b.total - a.total) || cmpText(a.name, b.name));
  }
}
function sortCategories(how) {
  const l = D().categories.map((c, i) => ({ c, i }));
  const used = c => budgeted(c) > 0 ? spentIn(c.name) / budgeted(c) : (spentIn(c.name) > 0 ? 9e9 : 0);
  switch (how) {
    case 'name': return l.sort((a, b) => cmpText(a.c.name, b.c.name));
    case 'spentHigh': return l.sort((a, b) => spentIn(b.c.name) - spentIn(a.c.name));
    case 'budgetHigh': return l.sort((a, b) => budgeted(b.c) - budgeted(a.c));
    case 'usedHigh': return l.sort((a, b) => used(b.c) - used(a.c));
    case 'leftHigh': return l.sort((a, b) => (budgeted(b.c) - spentIn(b.c.name)) - (budgeted(a.c) - spentIn(a.c.name)));
    default: return l;
  }
}
function sortDocs(list, how) {
  const l = [...list];
  switch (how) {
    case 'oldest': return l.sort((a, b) => cmpText(a.date, b.date) || cmpText(a.a.addedAt || '', b.a.addedAt || ''));
    case 'added': return l.sort((a, b) => cmpText(b.a.addedAt || '', a.a.addedAt || ''));
    case 'name': return l.sort((a, b) => cmpText(a.a.name, b.a.name));
    case 'owner': return l.sort((a, b) => cmpText(a.owner, b.owner) || cmpText(b.date, a.date));
    case 'kind': return l.sort((a, b) => (KINDS.indexOf(a.a.kind) - KINDS.indexOf(b.a.kind)) || cmpText(b.date, a.date));
    default: return l.sort((a, b) => cmpText(b.date, a.date) || cmpText(b.a.addedAt || '', a.a.addedAt || ''));
  }
}
function sortControl(which) {
  const cur = S.sort[which], label = (SORTS[which].find(o => o[0] === cur) || SORTS[which][0])[1];
  return `<label class="sortsel">⇅ Sorted by ${esc(label.toLowerCase())} ▾
    <select data-sort="${which}" aria-label="Sort">${SORTS[which].map(([v, t]) => `<option value="${v}" ${v === cur ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select></label>`;
}

/* ---------------- rendering ---------------- */
const TABS = [
  ['overview', 'Overview', '<path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z"/>'],
  ['payments', 'Payments', '<circle cx="12" cy="12" r="8.5"/><path d="M9 8h6M9 11h6M13 8c2 0 2 6-4 6l5 4"/>'],
  ['vendors', 'Vendors', '<circle cx="9" cy="9" r="3.2"/><circle cx="17" cy="10" r="2.6"/><path d="M3.5 19c.6-3.3 3-5 5.5-5s4.9 1.7 5.5 5M14.5 15.2c.8-.6 1.6-.9 2.5-.9 2 0 3.6 1.4 4 4.2"/>'],
  ['budget', 'Budget', '<path d="M12 3.5a8.5 8.5 0 1 0 8.5 8.5H12z"/><path d="M14.5 2.8A8.5 8.5 0 0 1 21 9.5h-6.5z"/>'],
  ['documents', 'Documents', '<path d="M7 3.5h7l4 4V20a.5.5 0 0 1-.5.5h-10A.5.5 0 0 1 7 20z"/><path d="M14 3.5V8h4M9.5 12h5M9.5 15.5h5"/>'],
];
function syncLine() {
  const s = S.sync;
  if (s.state === 'busy') return `<span class="dot busy"></span>${esc(s.msg || 'Syncing…')}`;
  if (s.state === 'bad') return `<span class="dot bad"></span>${esc(s.msg)}`;
  return `<span class="dot"></span>Synced with OneDrive${s.at ? ' · ' + timeStr(s.at) : ''}`;
}

function render() {
  const app = $('#app');
  if (S.phase === 'loading') { app.innerHTML = `<div class="welcome"><img class="logo" src="icons/icon-192.png" alt=""><div class="muted"><span class="spin"></span> Loading…</div></div>`; return; }
  if (S.phase === 'welcome' || S.phase === 'needsignin') { app.innerHTML = welcomeView(); return; }
  if (S.phase === 'import') { app.innerHTML = importView(); return; }
  if (S.phase === 'choose') {
    app.innerHTML = `<div class="welcome"><img class="logo" src="icons/icon-192.png" alt="">
      <h1>Where should Reno App keep your data?</h1>
      <p class="muted" style="margin:0">Pick a folder in your OneDrive (signed in as <b>${esc(S.account || '')}</b>). Everything — payments, bills, backups — is saved there. You can change it later.</p>
      <div id="browser"></div>
      <button class="linkbtn" data-act="signin-other" style="color:var(--muted)">Use a different Microsoft account</button></div>`;
    folderBrowser($('#browser'), { onChosen: f => chooseFolder(f) });
    return;
  }
  const scrollY = window.scrollY;
  const title = S.vendor ? S.vendor : (TABS.find(t => t[0] === S.tab) || TABS[0])[1];
  let body = '';
  try {
    body = S.vendor ? vendorDetailView(S.vendor) : ({ overview: overviewView, payments: paymentsView, vendors: vendorsView, budget: budgetView, documents: documentsView }[S.tab] || overviewView)();
  } catch (e) { console.error(e); body = `<div class="notice">Something went wrong showing this page: ${esc(e.message)}</div>`; }
  app.innerHTML = `
    <header class="top" id="top">
      <div class="topbar">
        ${S.vendor ? `<button class="back" data-act="back">‹ Vendors</button>` : ''}
        <h1 style="${S.vendor ? 'font-size:20px' : ''}">${esc(title)}</h1>
        ${topActions()}
        <button class="iconbtn" data-act="account" aria-label="Account and data">👤</button>
      </div>
      <div class="sync" id="syncline">${syncLine()}</div>
    </header>
    <main>${body}</main>
    <button class="fab" data-act="add-payment"><span style="font-size:20px;line-height:1">＋</span> Add Payment</button>
    <nav class="tabs">${TABS.map(([id, label, svg]) => `<button class="${!S.vendor && S.tab === id ? 'on' : ''}" data-tab="${id}">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${svg}</svg>${label}</button>`).join('')}</nav>`;
  window.scrollTo(0, scrollY);
}
function topActions() {
  if (S.vendor) return `<button class="btn sm sec" data-act="edit-vendor" data-name="${esc(S.vendor)}">Edit</button>`;
  if (S.tab === 'vendors') return `<button class="btn sm" data-act="new-vendor">Add Vendor</button>`;
  if (S.tab === 'budget') return `<button class="btn sm" data-act="new-category">Add Category</button>`;
  return '';
}

function welcomeView() {
  const again = S.phase === 'needsignin';
  return `<div class="welcome">
    <img class="logo" src="icons/icon-192.png" alt="">
    <h1>Reno App</h1>
    <p class="muted" style="margin:0">Your renovation payments, vendors, budget and bills — kept in your OneDrive, the same on your phone and your laptop.</p>
    <button class="btn wide" data-act="signin">${again ? 'Continue as ' + esc(S.account || 'your account') : 'Sign in with Microsoft'}</button>
    ${again ? `<button class="linkbtn" data-act="signin-other">Use a different account</button>` : `<p class="small faint" style="margin:0">Use the Microsoft account of the OneDrive where you want your data kept.</p>`}
    ${again ? `<p class="small faint" style="margin:0">Microsoft needs you to confirm your account again.</p>` : ''}
  </div>`;
}
function importView() {
  return `<div class="welcome">
    <img class="logo" src="icons/icon-192.png" alt="">
    <h1>Bring in your data</h1>
    <p class="muted" style="margin:0">Signed in as <b>${esc(S.account || '')}</b>. There is no Reno App data in <b>${esc(locFull())}</b> yet.</p>
    <div class="group" style="margin:0">
      <h4>From the Mac app</h4>
      <p class="small muted" style="margin:0 0 10px">On your laptop, choose the <b>Renovation Tracker</b> folder (the one with tracker.json, Attachments and Backups). Everything is copied into OneDrive; the original folder is not changed.</p>
      <label class="btn wide" style="position:relative">Choose Renovation Tracker folder…
        <input type="file" id="importDir" webkitdirectory directory multiple style="position:absolute;inset:0;opacity:0"></label>
      <label class="btn sec wide" style="position:relative;margin-top:8px">Or choose files (tracker.json + bills)…
        <input type="file" id="importFiles" multiple style="position:absolute;inset:0;opacity:0"></label>
      <div id="importStatus" class="small muted" style="margin-top:10px"></div>
    </div>
    <button class="linkbtn" data-act="start-empty">Start with empty data instead</button>
    <button class="linkbtn" data-act="change-folder">Choose a different folder</button>
    <button class="linkbtn" data-act="signin-other" style="color:var(--muted)">Use a different Microsoft account</button>
  </div>`;
}

function overviewView() {
  const d = D(), spent = totalSpent(), budget = d.totalBudget, planned = totalBudgeted();
  const cats = d.categories.map(c => ({ name: c.name, color: catColor(c.name), spent: spentIn(c.name), b: budgeted(c) })).filter(c => c.spent > 0).sort((a, b) => b.spent - a.spent);
  const pending = d.payments.filter(p => p.status && p.status !== 'Paid');
  const vstats = vendorStats().filter(v => v.total > 0).sort((a, b) => b.total - a.total);
  const over = d.categories.filter(c => budgeted(c) > 0 && spentIn(c.name) > budgeted(c));
  const max = Math.max(budget, spent, 1);
  // months
  const byMonth = {};
  for (const p of d.payments) { const m = (p.date || '').slice(0, 7); if (!m) continue; (byMonth[m] ||= {}); byMonth[m][p.category] = (byMonth[m][p.category] || 0) + (+p.amount || 0); }
  const months = Object.keys(byMonth).sort().slice(-8);
  const mmax = Math.max(1, ...months.map(m => Object.values(byMonth[m]).reduce((a, b) => a + b, 0)));
  const recent = sortPayments(d.payments, 'newest').slice(0, 5);
  return `<div class="stack">
    <div class="card hero">
      <div class="lbl">Spent so far</div>
      <div class="big num">${inr(spent)}</div>
      <div class="small muted">${budget > 0 ? `${(spent / budget * 100).toFixed(1)}% of ${inr(budget)} budget` : 'Set your total budget on the Budget tab'}</div>
      <div class="seg" style="margin-top:12px">${cats.map(c => `<i style="width:${c.spent / max * 100}%;background:${c.color}"></i>`).join('')}</div>
      <div class="legend">${cats.slice(0, 10).map(c => `<div><span class="sw" style="--c:${c.color}"></span><span class="n">${esc(c.name)}</span><b class="num">${lakh(c.spent)}</b></div>`).join('')}</div>
    </div>
    <div class="grid2 grid4">
      <div class="tile"><div class="lbl">💰 Cash left</div><div class="val num" style="color:${budget - spent < 0 ? 'var(--bad)' : 'inherit'}">${lakh(budget - spent)}</div><div class="sub">of ${lakh(budget)}</div></div>
      <div class="tile"><div class="lbl">🧾 Payments</div><div class="val num">${d.payments.length}</div><div class="sub">${pending.length ? pending.length + ' pending' : 'All paid'}</div></div>
      <div class="tile"><div class="lbl">⏳ Pending</div><div class="val num">${lakh(pending.reduce((s, p) => s + (+p.amount || 0), 0))}</div><div class="sub">${pending.length} payment${pending.length === 1 ? '' : 's'}</div></div>
      <div class="tile"><div class="lbl">👷 Vendors paid</div><div class="val num">${vstats.length}</div><div class="sub">${vstats[0] ? 'Top: ' + esc(vstats[0].name) : '–'}</div></div>
    </div>
    ${budget > 0 && planned > budget ? `<div class="notice"><span class="grow">Your category estimates add up to ${lakh(planned)} — ${lakh(planned - budget)} more than your ${lakh(budget)} budget.</span><button data-tab="budget">Review</button></div>` : ''}
    ${over.length ? `<div class="notice warn"><span class="grow">Over budget: ${over.map(c => esc(c.name)).join(', ')}.</span><button data-tab="budget">See</button></div>` : ''}
    ${months.length ? `<div class="card"><div class="sec-h" style="margin:0"><h2>Spending by month</h2></div>
      <div class="bars">${months.map(m => { const parts = byMonth[m], tot = Object.values(parts).reduce((a, b) => a + b, 0);
        return `<div class="col"><small class="num">${lakh(tot)}</small><div class="stackbar" style="height:${Math.max(3, tot / mmax * 100)}%">${Object.entries(parts).sort((a, b) => b[1] - a[1]).map(([c, v]) => `<i style="height:${v / tot * 100}%;background:${catColor(c)}"></i>`).join('')}</div><small>${monthShort(m)}</small></div>`; }).join('')}</div></div>` : ''}
    <div class="sec-h"><h2>Top vendors</h2><button class="linkbtn" data-tab="vendors">See all</button></div>
    <div class="list">${vstats.slice(0, 5).map(v => vendorRow(v, spent)).join('') || '<div class="empty">No payments yet.</div>'}</div>
    <div class="sec-h"><h2>Recent payments</h2><button class="linkbtn" data-tab="payments">See all</button></div>
    <div class="list">${recent.map(paymentRow).join('') || '<div class="empty">No payments yet. Tap Add Payment.</div>'}</div>
    <button class="btn sec wide" data-act="export">⬇︎ Export to Excel</button>
  </div>`;
}

function paymentRow(p) {
  const c = catColor(p.category), n = (p.attachments || []).length;
  return `<button class="row" data-act="edit-payment" data-id="${esc(p.id)}">
    <span class="ic" style="--c:${c}">${catIcon(p.category)}</span>
    <span class="main"><span class="t">${esc(p.item || p.category || 'Payment')}</span>
      <span class="s">${esc(p.vendor || '—')} · ${shortDate(p.date)}${p.mode ? ' · ' + esc(p.mode) : ''}</span>
      <span><span class="chip" style="--c:${c}">${esc(p.category || 'No category')}</span></span></span>
    <span class="amt"><span class="num">${inr(p.amount)}</span>
      <span style="display:flex;gap:4px">${p.status && p.status !== 'Paid' ? `<span class="pill warn">${esc(p.status)}</span>` : ''}${n ? `<span class="pill">📎 ${n}</span>` : ''}</span></span>
  </button>`;
}
function vendorRow(v, total) {
  const c = vendorColor(v.name);
  return `<button class="row" data-act="open-vendor" data-name="${esc(v.name)}">
    <span class="av" style="--c:${c}">${esc(initials(v.name))}</span>
    <span class="main"><span class="t">${esc(v.name)}</span>
      <span class="s">${v.count ? `${v.count} payment${v.count === 1 ? '' : 's'} · last ${shortDate(v.last)}` : 'No payments yet'}${v.trade ? ' · ' + esc(v.trade) : ''}</span>
      <span class="meter" style="height:4px"><i style="width:${Math.min(100, v.total / Math.max(total, 1) * 100)}%;--c:${c}"></i></span></span>
    <span class="amt num">${inr(v.total)}</span>
  </button>`;
}

function paymentsView() {
  const d = D(), q = key(S.filters.paySearch), cat = S.filters.payCat, how = S.sort.payments;
  let list = d.payments.filter(p => (!cat || p.category === cat) &&
    (!q || [p.item, p.vendor, p.category, p.notes, p.invoice, p.mode, String(p.amount), ...(p.attachments || []).map(a => a.name)].join(' ').toLowerCase().includes(q)));
  list = sortPayments(list, how);
  const sum = list.reduce((s, p) => s + (+p.amount || 0), 0);
  let groups = '';
  if (how === 'newest' || how === 'oldest') {
    const by = {}; for (const p of list) (by[(p.date || '').slice(0, 7)] ||= []).push(p);
    const ms = Object.keys(by).sort(); if (how === 'newest') ms.reverse();
    groups = ms.map(m => `<div class="list"><div class="gh"><span>${m ? monthLabel(m) : 'No date'}</span><span class="num">${inr(by[m].reduce((s, p) => s + (+p.amount || 0), 0))}</span></div>${by[m].map(paymentRow).join('')}</div>`).join('');
  } else groups = `<div class="list">${list.map(paymentRow).join('')}</div>`;
  const usedCats = d.categories.filter(c => d.payments.some(p => p.category === c.name));
  return `<div class="stack">
    <div class="tools">
      <div class="search">🔍<input type="search" id="paySearch" placeholder="Search item, vendor, notes, amount" value="${esc(S.filters.paySearch)}"></div>
      <div class="chips"><button class="fchip ${!cat ? 'on' : ''}" data-paycat="">All</button>${usedCats.map(c => `<button class="fchip ${cat === c.name ? 'on' : ''}" style="--c:${catColor(c.name)}" data-paycat="${esc(c.name)}"><span class="sw" style="--c:${catColor(c.name)}"></span>${esc(c.name)}</button>`).join('')}</div>
      <div class="sortbar"><span class="num">${list.length} payment${list.length === 1 ? '' : 's'} · ${inr(sum)}</span>${sortControl('payments')}</div>
    </div>
    ${list.length ? groups : `<div class="list"><div class="empty">${d.payments.length ? 'No payments match.' : 'No payments yet. Tap Add Payment.'}</div></div>`}
  </div>`;
}

function vendorsView() {
  const stats = sortVendors(vendorStats().filter(v => v.name), S.sort.vendors);
  return `<div class="stack">
    <div class="sortbar"><span>${stats.length} vendor${stats.length === 1 ? '' : 's'}</span>${sortControl('vendors')}</div>
    <div class="list">${stats.map(v => vendorRow(v, totalSpent())).join('') || '<div class="empty">No vendors yet.</div>'}</div>
  </div>`;
}

function vendorDetailView(name) {
  const s = vendorStats().find(v => v.name === name) || { name, total: 0, count: 0, byCat: {} };
  const v = D().vendors.find(x => x.name === name);
  const pays = sortPayments(D().payments.filter(p => p.vendor === name), S.sort.vendorPays);
  const docs = allDocuments().filter(x => x.owner === name || (x.ref.type === 'payment' && pays.some(p => p.id === x.ref.id)));
  const c = vendorColor(name);
  return `<div class="stack">
    <div style="display:flex;gap:14px;align-items:center">
      <span class="av" style="--c:${c};width:56px;height:56px;font-size:19px">${esc(initials(name))}</span>
      <div style="min-width:0"><div style="font-size:22px;font-weight:800">${esc(name)}</div>
      <div class="small muted">${esc([v?.trade, v?.phone].filter(Boolean).join(' · ') || 'No trade or phone saved')}</div></div>
    </div>
    <button class="btn wide" style="background:${c}" data-act="add-payment" data-vendor="${esc(name)}">＋ Add payment to ${esc(name.split(' ')[0])}</button>
    <div class="grid2 grid4">
      <div class="tile"><div class="lbl">Total paid</div><div class="val num">${inr(s.total)}</div></div>
      <div class="tile"><div class="lbl">Payments</div><div class="val num">${s.count}</div></div>
      <div class="tile"><div class="lbl">First</div><div class="val" style="font-size:17px">${shortDate(s.first)}</div></div>
      <div class="tile"><div class="lbl">Latest</div><div class="val" style="font-size:17px">${shortDate(s.last)}</div></div>
    </div>
    ${Object.keys(s.byCat).length ? `<div class="card"><div class="sec-h" style="margin:0 0 10px"><h2 style="font-size:16px">By category</h2></div>
      ${Object.entries(s.byCat).sort((a, b) => b[1] - a[1]).map(([cn, amt]) => `<div style="margin-bottom:10px"><div style="display:flex;justify-content:space-between;margin-bottom:5px"><span class="chip" style="--c:${catColor(cn)}">${esc(cn)}</span><b class="num small">${inr(amt)}</b></div><div class="meter" style="height:6px"><i style="width:${amt / Math.max(s.total, 1) * 100}%;--c:${catColor(cn)}"></i></div></div>`).join('')}</div>` : ''}
    <div class="sec-h"><h2 style="font-size:16px">Documents</h2></div>
    <div class="list">${docs.map(docRow).join('') || '<div class="empty">No bills, receipts or quotations yet.</div>'}</div>
    <div class="sec-h"><h2 style="font-size:16px">Payments</h2>${sortControl('vendorPays')}</div>
    <div class="list">${pays.map(paymentRow).join('') || '<div class="empty">No payments yet.</div>'}</div>
  </div>`;
}

function budgetView() {
  const d = D(), budget = d.totalBudget, planned = totalBudgeted(), gap = budget - planned;
  return `<div class="stack">
    <div class="grid2 grid4">
      <button class="tile" style="text-align:left" data-act="edit-budget"><div class="lbl">🎯 Total budget ✎</div><div class="val num">${lakh(budget)}</div><div class="sub">Tap to change</div></button>
      <div class="tile"><div class="lbl">📋 Planned</div><div class="val num">${lakh(planned)}</div><div class="sub">category estimates</div></div>
      <div class="tile"><div class="lbl">${gap < 0 ? '⚠️ Plan over by' : '✅ Not planned'}</div><div class="val num" style="color:${gap < 0 ? 'var(--bad)' : 'var(--good)'}">${lakh(Math.abs(gap))}</div></div>
      <div class="tile"><div class="lbl">₹ Spent</div><div class="val num">${lakh(totalSpent())}</div></div>
    </div>
    <div class="sortbar"><h2 style="font-size:18px;color:var(--ink)">Categories</h2>${sortControl('categories')}</div>
    ${sortCategories(S.sort.categories).map(({ c }) => categoryCard(c)).join('') || '<div class="list"><div class="empty">No categories yet.</div></div>'}
  </div>`;
}
function categoryCard(c) {
  const b = budgeted(c), spent = spentIn(c.name), rem = b - spent, frac = b > 0 ? spent / b : (spent > 0 ? 1 : 0), over = b > 0 && spent > b, col = catColor(c.name);
  const n = (c.attachments || []).length;
  return `<button class="card" style="text-align:left;width:100%" data-act="edit-category" data-id="${esc(c.id)}">
    <div style="display:flex;gap:12px;align-items:center">
      <span class="ic" style="--c:${col}">${catIcon(c.name)}</span>
      <div style="flex:1;min-width:0"><div class="t" style="font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(c.name)}</div>
        <div class="small faint" style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(c.scope || (c.estimateType || 'Rough') + ' estimate')}</div></div>
      ${b === 0 && spent > 0 ? '<span class="pill warn">No budget</span>' : over ? '<span class="pill bad">Over</span>' : `<b style="color:${col}">${b > 0 ? Math.round(frac * 100) + '%' : '—'}</b>`}
    </div>
    <div class="meter" style="margin:12px 0"><i style="width:${Math.min(100, frac * 100)}%;--c:${over ? 'var(--bad)' : col}"></i></div>
    <div style="display:flex;justify-content:space-between">
      <div><div class="small faint">Spent</div><b class="num">${inr(spent)}</b></div>
      <div style="text-align:right"><div class="small faint">${rem < 0 ? 'Over by' : 'Left of ' + lakh(b)}</div><b class="num" style="color:${rem < 0 ? 'var(--bad)' : 'inherit'}">${inr(Math.abs(rem))}</b></div>
    </div>
    ${n ? `<div class="small" style="margin-top:8px;color:var(--accent);font-weight:700">📎 ${n} quotation${n === 1 ? '' : 's'} attached</div>` : ''}
  </button>`;
}

function docRow(x) {
  const col = KIND_COLOR[x.a.kind] || '#64748B';
  return `<button class="row" data-act="view-doc" data-file="${esc(x.a.file)}" data-name="${esc(x.a.name)}">
    <span class="ic" style="--c:${col}">${docIcon(x.a.file)}</span>
    <span class="main"><span class="t">${esc(x.a.name)}</span><span class="s">${esc(x.owner)} · ${esc(x.detail)}</span>
      <span style="display:flex;gap:6px;align-items:center"><span class="chip" style="--c:${col}">${esc(x.a.kind)}</span><span class="small faint">${shortDate(x.date)}</span></span></span>
    <span style="color:var(--accent)">👁</span>
  </button>`;
}
function documentsView() {
  const all = allDocuments(), q = key(S.filters.docSearch), kind = S.filters.docKind;
  const docs = sortDocs(all.filter(x => (!kind || x.a.kind === kind) && (!q || [x.a.name, x.owner, x.detail].join(' ').toLowerCase().includes(q))), S.sort.documents);
  return `<div class="stack">
    <div class="tools">
      <div class="search">🔍<input type="search" id="docSearch" placeholder="Search documents" value="${esc(S.filters.docSearch)}"></div>
      <div class="chips"><button class="fchip ${!kind ? 'on' : ''}" data-dockind="">All · ${all.length}</button>${KINDS.map(k => `<button class="fchip ${kind === k ? 'on' : ''}" style="--c:${KIND_COLOR[k]}" data-dockind="${k}">${k}s · ${all.filter(x => x.a.kind === k).length}</button>`).join('')}</div>
      <div class="sortbar"><span>${docs.length} document${docs.length === 1 ? '' : 's'}</span>${sortControl('documents')}</div>
    </div>
    <div class="list">${docs.map(docRow).join('') || `<div class="empty">${all.length ? 'No documents match.' : 'No documents yet. Open a payment and add its bill or receipt.'}</div>`}</div>
  </div>`;
}

/* ---------------- sheets (forms) ---------------- */
let sheetCleanup = null;
function openSheet(title, bodyHTML, { saveLabel = 'Save', onSave, onCancel, noSave } = {}) {
  closeSheet(true);
  const layer = $('#layer');
  layer.innerHTML = `<div class="scrim" id="scrim"><div class="sheet" role="dialog" aria-modal="true">
    <div class="sheet-h"><button data-sheet="cancel">${noSave ? '' : 'Cancel'}</button><h3>${esc(title)}</h3><button class="save" data-sheet="save">${noSave ? 'Done' : esc(saveLabel)}</button></div>
    <div class="sheet-b" id="sheetBody">${bodyHTML}</div></div></div>`;
  document.body.style.overflow = 'hidden';
  const done = async (save) => {
    if (save && onSave) { const btn = $('[data-sheet="save"]'); btn.disabled = true; const ok = await onSave(); btn.disabled = false; if (ok === false) return; }
    else if (!save && onCancel) onCancel();
    closeSheet();
  };
  $('#scrim').addEventListener('click', e => {
    if (e.target.id === 'scrim') done(false);
    const b = e.target.closest('[data-sheet]'); if (!b) return;
    done(b.dataset.sheet === 'save' ? !noSave : (noSave ? false : false));
  });
  sheetCleanup = () => {};
}
function closeSheet(silent) { $('#layer').innerHTML = ''; document.body.style.overflow = ''; sheetCleanup = null; }

/* Attachments editor used inside forms */
function docsEditor(listId, defaultKind) {
  return `<div class="group"><h4>Bills, receipts & quotations</h4><div id="${listId}"></div>
    <div class="adddoc"><select id="${listId}-kind">${KINDS.map(k => `<option ${k === defaultKind ? 'selected' : ''}>${k}</option>`).join('')}</select>
    <label class="btn sm sec" style="position:relative">📷 Add photo or file<input type="file" id="${listId}-file" accept="image/*,application/pdf,.pdf,.heic" multiple style="position:absolute;inset:0;opacity:0"></label></div>
    <div class="small faint" id="${listId}-status" style="margin-top:6px"></div></div>`;
}
function wireDocs(listId, ctx, ownerFn, labelFn) {
  const draw = () => {
    $('#' + listId).innerHTML = ctx.atts.length ? ctx.atts.map((a, i) => `<div class="doc"><span class="ic" style="--c:${KIND_COLOR[a.kind] || '#64748B'};width:32px;height:32px;font-size:15px">${docIcon(a.file)}</span>
      <button class="n" style="border:0;background:none;text-align:left;padding:0" data-act="view-doc" data-file="${esc(a.file)}" data-name="${esc(a.name)}"><b>${esc(a.name)}</b><span class="small faint">${esc(a.kind)}</span></button>
      <button type="button" class="x" data-rm="${i}" aria-label="Remove">✕</button></div>`).join('') : '<div class="small faint">None attached yet.</div>';
  };
  draw();
  $('#' + listId).addEventListener('click', e => {
    const b = e.target.closest('[data-rm]'); if (!b) return;
    const a = ctx.atts[+b.dataset.rm];
    if (!confirm(`Remove “${a.name}”? It is kept in ${locShort()} › Attachments › Removed.`)) return;
    ctx.atts.splice(+b.dataset.rm, 1); ctx.removed.push(a); draw();
  });
  $('#' + listId + '-file').addEventListener('change', async e => {
    const files = [...e.target.files]; e.target.value = '';
    const st = $('#' + listId + '-status'), kind = $('#' + listId + '-kind').value;
    ctx.uploading = (ctx.uploading || 0) + 1;
    try {
      for (let i = 0; i < files.length; i++) {
        st.innerHTML = `<span class="spin" style="width:12px;height:12px"></span> Uploading ${i + 1} of ${files.length} to OneDrive…`;
        const a = await uploadAttachment(files[i], ownerFn(), labelFn(), kind);
        ctx.atts.push(a); ctx.added.push(a); draw();
      }
      st.textContent = 'Uploaded. Tap Save to keep it.';
    } catch (err) { st.textContent = 'Upload failed: ' + err.message; }
    ctx.uploading--;
  });
}
const docsCtx = existing => ({ atts: [...(existing || [])], added: [], removed: [], uploading: 0 });
function finishDocs(ctx, saved) {
  // Saved: archive removed ones. Cancelled: archive the ones uploaded for nothing.
  for (const a of (saved ? ctx.removed : ctx.added)) moveToRemoved(a);
}
const opt = (v, cur) => `<option ${v === cur ? 'selected' : ''}>${esc(v)}</option>`;

function paymentForm(p, presetVendor) {
  const d = D(), isNew = !p;
  p = p || { id: uuid(), date: todayStr(), category: S.filters.payCat || '', item: '', vendor: presetVendor || '', mode: 'UPI', amount: '', invoice: '', status: 'Paid', notes: '' };
  const ctx = docsCtx(p.attachments);
  const html = `
    <div class="field"><label>Amount (₹)</label><input class="amount-in num" id="f-amount" inputmode="decimal" placeholder="0" value="${p.amount === '' ? '' : esc(p.amount)}" autocomplete="off">
      <span class="hint">You can type a sum, e.g. 55090-1690</span></div>
    <div class="two"><div class="field"><label>Date</label><input type="date" id="f-date" value="${esc(p.date)}"></div>
      <div class="field"><label>Paid by</label><select id="f-mode">${[...new Set([...MODES, p.mode].filter(Boolean))].map(m => opt(m, p.mode)).join('')}</select></div></div>
    <div class="field"><label>Vendor</label><input id="f-vendor" list="dl-vendors" value="${esc(p.vendor)}" placeholder="Who was paid" autocomplete="off">
      <datalist id="dl-vendors">${vendorStats().map(v => `<option value="${esc(v.name)}">`).join('')}</datalist></div>
    <div class="field"><label>Category</label><select id="f-cat"><option value="">Choose…</option>${d.categories.map(c => opt(c.name, p.category)).join('')}${p.category && !d.categories.some(c => c.name === p.category) ? opt(p.category, p.category) : ''}<option value="__new">＋ New category…</option></select>
      <input id="f-catnew" placeholder="New category name" style="display:none;margin-top:6px"></div>
    <div class="field"><label>Item / work</label><input id="f-item" list="dl-items" value="${esc(p.item)}" placeholder="e.g. Labor, Cement, Tiles" autocomplete="off">
      <datalist id="dl-items">${d.items.map(i => `<option value="${esc(i)}">`).join('')}</datalist></div>
    <div class="field"><label>Status</label><div class="seg-in" id="f-status">${['Paid', 'Pending'].map(s => `<button type="button" class="${(p.status || 'Paid') === s ? 'on' : ''}" data-v="${s}">${s}</button>`).join('')}</div></div>
    <div class="field"><label>Invoice / bill no.</label><input id="f-invoice" value="${esc(p.invoice)}"></div>
    <div class="field"><label>Notes</label><textarea id="f-notes">${esc(p.notes)}</textarea></div>
    ${docsEditor('payDocs', 'Bill')}
    ${isNew ? '' : `<button class="btn danger wide" type="button" id="f-delete">Delete payment</button>`}`;
  openSheet(isNew ? 'New Payment' : 'Edit Payment', html, {
    onSave: async () => {
      if (ctx.uploading) { toast('Please wait for the upload to finish.'); return false; }
      const amount = parseAmount($('#f-amount').value);
      if (!isFinite(amount) || amount <= 0) { toast('Please enter the amount.'); $('#f-amount').focus(); return false; }
      let category = $('#f-cat').value; if (category === '__new') category = $('#f-catnew').value.trim();
      if (!category) { toast('Please choose a category.'); return false; }
      const vendor = $('#f-vendor').value.trim(), itemName = $('#f-item').value.trim();
      const fields = { date: $('#f-date').value || todayStr(), amount, mode: $('#f-mode').value, vendor, category, item: itemName,
        status: $('#f-status .on')?.dataset.v || 'Paid', invoice: $('#f-invoice').value.trim(), notes: $('#f-notes').value.trim() };
      const atts = ctx.atts.slice(), id = p.id;
      const ok = await mutate(data => {
        rememberNames(data, vendor, category, itemName);
        let x = data.payments.find(q => q.id === id);
        if (!x) { x = { id, createdAt: isoNow() }; data.payments.push(x); }
        Object.assign(x, fields, { updatedAt: isoNow() });
        if (atts.length) x.attachments = atts; else delete x.attachments;
      });
      if (ok) { finishDocs(ctx, true); toast(isNew ? 'Payment added' : 'Payment saved'); }
      return ok;
    },
    onCancel: () => finishDocs(ctx, false),
  });
  wireDocs('payDocs', ctx, () => $('#f-vendor').value || $('#f-cat').value, () => $('#f-item').value);
  $('#f-cat').addEventListener('change', e => { $('#f-catnew').style.display = e.target.value === '__new' ? '' : 'none'; if (e.target.value === '__new') $('#f-catnew').focus(); });
  $('#f-status').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; $('#f-status').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); });
  $('#f-delete')?.addEventListener('click', async () => {
    if (!confirm(`Delete this payment of ${inr(p.amount)} to ${p.vendor || 'vendor'}? Its bills are moved to the Removed folder in OneDrive.`)) return;
    const ok = await mutate(data => { data.payments = data.payments.filter(q => q.id !== p.id); });
    if (ok) { for (const a of p.attachments || []) moveToRemoved(a); finishDocs(ctx, false); closeSheet(); toast('Payment deleted'); }
  });
  if (isNew) setTimeout(() => $('#f-amount')?.focus(), 250);
}
/** Remembers new vendors, categories and items so they appear in the lists next time. */
function rememberNames(data, vendor, category, itemName) {
  if (vendor && !data.vendors.some(v => key(v.name) === key(vendor))) {
    let id = slug(vendor); while (data.vendors.some(v => v.id === id)) id += '-2';
    data.vendors.push({ id, name: vendor, phone: '', trade: '' });
  }
  if (category && !data.categories.some(c => key(c.name) === key(category))) {
    let id = slug(category); while (data.categories.some(c => c.id === id)) id += '-2';
    data.categories.push({ id, name: category, scope: '', quotedBy: '', estimateType: 'Rough', quoted: 0, contingency: 0.1, notes: '', colorIndex: data.categories.length % PAL.length });
  }
  if (itemName && !data.items.some(i => key(i) === key(itemName))) data.items.push(itemName);
}

function vendorForm(v) {
  const isNew = !v;
  v = v || { id: null, name: '', trade: '', phone: '' };
  const ctx = docsCtx(v.attachments), oldName = v.name;
  const used = !isNew && D().payments.some(p => p.vendor === oldName);
  openSheet(isNew ? 'New Vendor' : 'Edit Vendor', `
    <div class="field"><label>Name</label><input id="v-name" value="${esc(v.name)}" autocomplete="off"></div>
    <div class="field"><label>Trade</label><input id="v-trade" value="${esc(v.trade)}" placeholder="e.g. Electrician"></div>
    <div class="field"><label>Phone</label><input id="v-phone" type="tel" value="${esc(v.phone)}"></div>
    ${docsEditor('venDocs', 'Quotation')}
    ${isNew ? '' : used ? `<p class="small faint">This vendor has payments, so it can't be deleted. Renaming updates all their payments.</p>` : `<button class="btn danger wide" type="button" id="v-delete">Delete vendor</button>`}`, {
    onSave: async () => {
      if (ctx.uploading) { toast('Please wait for the upload to finish.'); return false; }
      const name = $('#v-name').value.trim(); if (!name) { toast('Please enter a name.'); return false; }
      if (key(name) !== key(oldName) && D().vendors.some(x => key(x.name) === key(name))) { toast('A vendor with that name already exists.'); return false; }
      const trade = $('#v-trade').value.trim(), phone = $('#v-phone').value.trim(), atts = ctx.atts.slice();
      const ok = await mutate(data => {
        let x = isNew ? null : data.vendors.find(q => q.id === v.id) || data.vendors.find(q => q.name === oldName);
        if (!x) { let id = slug(name); while (data.vendors.some(q => q.id === id)) id += '-2'; x = { id }; data.vendors.push(x); }
        Object.assign(x, { name, trade, phone });
        if (atts.length) x.attachments = atts; else delete x.attachments;
        if (!isNew && oldName && oldName !== name) for (const p of data.payments) if (p.vendor === oldName) { p.vendor = name; p.updatedAt = isoNow(); }
      });
      if (ok) { finishDocs(ctx, true); if (S.vendor === oldName) S.vendor = name; render(); }
      return ok;
    },
    onCancel: () => finishDocs(ctx, false),
  });
  wireDocs('venDocs', ctx, () => $('#v-name').value, () => 'Vendor');
  $('#v-delete')?.addEventListener('click', async () => {
    if (!confirm(`Delete vendor “${oldName}”?`)) return;
    const ok = await mutate(data => { data.vendors = data.vendors.filter(q => q.id !== v.id); });
    if (ok) { for (const a of v.attachments || []) moveToRemoved(a); closeSheet(); S.vendor = null; render(); }
  });
}

function categoryForm(c) {
  const isNew = !c, d = D();
  c = c || { id: null, name: '', scope: '', quotedBy: '', estimateType: 'Rough', quoted: 0, contingency: 0.1, notes: '', colorIndex: d.categories.length % PAL.length };
  const ctx = docsCtx(c.attachments), oldName = c.name, used = !isNew && d.payments.some(p => p.category === oldName);
  let color = c.colorIndex ?? d.categories.indexOf(c);
  openSheet(isNew ? 'New Category' : 'Edit Category', `
    <div class="field"><label>Name</label><input id="c-name" value="${esc(c.name)}"></div>
    <div class="field"><label>What it covers</label><input id="c-scope" value="${esc(c.scope)}" placeholder="e.g. Mason, cement, sand"></div>
    <div class="two"><div class="field"><label>Estimate (₹)</label><input id="c-quoted" inputmode="decimal" value="${c.quoted ? esc(c.quoted) : ''}" placeholder="0"></div>
      <div class="field"><label>Extra buffer %</label><input id="c-cont" inputmode="decimal" value="${Math.round((+c.contingency || 0) * 100)}"></div></div>
    <div class="two"><div class="field"><label>Estimate type</label><select id="c-type">${[...new Set([...EST_TYPES, c.estimateType].filter(Boolean))].map(t => opt(t, c.estimateType)).join('')}</select></div>
      <div class="field"><label>Quoted by</label><input id="c-by" value="${esc(c.quotedBy)}"></div></div>
    <div class="field"><label>Budget incl. buffer</label><div id="c-total" class="num" style="font-weight:700;font-size:18px"></div></div>
    <div class="field"><label>Colour</label><div class="palette" id="c-pal">${PAL.map((p, i) => `<button type="button" style="--c:${p}" data-i="${i}" class="${i === color % PAL.length ? 'on' : ''}"></button>`).join('')}</div></div>
    <div class="field"><label>Notes</label><textarea id="c-notes">${esc(c.notes)}</textarea></div>
    ${docsEditor('catDocs', 'Quotation')}
    ${isNew ? '' : used ? `<p class="small faint">This category has payments, so it can't be deleted. Renaming updates all its payments.</p>` : `<button class="btn danger wide" type="button" id="c-delete">Delete category</button>`}`, {
    onSave: async () => {
      if (ctx.uploading) { toast('Please wait for the upload to finish.'); return false; }
      const name = $('#c-name').value.trim(); if (!name) { toast('Please enter a name.'); return false; }
      if (key(name) !== key(oldName) && D().categories.some(x => key(x.name) === key(name))) { toast('A category with that name already exists.'); return false; }
      const quoted = $('#c-quoted').value.trim() ? parseAmount($('#c-quoted').value) : 0;
      if (!isFinite(quoted)) { toast('Please check the estimate.'); return false; }
      const contingency = (parseFloat($('#c-cont').value) || 0) / 100;
      const fields = { name, scope: $('#c-scope').value.trim(), quoted, contingency, estimateType: $('#c-type').value, quotedBy: $('#c-by').value.trim(), notes: $('#c-notes').value.trim(), colorIndex: color };
      const atts = ctx.atts.slice();
      const ok = await mutate(data => {
        let x = isNew ? null : data.categories.find(q => q.id === c.id) || data.categories.find(q => q.name === oldName);
        if (!x) { let id = slug(name); while (data.categories.some(q => q.id === id)) id += '-2'; x = { id }; data.categories.push(x); }
        Object.assign(x, fields);
        if (atts.length) x.attachments = atts; else delete x.attachments;
        if (!isNew && oldName && oldName !== name) for (const p of data.payments) if (p.category === oldName) { p.category = name; p.updatedAt = isoNow(); }
      });
      if (ok) finishDocs(ctx, true);
      return ok;
    },
    onCancel: () => finishDocs(ctx, false),
  });
  const upd = () => { const q = parseAmount($('#c-quoted').value) || 0, k = (parseFloat($('#c-cont').value) || 0) / 100; $('#c-total').textContent = inr(q * (1 + k)); };
  $('#c-quoted').addEventListener('input', upd); $('#c-cont').addEventListener('input', upd); upd();
  $('#c-pal').addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return; color = +b.dataset.i; $('#c-pal').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); });
  wireDocs('catDocs', ctx, () => $('#c-name').value, () => 'Quotation');
  $('#c-delete')?.addEventListener('click', async () => {
    if (!confirm(`Delete category “${oldName}”?`)) return;
    const ok = await mutate(data => { data.categories = data.categories.filter(q => q.id !== c.id); });
    if (ok) { for (const a of c.attachments || []) moveToRemoved(a); closeSheet(); }
  });
}

function budgetForm() {
  openSheet('Total Budget', `<div class="field"><label>Your whole renovation budget (₹)</label>
    <input class="amount-in num" id="b-total" inputmode="decimal" value="${D().totalBudget || ''}" placeholder="e.g. 4000000"></div>
    <p class="small faint">Currently ${lakh(D().totalBudget)}.</p>`, {
    onSave: async () => {
      const v = parseAmount($('#b-total').value);
      if (!isFinite(v) || v < 0) { toast('Please enter the budget.'); return false; }
      return mutate(data => { data.totalBudget = v; });
    },
  });
}

function accountSheet() {
  const d = D();
  openSheet('Account & Data', `
    <div class="group"><h4>Microsoft account</h4>
      <div style="display:flex;gap:12px;align-items:center"><span class="av" style="--c:var(--accent)">${esc(initials(S.account || 'M'))}</span>
      <div style="min-width:0"><b style="display:block;overflow:hidden;text-overflow:ellipsis">${esc(S.account || 'Signed in')}</b><span class="small" style="color:var(--good)">✓ Signed in to OneDrive</span></div></div>
      <button class="btn danger wide" style="margin-top:12px" id="a-signout">Sign out</button></div>
    <div class="group"><h4>OneDrive data</h4>
      <div class="kv"><span>Folder</span><span>${esc(locFull())}</span></div>
      <div class="kv"><span>Last synced</span><span>${S.sync.at ? timeStr(S.sync.at) : 'Not yet'}</span></div>
      <div class="kv"><span>Daily backups</span><span>${esc(locShort())} › Backups</span></div>
      <button class="btn sec wide" style="margin-top:10px" id="a-folder">📁 Change folder…</button>
      <button class="btn sec wide" style="margin-top:10px" id="a-sync">↻ Sync now</button>
      <button class="btn sec wide" style="margin-top:8px" id="a-export">⬇︎ Export to Excel</button></div>
    <div class="group"><h4>Your data</h4>
      <div class="kv"><span>Payments</span><span>${d.payments.length} · ${inr(totalSpent())}</span></div>
      <div class="kv"><span>Vendors</span><span>${d.vendors.length}</span></div>
      <div class="kv"><span>Categories</span><span>${d.categories.length}</span></div>
      <div class="kv"><span>Documents</span><span>${allDocuments().length}</span></div></div>
    <div class="group"><h4>Appearance</h4>
      <div class="seg-in" id="a-theme">${[['auto', 'Auto'], ['light', '☀️ Light'], ['dark', '🌙 Dark']].map(([v, t]) => `<button type="button" data-theme-v="${v}" class="${(ls.get('ra.theme') || 'auto') === v ? 'on' : ''}">${t}</button>`).join('')}</div>
      <div class="small faint" style="margin-top:8px">Auto follows your iPhone or Mac setting.</div></div>
    <div class="group"><h4>Security</h4>
      <button class="btn sec wide" id="a-lock">🔒 Passcode & Face ID / Touch ID · ${getLock() ? 'On' : 'Off'}</button></div>
    <div class="group"><h4>Add to your Home Screen</h4>
      <ol class="steps"><li><b>iPhone:</b> open this page in Safari → tap Share → <b>Add to Home Screen</b>.</li>
      <li><b>Mac:</b> download the Mac app from the install page below.</li></ol>
      <a class="btn sec wide" style="margin-top:10px;text-decoration:none" href="download/" target="_blank" rel="noopener">Install page — share this link</a>
      <div class="small faint" style="margin-top:6px;text-align:center">dijango999.github.io/reno-app/download</div></div>
    <p class="small faint" style="text-align:center">Reno App ${CFG.version}</p>`, { noSave: true });
  $('#a-signout').onclick = () => { if (confirm('Sign out? Your data stays safe in OneDrive.')) signOut(); };
  $('#a-folder').onclick = () => changeFolderSheet();
  $('#a-sync').onclick = async () => { await pull(true); toast('Up to date'); closeSheet(); };
  $('#a-export').onclick = exportExcel;
  $('#a-lock').onclick = () => lockSettingsSheet();
  $('#a-theme').onclick = e => { const b = e.target.closest('[data-theme-v]'); if (!b) return; setTheme(b.dataset.themeV);
    $('#a-theme').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); };
}

function setTheme(v) {
  if (v === 'light' || v === 'dark') { document.documentElement.dataset.theme = v; ls.set('ra.theme', v); }
  else { delete document.documentElement.dataset.theme; ls.del('ra.theme'); }
}

/* ---------------- document viewer ---------------- */
async function viewDoc(file, name) {
  const v = document.createElement('div'); v.className = 'viewer';
  v.innerHTML = `<div class="vh"><button id="v-close">✕ Close</button><b>${esc(name)}</b><a id="v-open" href="#" target="_blank" rel="noopener" style="display:none">Open</a></div><div class="vb"><span class="spin" style="color:#fff"></span></div>`;
  document.body.appendChild(v);
  v.querySelector('#v-close').onclick = () => v.remove();
  try {
    const meta = await (await g(item('Attachments/' + file))).json();
    const url = meta['@microsoft.graph.downloadUrl'];
    const open = v.querySelector('#v-open'); open.href = url; open.style.display = '';
    const body = v.querySelector('.vb');
    if (isImageName(file)) body.innerHTML = `<img src="${esc(url)}" alt="${esc(name)}">`;
    else if (/\.pdf$/i.test(file) && window.matchMedia('(min-width:700px)').matches) {
      const blob = await (await fetch(url)).blob();
      body.innerHTML = `<iframe src="${URL.createObjectURL(new Blob([blob], { type: 'application/pdf' }))}"></iframe>`;
    } else body.innerHTML = `<div style="color:#fff;text-align:center"><div style="font-size:48px">📄</div><p>${esc(name)}</p><a class="btn" href="${esc(url)}" target="_blank" rel="noopener">Open document</a></div>`;
  } catch (e) {
    v.querySelector('.vb').innerHTML = `<p style="color:#fff">${e.status === 404 ? 'This file is not in ' + esc(locFull()) + ' › Attachments.' : esc(e.message)}</p>`;
  }
}

/* ---------------- Excel export ---------------- */
function loadScript(src) { return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('Could not load the Excel tool. Check your connection.')); document.head.appendChild(s); }); }
async function exportExcel() {
  try {
    toast('Preparing Excel file…');
    if (!window.XLSX) await loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');
    const d = D(), X = window.XLSX, wb = X.utils.book_new();
    const add = (name, rows, widths) => { const ws = X.utils.aoa_to_sheet(rows); ws['!cols'] = widths.map(w => ({ wch: w })); X.utils.book_append_sheet(wb, ws, name); };
    const spent = totalSpent();
    add('Summary', [['Reno App — exported', new Date().toLocaleString('en-IN')], [], ['Total budget', d.totalBudget], ['Spent so far', spent], ['Cash left', d.totalBudget - spent],
      ['Planned (category estimates)', totalBudgeted()], ['Payments', d.payments.length], ['Pending amount', d.payments.filter(p => p.status !== 'Paid').reduce((s, p) => s + (+p.amount || 0), 0)]], [34, 22]);
    add('Payments', [['Date', 'Vendor', 'Category', 'Item', 'Amount', 'Mode', 'Status', 'Invoice', 'Notes', 'Documents'],
      ...sortPayments(d.payments, 'oldest').map(p => [p.date, p.vendor, p.category, p.item, +p.amount || 0, p.mode, p.status, p.invoice, p.notes, (p.attachments || []).map(a => a.file).join('; ')])],
      [12, 26, 22, 26, 14, 14, 10, 14, 34, 50]);
    add('Vendors', [['Vendor', 'Trade', 'Phone', 'Total paid', 'Payments', 'First', 'Latest'],
      ...sortVendors(vendorStats(), 'spentHigh').map(v => [v.name, v.trade, v.phone, v.total, v.count, v.first || '', v.last || ''])], [30, 18, 16, 14, 10, 12, 12]);
    add('Categories', [['Category', 'Covers', 'Estimate type', 'Quoted', 'Buffer %', 'Budget', 'Spent', 'Left', '% used'],
      ...d.categories.map(c => { const b = budgeted(c), s = spentIn(c.name); return [c.name, c.scope, c.estimateType, +c.quoted || 0, Math.round((+c.contingency || 0) * 100), Math.round(b), s, Math.round(b - s), b > 0 ? Math.round(s / b * 100) : '']; })],
      [24, 40, 14, 14, 10, 14, 14, 14, 9]);
    const cats = d.categories.map(c => c.name), vs = sortVendors(vendorStats().filter(v => v.total > 0), 'name');
    add('Vendor x Category', [['Vendor', ...cats, 'Total'], ...vs.map(v => [v.name, ...cats.map(c => v.byCat[c] || 0), v.total])], [28, ...cats.map(() => 14), 14]);
    X.writeFile(wb, `Reno App ${todayStr()}.xlsx`);
  } catch (e) { toast(e.message); }
}

/* ---------------- import from the Mac app's folder ---------------- */
async function importFrom(fileList) {
  const st = $('#importStatus'); const say = h => { if (st) st.innerHTML = h; };
  const files = [...fileList].filter(f => !f.name.startsWith('.'));
  // Path inside the chosen folder, e.g. "Attachments/x.jpg" (or just the name when files were picked one by one)
  const rel = f => { const parts = (f.webkitRelativePath || f.name).split('/'); return f.webkitRelativePath ? parts.slice(1).join('/') : parts.join('/'); };
  const trackerFile = files.find(f => rel(f) === 'tracker.json');
  if (!trackerFile) { say('<span style="color:var(--bad)">No tracker.json found. Choose the “Renovation Tracker” folder itself.</span>'); return; }
  let text, data;
  try { text = await trackerFile.text(); data = normalize(JSON.parse(text)); if (!Array.isArray(data.payments)) throw 0; }
  catch { say('<span style="color:var(--bad)">That tracker.json could not be read.</span>'); return; }
  try {
    if (await loadRemote()) { say('OneDrive already has data — nothing was copied.'); await pull(true); return; }
    const referenced = new Set([...data.payments, ...data.vendors, ...data.categories].flatMap(x => (x.attachments || []).map(a => a.file)));
    const uploads = [];
    for (const f of files) {
      const r = rel(f);
      if (r === 'tracker.json') continue;
      if (/^(Attachments|Backups)\//.test(r)) uploads.push([r, f]);
      else if (!f.webkitRelativePath && referenced.has(f.name)) uploads.push(['Attachments/' + f.name, f]);
    }
    const have = new Set(uploads.map(u => u[0].replace(/^Attachments\//, '')));
    const missing = [...referenced].filter(n => !have.has(n));
    for (let i = 0; i < uploads.length; i++) {
      say(`<span class="spin" style="width:12px;height:12px"></span> Copying ${i + 1} of ${uploads.length}: ${esc(uploads[i][0])}`);
      await uploadBlob(uploads[i][0], uploads[i][1]);
    }
    say('<span class="spin" style="width:12px;height:12px"></span> Saving your data…');
    const etag = await putTracker(data, null, true);   // create only — never overwrites existing data
    ls.set('ra.backup', '');
    applyData(data, etag); setSync('ok');
    toast(`Imported ${data.payments.length} payments and ${uploads.filter(u => u[0].startsWith('Attachments/')).length} documents` + (missing.length ? ` (${missing.length} bill${missing.length === 1 ? '' : 's'} not found)` : ''));
  } catch (e) { say(`<span style="color:var(--bad)">Import stopped: ${esc(e.message)}. Nothing was overwritten; you can try again.</span>`); }
}
async function startEmpty() {
  if (!confirm('Start with empty data? Only do this if you have no existing data to bring in.')) return;
  try {
    if (await loadRemote()) { await pull(true); return; }
    const data = { totalBudget: 0, categories: [], vendors: [], payments: [], items: [], savedAt: isoNow() };
    applyData(data, await putTracker(data, null, true));
  } catch (e) { toast(e.message); }
}

/* ---------------- toast ---------------- */
let toastTimer;
function toast(msg) {
  document.querySelector('.toast')?.remove();
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t);
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.remove(), 3200);
}

/* ---------------- events ---------------- */
document.addEventListener('click', e => {
  const t = e.target.closest('[data-tab],[data-act],[data-paycat],[data-dockind]');
  if (!t || t.closest('#scrim') && !t.dataset.act) return;
  if (t.dataset.tab) { S.tab = t.dataset.tab; S.vendor = null; ls.set('ra.tab', S.tab); closeSheet(); render(); window.scrollTo(0, 0); return; }
  if ('paycat' in t.dataset) { S.filters.payCat = t.dataset.paycat; render(); return; }
  if ('dockind' in t.dataset) { S.filters.docKind = t.dataset.dockind; render(); return; }
  const id = t.dataset.id;
  switch (t.dataset.act) {
    case 'signin': signIn(false); break;
    case 'signin-other': signIn(true); break;
    case 'start-empty': startEmpty(); break;
    case 'change-folder': changeFolderSheet(); break;
    case 'account': accountSheet(); break;
    case 'export': exportExcel(); break;
    case 'add-payment': paymentForm(null, t.dataset.vendor); break;
    case 'edit-payment': { const p = D().payments.find(x => x.id === id); if (p) paymentForm(structuredClone(p)); break; }
    case 'open-vendor': S.vendor = t.dataset.name; render(); window.scrollTo(0, 0); break;
    case 'back': S.vendor = null; render(); break;
    case 'new-vendor': vendorForm(null); break;
    case 'edit-vendor': { const n = t.dataset.name; vendorForm(structuredClone(D().vendors.find(v => v.name === n) || { id: null, name: n, trade: '', phone: '' })); break; }
    case 'new-category': categoryForm(null); break;
    case 'edit-category': { const c = D().categories.find(x => x.id === id); if (c) categoryForm(structuredClone(c)); break; }
    case 'edit-budget': budgetForm(); break;
    case 'view-doc': viewDoc(t.dataset.file, t.dataset.name); break;
  }
});
document.addEventListener('change', e => {
  const s = e.target.closest('[data-sort]');
  if (s) { S.sort[s.dataset.sort] = s.value; ls.set('sort.' + s.dataset.sort, s.value); render(); return; }
  if (e.target.id === 'importDir' || e.target.id === 'importFiles') importFrom(e.target.files);
});
let searchTimer;
document.addEventListener('input', e => {
  if (e.target.id !== 'paySearch' && e.target.id !== 'docSearch') return;
  S.filters[e.target.id] = e.target.value;
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => {
    const id = e.target.id, pos = e.target.selectionStart; render();
    const el = $('#' + id); if (el) { el.focus(); try { el.setSelectionRange(pos, pos); } catch {} }
  }, 250);
});
window.addEventListener('scroll', () => $('#top')?.classList.toggle('scrolled', window.scrollY > 4), { passive: true });
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.phase === 'ready') pull(false); });
window.addEventListener('online', () => S.phase === 'ready' && pull(false));
setInterval(() => { if (document.visibilityState === 'visible' && S.phase === 'ready' && !$('#scrim')) pull(false); }, 20000);

/** Renews the sign-in in the background (at most once every 2 minutes, so it can never loop). */
function trySilent() {
  const last = +(ls.get('ra.silentAt') || 0);
  if (!S.account || Date.now() - last < 120000 || !navigator.onLine) return false;
  S.phase = 'loading'; render(); signIn(false, true); return true;
}

/* ---------------- app lock: passcode + Face ID / Touch ID ----------------
 * A privacy lock for this device. The passcode is stored only as a salted PBKDF2 hash.
 * Face ID / Touch ID uses the phone's own biometric check (WebAuthn platform authenticator);
 * the fingerprint or face never leaves the device. Forgot the passcode? Sign out and sign in
 * with Microsoft again — that clears the lock and the copy kept on this device.
 */
const LOCK_KEY = 'ra.lock';
const getLock = () => { try { return JSON.parse(ls.get(LOCK_KEY) || 'null'); } catch { return null; } };
const setLock = v => v ? ls.set(LOCK_KEY, JSON.stringify(v)) : ls.del(LOCK_KEY);
const fromB64url = s => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4)), c => c.charCodeAt(0));
async function hashPin(pin, saltB64) {
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(pin), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt: fromB64url(saltB64), iterations: 150000 }, key, 256);
  return b64url(bits);
}
async function bioAvailable() {
  try { return !!(window.PublicKeyCredential && await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable()); } catch { return false; }
}
const bioName = () => /iPhone|iPad/.test(navigator.userAgent) ? 'Face ID' : /Macintosh/.test(navigator.userAgent) ? 'Touch ID' : 'fingerprint / face';
async function bioRegister() {
  const rnd = n => crypto.getRandomValues(new Uint8Array(n));
  const cred = await navigator.credentials.create({ publicKey: {
    challenge: rnd(32), rp: { name: 'Reno App', id: location.hostname },
    user: { id: rnd(16), name: S.account || 'Reno App', displayName: 'Reno App lock' },
    pubKeyCredParams: [{ type: 'public-key', alg: -7 }, { type: 'public-key', alg: -257 }],
    authenticatorSelection: { authenticatorAttachment: 'platform', userVerification: 'required', residentKey: 'discouraged' },
    timeout: 60000 } });
  return b64url(cred.rawId);
}
async function bioCheck(id) {
  const a = await navigator.credentials.get({ publicKey: {
    challenge: crypto.getRandomValues(new Uint8Array(32)),
    allowCredentials: [{ type: 'public-key', id: fromB64url(id), transports: ['internal', 'hybrid'] }],
    userVerification: 'required', timeout: 60000 } });
  const flags = new Uint8Array(a.response.authenticatorData)[32];
  return !!(flags & 0x04);                         // the device confirmed it was really you
}

let locked = false, hiddenAt = 0, pinBuf = '', failCount = +(ls.get('ra.lockFails') || 0), lockUntil = +(ls.get('ra.lockUntil') || 0);
function showLock() {
  const L = getLock(); if (!L || locked) return;
  locked = true; pinBuf = '';
  document.documentElement.classList.add('is-locked');
  const el = document.createElement('div'); el.id = 'lock'; el.className = 'lock';
  el.innerHTML = `<img class="logo" src="icons/icon-192.png" alt="">
    <h2>Reno App is locked</h2><p class="muted" id="lockMsg">Enter your passcode</p>
    <div class="pindots" id="pinDots">${'<i></i>'.repeat(L.len)}</div>
    <div class="keypad">${[1, 2, 3, 4, 5, 6, 7, 8, 9].map(n => `<button data-k="${n}">${n}</button>`).join('')}
      ${L.bio ? `<button data-k="bio" aria-label="Unlock with ${bioName()}">${bioName() === 'Face ID' ? '🙂' : '👆'}</button>` : '<span></span>'}
      <button data-k="0">0</button><button data-k="del" aria-label="Delete">⌫</button></div>
    ${L.bio ? `<button class="linkbtn" data-k="bio">Unlock with ${bioName()}</button>` : ''}
    <button class="linkbtn" data-k="forgot" style="color:var(--muted)">Forgot passcode?</button>`;
  document.body.appendChild(el);
  el.addEventListener('click', onLockKey);
  if (L.bio) setTimeout(() => tryBio(), 350);
}
function hideLock() {
  locked = false; document.getElementById('lock')?.remove(); document.documentElement.classList.remove('is-locked');
  failCount = 0; ls.set('ra.lockFails', '0');
}
async function tryBio() {
  const L = getLock(); if (!L?.bio || !locked) return;
  try { if (await bioCheck(L.bio)) hideLock(); }
  catch { const m = $('#lockMsg'); if (m) m.textContent = 'Enter your passcode'; }
}
function drawDots(err) {
  const d = $('#pinDots'); if (!d) return;
  d.querySelectorAll('i').forEach((x, i) => x.classList.toggle('on', i < pinBuf.length));
  d.classList.toggle('shake', !!err);
}
async function onLockKey(e) {
  const b = e.target.closest('[data-k]'); if (!b) return;
  const k = b.dataset.k, L = getLock(), msg = $('#lockMsg');
  if (k === 'bio') return tryBio();
  if (k === 'forgot') {
    if (confirm('Sign out of Reno App on this device? Your data stays safe in OneDrive. After signing in with Microsoft again you can set a new passcode.')) {
      setLock(null); hideLock(); signOut();
    }
    return;
  }
  if (Date.now() < lockUntil) { msg.textContent = `Too many tries. Wait ${Math.ceil((lockUntil - Date.now()) / 1000)} seconds.`; return; }
  if (k === 'del') { pinBuf = pinBuf.slice(0, -1); drawDots(); return; }
  if (pinBuf.length >= L.len) return;
  pinBuf += k; drawDots();
  if (pinBuf.length < L.len) return;
  const ok = (await hashPin(pinBuf, L.salt)) === L.hash;
  if (ok) { hideLock(); return; }
  failCount++; ls.set('ra.lockFails', String(failCount));
  if (failCount >= 5) { lockUntil = Date.now() + 30000 * Math.pow(2, Math.min(failCount - 5, 5)); ls.set('ra.lockUntil', String(lockUntil)); }
  msg.textContent = failCount >= 5 ? 'Too many tries. Please wait a moment.' : 'Wrong passcode. Try again.';
  drawDots(true); setTimeout(() => { pinBuf = ''; drawDots(); }, 450);
}
document.addEventListener('visibilitychange', () => {
  const L = getLock(); if (!L || S.phase === 'welcome' || S.phase === 'needsignin') return;
  if (document.visibilityState === 'hidden') { hiddenAt = Date.now(); if (!L.after) showLock(); }   // hide content in the app switcher
  else if (Date.now() - hiddenAt >= (L.after || 0) * 1000) showLock();
});

async function lockSettingsSheet() {
  const L = getLock(), canBio = await bioAvailable();
  openSheet('Passcode & ' + bioName(), `
    <div class="group"><h4>App lock</h4>
      <div class="kv"><span>Status</span><span>${L ? '🔒 On' : 'Off'}</span></div>
      ${L ? `<div class="kv"><span>${esc(bioName())}</span><span>${L.bio ? 'On' : 'Off'}</span></div>` : ''}
      <p class="small faint" style="margin:8px 0 0">The lock is for this ${/iPhone|iPad/.test(navigator.userAgent) ? 'iPhone' : 'device'} only. Each phone sets its own passcode.</p></div>
    <div class="group"><h4>${L ? 'Change passcode' : 'Turn on with a passcode'}</h4>
      <div class="field"><label>New passcode (4–6 digits)</label><input id="pin1" type="password" inputmode="numeric" autocomplete="new-password" maxlength="6"></div>
      <div class="field"><label>Type it again</label><input id="pin2" type="password" inputmode="numeric" autocomplete="new-password" maxlength="6"></div>
      ${canBio ? `<label style="display:flex;gap:10px;align-items:center;font-weight:600"><input type="checkbox" id="pinBio" ${!L || L.bio ? 'checked' : ''} style="width:22px;height:22px"> Also unlock with ${esc(bioName())}</label>` : `<p class="small faint">${esc(bioName())} isn't available in this browser; the passcode will be used.</p>`}
      <button class="btn wide" id="pinSave" style="margin-top:12px">${L ? 'Save new passcode' : 'Turn on lock'}</button>
      <div class="small" id="pinMsg" style="margin-top:8px;color:var(--bad)"></div></div>
    ${L ? `<div class="group"><h4>Lock again after</h4><div class="seg-in" id="lockAfter">${[[0, 'Right away'], [60, '1 min'], [300, '5 min']].map(([v, t]) => `<button type="button" data-v="${v}" class="${(L.after || 0) === v ? 'on' : ''}">${t}</button>`).join('')}</div></div>
      <button class="btn danger wide" id="lockOff">Turn off lock</button>` : ''}`, { noSave: true });
  $('#pinSave').onclick = async () => {
    const a = $('#pin1').value.trim(), b = $('#pin2').value.trim(), m = $('#pinMsg');
    if (!/^\d{4,6}$/.test(a)) { m.textContent = 'Use 4 to 6 digits.'; return; }
    if (a !== b) { m.textContent = 'The two passcodes are different.'; return; }
    const salt = randStr(22);
    const next = { len: a.length, salt, hash: await hashPin(a, salt), after: L?.after ?? 0, bio: null };
    if ($('#pinBio')?.checked) {
      try { next.bio = L?.bio || await bioRegister(); }
      catch { m.textContent = `${bioName()} wasn't set up (cancelled). The passcode still works.`; }
    }
    setLock(next); toast(L ? 'Passcode changed' : 'App lock is on'); if (!m.textContent) closeSheet(); else lockSettingsSheet();
  };
  $('#lockAfter')?.addEventListener('click', e => { const b = e.target.closest('button'); if (!b) return;
    const cur = getLock(); cur.after = +b.dataset.v; setLock(cur);
    $('#lockAfter').querySelectorAll('button').forEach(x => x.classList.toggle('on', x === b)); });
  $('#lockOff')?.addEventListener('click', () => { if (confirm('Turn off the app lock on this device?')) { setLock(null); toast('App lock is off'); closeSheet(); } });
}

/* ---------------- start ---------------- */
(async function start() {
  if (getLock() && tokens) showLock();
  render();
  try { await handleRedirect(); } catch (e) { toast(e.message); }
  if (!tokens) { if (S.account && trySilent()) return; S.phase = S.account ? 'needsignin' : 'welcome'; render(); return; }
  // Show the last synced copy straight away, then refresh from OneDrive.
  try { const c = JSON.parse(ls.get('ra.cache') || 'null'); if (c) { S.data = normalize(c); S.etag = ls.get('ra.etag') || null; S.phase = 'ready'; render(); } } catch {}
  try { await accessToken(); } catch (e) { if (e instanceof AuthError) { if (trySilent()) return; S.phase = 'needsignin'; render(); return; } }
  try {
    if (!(await resolveLocation())) { S.phase = 'choose'; render(); return; }
  } catch (e) { if (!S.data) { S.phase = 'ready'; } setSync('bad', e.status === 0 ? 'Offline — showing the last synced copy' : e.message); }
  await pull(true);
  if (S.phase === 'loading') { S.phase = 'ready'; if (!S.data) S.phase = 'needsignin'; render(); }
})();
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});
