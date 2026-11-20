/* My Gold — family gold record, kept in OneDrive.
 *
 * Data lives in a OneDrive folder you choose (e.g. "My Gold"):
 *   my-gold.json        every piece of gold, family members, settings and zakat history
 *   Photos/             photos of each piece
 *   Bills/              bills and receipts
 *   Backups/            one copy of my-gold.json per day, taken before the first change of the day
 *
 * Every change is applied to the newest copy from OneDrive and saved with an eTag check, so two
 * phones editing at the same time never overwrite each other. Sign-in is the same Microsoft
 * sign-in as Reno App (PKCE, no secrets in the code).
 *
 * Gold prices (no key needed, free):
 *   live spot  – api.gold-api.com (USD/oz), currencies from fawazahmed0 currency-api
 *   any date   – fawazahmed0 currency-api (from Mar 2024), freegoldapi.com for older dates (USD)
 */
'use strict';

const CFG = {
  clientId: 'c010a16c-4e83-493b-b778-86ad96b766e9',
  authority: 'https://login.microsoftonline.com/consumers/oauth2/v2.0',
  scopes: 'Files.ReadWrite User.Read offline_access',
  graph: 'https://graph.microsoft.com/v1.0',
  dataFile: 'my-gold.json',
  pointer: 'Apps/My Gold/location.json',
  version: '1.0',
};
const CANON = 'https://dijango999.github.io/reno-app/my-gold/';
const ON_SITE = location.hostname === 'dijango999.github.io';
if (ON_SITE && location.origin + location.pathname !== CANON) location.replace(CANON + location.search);
const REDIRECT = ON_SITE ? CANON : location.origin + location.pathname.replace(/index\.html$/, '');

const OZ = 31.1034768;                 // grams in a troy ounce
const NISAB_G = 85;                    // nisab in grams of pure gold
const HAWL_DAYS = 354;                 // one lunar year
const KARATS = [
  { k: 24, p: 0.999, label: '24K' }, { k: 22, p: 0.916, label: '22K' }, { k: 21, p: 0.875, label: '21K' },
  { k: 18, p: 0.750, label: '18K' }, { k: 14, p: 0.585, label: '14K' },
];
const TYPES = {
  coin: { label: 'Coin', plural: 'Coins', icon: '🪙', color: '#D4A017' },
  bar: { label: 'Bar / Biscuit', plural: 'Bars', icon: '🧈', color: '#B45309' },
  jewellery: { label: 'Jewellery', plural: 'Jewellery', icon: '💍', color: '#DB2777' },
};
const CURS = { AED: { sym: 'AED ', loc: 'en-AE' }, INR: { sym: '₹', loc: 'en-IN' }, USD: { sym: '$', loc: 'en-US' } };
const OWNER_COLORS = ['#C99A2E', '#DB2777', '#7C3AED', '#0EA5E9', '#16A34A', '#EA580C'];
const AED_PER_USD = 3.6725;            // AED has been pegged to the dollar since 1997

/* ---------------- helpers ---------------- */
const $ = (s, el = document) => el.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ls = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch {} },
  del(k) { try { localStorage.removeItem(k); } catch {} },
};
const lsJSON = (k, d = null) => { try { return JSON.parse(ls.get(k) || 'null') ?? d; } catch { return d; } };
const uuid = () => (crypto.randomUUID ? crypto.randomUUID() : 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => {
  const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }));
const isoNow = () => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
const ymd = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const todayStr = () => ymd(new Date());
const parseDay = s => { const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s || ''); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const addDays = (s, n) => { const d = parseDay(s); d.setDate(d.getDate() + n); return ymd(d); };
const daysBetween = (a, b) => Math.round((parseDay(b) - parseDay(a)) / 864e5);
const shortDate = s => { const d = parseDay(s); return d ? d.toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '–'; };
const timeStr = d => d ? new Date(d).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '';
const cmpText = (a, b) => String(a).localeCompare(String(b), undefined, { sensitivity: 'base', numeric: true });
const num = v => { const n = parseFloat(String(v ?? '').replace(/[,\s₹]/g, '').replace(/^AED/i, '')); return isFinite(n) ? n : NaN; };
const round2 = n => Math.round(n * 100) / 100;
const trimZ = s => s.includes('.') ? s.replace(/\.?0+$/, '') : s;
const grams = (n, dp = 3) => trimZ((+n || 0).toFixed(dp)) + ' g';
function money(n, cur = S.data?.settings?.currency || 'AED', dp) {
  if (n == null || !isFinite(n)) return '–';
  const c = CURS[cur] || { sym: cur + ' ', loc: 'en-US' };
  const a = Math.abs(n), d = dp ?? (a < 1000 ? 2 : 0);
  return (n < 0 ? '-' : '') + c.sym + a.toLocaleString(c.loc, { minimumFractionDigits: d, maximumFractionDigits: d });
}
function moneyK(n, c = cur()) {
  if (n == null || !isFinite(n)) return '–';
  const a = Math.abs(n), sym = (CURS[c] || { sym: c + ' ' }).sym, sg = n < 0 ? '-' : '';
  if (c === 'INR' && a >= 1e7) return sg + sym + trimZ((a / 1e7).toFixed(2)) + ' Cr';
  if (c === 'INR' && a >= 1e5) return sg + sym + trimZ((a / 1e5).toFixed(2)) + ' L';
  if (a >= 1e6) return sg + sym + trimZ((a / 1e6).toFixed(2)) + 'M';
  if (a >= 1e4) return sg + sym + trimZ((a / 1e3).toFixed(1)) + 'K';
  return money(n, c, 0);
}
const pct = n => (n > 0 ? '+' : '') + (isFinite(n) ? n.toFixed(1) : '0') + '%';
const initials = n => String(n).split(/\s+/).filter(Boolean).slice(0, 2).map(w => w[0]).join('').toUpperCase() || '?';
const isImageName = n => /\.(jpe?g|png|gif|webp|heic|heif|bmp)$/i.test(n || '');

/* ---------------- state ---------------- */
const S = {
  data: null, etag: null, account: ls.get('mg.account'), demo: ls.get('mg.demo') === '1',
  phase: 'loading',                 // loading | welcome | needsignin | folder | setup | ready
  tab: ls.get('mg.tab') || 'home',
  f: { owner: '', type: '', q: '', sort: ls.get('mg.sort') || 'newest', sold: false },
  sync: { state: 'idle', at: null, msg: '' },
  rates: lsJSON('mg.rates'), ratesBusy: false, ratesErr: '',
};
let tokens = lsJSON('mg.tokens');

/* ---------------- Microsoft sign-in (PKCE, no secrets) ---------------- */
class AuthError extends Error {}
class HttpError extends Error { constructor(status, msg) { super(msg); this.status = status; } }
const b64url = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const randStr = n => { const a = new Uint8Array(n); crypto.getRandomValues(a); return b64url(a).slice(0, n); };

async function signIn(chooseAccount, silent) {
  const verifier = randStr(64);
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = randStr(24);
  ls.set('mg.pkce', JSON.stringify({ verifier, state }));
  const p = new URLSearchParams({
    client_id: CFG.clientId, response_type: 'code', redirect_uri: REDIRECT, scope: CFG.scopes,
    code_challenge: challenge, code_challenge_method: 'S256', state, response_mode: 'query',
  });
  if (S.account && !chooseAccount) p.set('login_hint', S.account); else p.set('prompt', 'select_account');
  if (silent) { p.set('prompt', 'none'); ls.set('mg.silentAt', String(Date.now())); }
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
  const saved = lsJSON('mg.pkce'); ls.del('mg.pkce');
  if (!saved || saved.state !== q.get('state')) { toast('That sign-in expired. Please try again.'); return; }
  await tokenRequest({ grant_type: 'authorization_code', code: q.get('code'), redirect_uri: REDIRECT, code_verifier: saved.verifier });
  try {
    const me = await (await g('/me?$select=userPrincipalName,displayName,mail')).json();
    S.account = me.userPrincipalName || me.mail || me.displayName || 'Microsoft account';
    ls.set('mg.account', S.account);
  } catch {}
}
async function tokenRequest(fields) {
  const body = new URLSearchParams({ client_id: CFG.clientId, scope: CFG.scopes, ...fields });
  let r; try { r = await fetch(`${CFG.authority}/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }); }
  catch { throw new HttpError(0, 'No internet connection.'); }
  const j = await r.json().catch(() => ({}));
  if (!r.ok || !j.access_token) throw new AuthError(j.error_description || 'Microsoft sign-in failed.');
  tokens = { access: j.access_token, refresh: j.refresh_token || tokens?.refresh, exp: Date.now() + ((j.expires_in || 3600) - 120) * 1000 };
  ls.set('mg.tokens', JSON.stringify(tokens));
}
let refreshing = null;
async function accessToken() {
  if (tokens?.access && tokens.exp > Date.now()) return tokens.access;
  if (!tokens?.refresh) throw new AuthError('Please sign in.');
  refreshing ||= tokenRequest({ grant_type: 'refresh_token', refresh_token: tokens.refresh }).finally(() => { refreshing = null; });
  await refreshing; return tokens.access;
}
function trySilent() {
  const last = +(ls.get('mg.silentAt') || 0);
  if (!S.account || Date.now() - last < 120000 || !navigator.onLine) return false;
  S.phase = 'loading'; render(); signIn(false, true); return true;
}
function signOut() {
  for (const k of ['mg.account', 'mg.tokens', 'mg.cache', 'mg.etag', 'mg.loc', 'mg.demo', 'mg.demoData']) ls.del(k);
  tokens = null; LOC = null; S.account = null; S.demo = false; S.data = null; S.etag = null;
  S.phase = 'welcome'; closeSheet(); render();
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
const DEFAULT_PATH = 'Documents/02 ACCOUNTS/My Gold app';
let LOC = lsJSON('mg.loc');
const encRel = rel => rel.split('/').filter(Boolean).map(encodeURIComponent).join('/');
const item = rel => { if (!LOC) throw new Error('Choose where to keep your data first.'); return `/drives/${LOC.driveId}/items/${LOC.id}:/${encRel(rel)}`; };
const locFull = (l = LOC) => l ? 'OneDrive › ' + (l.path ? l.path.split('/').join(' › ') : 'My files') : 'not chosen';
function folderPath(meta) {
  const parent = decodeURIComponent((meta.parentReference?.path || '').replace(/^\/drive(s\/[^/]+)?\/root:?\/?/, ''));
  return meta.root ? '' : [parent, meta.name].filter(Boolean).join('/');
}
function saveLoc(l) {
  const changed = !LOC || LOC.id !== l.id;
  LOC = { driveId: l.driveId, id: l.id, path: l.path || '' }; ls.set('mg.loc', JSON.stringify(LOC));
  if (changed) { ls.del('mg.cache'); ls.del('mg.etag'); S.data = null; S.etag = null; }
}
async function readPointer() {
  try {
    const m = await (await g(`/me/drive/root:/${encRel(CFG.pointer)}`)).json();
    const p = JSON.parse(await (await fetch(m['@microsoft.graph.downloadUrl'], { cache: 'no-store' })).text());
    return p && p.id && p.driveId ? p : null;
  } catch (e) { if (e.status === 404 || e instanceof SyntaxError) return null; throw e; }
}
async function writePointer(l) {
  await g(`/me/drive/root:/${encRel(CFG.pointer)}:/content`, { method: 'PUT', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ driveId: l.driveId, id: l.id, path: l.path, savedAt: isoNow(), by: S.account || '' }, null, 2) });
}
async function resolveLocation() {
  const p = await readPointer();
  if (p) {
    try { const m = await (await g(`/drives/${p.driveId}/items/${p.id}?$select=id,name,root,parentReference`)).json();
      saveLoc({ driveId: p.driveId, id: p.id, path: folderPath(m) }); return true; }
    catch (e) { if (e.status !== 404) throw e; }
  }
  if (LOC) { await writePointer(LOC).catch(() => {}); return true; }
  try {   // the family's chosen folder: Documents › 02 ACCOUNTS › My Gold app
    const f = await (await g(`/me/drive/root:/${encRel(DEFAULT_PATH)}?$select=id,name,parentReference`)).json();
    const l = { driveId: f.parentReference.driveId, id: f.id, path: DEFAULT_PATH };
    saveLoc(l); await writePointer(l).catch(() => {}); return true;
  } catch (e) { if (e.status !== 404) throw e; }
  return false;
}
async function loadRemote() {
  let meta;
  try { meta = await (await g(item(CFG.dataFile))).json(); }
  catch (e) { if (e.status === 404) return null; throw e; }
  const r = await fetch(meta['@microsoft.graph.downloadUrl'], { cache: 'no-store' });
  if (!r.ok) throw new HttpError(r.status, 'Could not download your data.');
  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { throw new Error(`${CFG.dataFile} in OneDrive could not be read. Nothing was changed.`); }
  return { data: normalize(data), etag: meta.eTag, text };
}
function normalize(d) {
  d.owners ||= []; d.items ||= []; d.zakatLog ||= [];
  d.settings = { currency: 'AED', theme: 'auto', zakatMethod: 'majority', zakatDate: '', premium: { AED: 0, INR: 0, USD: 0 }, ...(d.settings || {}) };
  d.settings.premium = { AED: 0, INR: 0, USD: 0, ...(d.settings.premium || {}) };
  for (const it of d.items) { it.photos ||= []; it.docs ||= []; it.status ||= 'held'; it.zakat ||= 'auto'; }
  return d;
}
const stableJSON = v => JSON.stringify(v, null, 2);
async function putData(data, etag, createOnly) {
  const headers = { 'Content-Type': 'application/json' };
  if (createOnly) headers['If-None-Match'] = '*'; else if (etag) headers['If-Match'] = etag;
  const r = await g(item(CFG.dataFile) + ':/content', { method: 'PUT', headers, body: stableJSON(data) });
  return (await r.json()).eTag;
}
async function dailyBackup(text) {
  const day = todayStr();
  if (ls.get('mg.backup') === day) return;
  try {
    await g(item(`Backups/my-gold-${day}.json`) + ':/content?@microsoft.graph.conflictBehavior=fail',
      { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: text });
  } catch (e) { if (e.status !== 409) { console.warn('backup', e); return; } }
  ls.set('mg.backup', day);
}

let saving = Promise.resolve();
/** Applies `change` to the newest data and saves it. Retries if someone else saved first. */
function mutate(change) {
  const run = async () => {
    if (S.demo) {
      const d = normalize(JSON.parse(JSON.stringify(S.data))); change(d); d.savedAt = isoNow();
      ls.set('mg.demoData', JSON.stringify(d)); applyData(d, null); return true;
    }
    if (!navigator.onLine) throw new Error('You are offline. Connect to the internet to save changes.');
    setSync('busy', 'Saving…');
    for (let attempt = 0; attempt < 5; attempt++) {
      const remote = await loadRemote();
      if (!remote) throw new Error('Your data file is missing in OneDrive. Nothing was saved.');
      await dailyBackup(remote.text);
      const d = remote.data; change(d); d.savedAt = isoNow();
      try { const etag = await putData(d, remote.etag); applyData(d, etag); setSync('ok'); return true; }
      catch (e) { if (e.status === 412 || e.status === 409) continue; throw e; }
    }
    throw new Error('OneDrive is busy. Please try again.');
  };
  const p = saving.then(run, run);
  saving = p.catch(() => {});
  return p.catch(e => { setSync('bad', e.message); toast(e.message); return false; });
}
function applyData(d, etag) {
  S.data = d; S.etag = etag;
  if (!S.demo) { ls.set('mg.cache', JSON.stringify(d)); ls.set('mg.etag', etag || ''); }
  applyTheme();
  if (S.phase !== 'ready') S.phase = 'ready';
  render();
}
async function pull(force) {
  if (S.demo || !tokens) return;
  try {
    if (!force && S.etag) {
      const meta = await (await g(item(CFG.dataFile) + '?$select=eTag')).json();
      if (meta.eTag === S.etag) { setSync('ok'); return; }
    }
    setSync('busy', 'Syncing…');
    const remote = await loadRemote();
    if (!remote) { S.phase = 'setup'; render(); return; }
    applyData(remote.data, remote.etag); setSync('ok');
  } catch (e) {
    if (e instanceof AuthError) { if (trySilent()) return; S.phase = 'needsignin'; render(); return; }
    setSync('bad', e.status === 0 ? 'Offline — showing the last synced copy' : e.message);
  }
}
function setSync(state, msg = '') {
  S.sync = { state, msg, at: state === 'ok' ? new Date() : S.sync.at };
  const el = $('#syncline'); if (el) el.innerHTML = syncLine();
}
function syncLine() {
  if (S.demo) return `<span class="dot" style="background:var(--info)"></span>Sample data on this device only — nothing is saved to OneDrive`;
  const s = S.sync;
  if (s.state === 'busy') return `<span class="dot busy"></span>${esc(s.msg)}`;
  if (s.state === 'bad') return `<span class="dot bad"></span>${esc(s.msg)}`;
  return `<span class="dot"></span>Saved in ${esc(locFull())}${s.at ? ' · ' + timeStr(s.at) : ''}`;
}

/* ---------------- photos and bills ---------------- */
async function loadBitmap(file) {
  try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch {
    const url = URL.createObjectURL(file);
    try { const img = new Image(); img.src = url; await img.decode(); return img; } finally { setTimeout(() => URL.revokeObjectURL(url), 2000); }
  }
}
async function resizeImage(file, max, q) {
  const bmp = await loadBitmap(file);
  const w = bmp.width || bmp.naturalWidth, h = bmp.height || bmp.naturalHeight;
  const sc = Math.min(1, max / Math.max(w, h));
  const c = document.createElement('canvas'); c.width = Math.round(w * sc); c.height = Math.round(h * sc);
  c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
  return { blob: await new Promise(r => c.toBlob(r, 'image/jpeg', q)), dataUrl: c.toDataURL('image/jpeg', q) };
}
/** Square-cropped small preview kept inside my-gold.json so lists show photos instantly, even offline. */
async function makeThumb(file) {
  const bmp = await loadBitmap(file);
  const w = bmp.width || bmp.naturalWidth, h = bmp.height || bmp.naturalHeight, s = Math.min(w, h), T = 220;
  const c = document.createElement('canvas'); c.width = c.height = T;
  c.getContext('2d').drawImage(bmp, (w - s) / 2, (h - s) / 2, s, s, 0, 0, T, T);
  return c.toDataURL('image/jpeg', 0.72);
}
async function uploadBlob(rel, blob) {
  if (blob.size <= 4 * 1024 * 1024) {
    await g(item(rel) + ':/content?@microsoft.graph.conflictBehavior=replace',
      { method: 'PUT', headers: { 'Content-Type': blob.type || 'application/octet-stream' }, body: blob });
    return;
  }
  const s = await (await g(item(rel) + ':/createUploadSession', { method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'replace' } }) })).json();
  const chunk = 320 * 1024 * 16;
  for (let start = 0; start < blob.size; start += chunk) {
    const end = Math.min(start + chunk, blob.size);
    const r = await fetch(s.uploadUrl, { method: 'PUT', headers: { 'Content-Range': `bytes ${start}-${end - 1}/${blob.size}` }, body: blob.slice(start, end) });
    if (!r.ok) throw new HttpError(r.status, 'Upload failed. Please try again.');
  }
}
const cleanName = s => String(s || '').replace(/[\\/:*?"<>|#%{}~&]+/g, ' ').replace(/\s+/g, ' ').trim();
function fileName(ext, owner, title, kind) {
  const code = Math.random().toString(16).slice(2, 6).toUpperCase();
  return [todayStr(), cleanName(owner), cleanName(title), kind, code].filter(Boolean).join(' ').slice(0, 140) + '.' + ext;
}
/** Photo → Photos/ in OneDrive (max 2000px) plus a small preview for the app. */
async function savePhoto(file, owner, title) {
  const big = await resizeImage(file, 2000, 0.86);
  const thumb = await makeThumb(file);
  const name = fileName('jpg', owner, title, 'Photo');
  if (S.demo) return { id: uuid(), file: name, thumb, local: (await resizeImage(file, 1200, 0.8)).dataUrl, addedAt: isoNow() };
  await uploadBlob('Photos/' + name, big.blob);
  return { id: uuid(), file: name, thumb, addedAt: isoNow() };
}
/** Bill or receipt (photo or PDF) → Bills/ in OneDrive. */
async function saveDoc(file, owner, title, kind) {
  let blob = file, ext = (/\.(\w{1,5})$/.exec(file.name || '')?.[1] || (file.type.split('/')[1] || 'bin')).toLowerCase(), thumb = '';
  if (/^image\//.test(file.type)) {
    try { blob = (await resizeImage(file, 2400, 0.85)).blob; ext = 'jpg'; thumb = await makeThumb(file); } catch {}
  }
  const name = fileName(ext, owner, title, kind);
  if (S.demo) return { id: uuid(), file: name, name: file.name || name, kind, thumb, addedAt: isoNow() };
  await uploadBlob('Bills/' + name, blob);
  return { id: uuid(), file: name, name: file.name || name, kind, thumb, addedAt: isoNow() };
}
async function downloadUrl(rel) {
  const m = await (await g(item(rel) + '?$select=id,@microsoft.graph.downloadUrl')).json();
  return m['@microsoft.graph.downloadUrl'];
}
async function showPhoto(p) {
  const v = document.createElement('div'); v.className = 'viewer';
  v.innerHTML = `<img alt="" src="${p.local || p.thumb}"><button aria-label="Close">✕</button>`;
  v.addEventListener('click', () => v.remove()); document.body.appendChild(v);
  if (!p.local && !S.demo) {
    try { const url = await downloadUrl('Photos/' + p.file); const img = $('img', v); if (img) img.src = url; }
    catch (e) { toast('Showing the preview — full photo needs internet.'); }
  }
}
async function openDoc(d) {
  if (S.demo) { toast('Sample data — files are only kept when signed in to OneDrive.'); return; }
  const w = window.open('', '_blank');
  try { const url = await downloadUrl('Bills/' + d.file); if (w) w.location = url; else location.href = url; }
  catch (e) { w?.close(); toast(e.message); }
}
async function removeFile(folder, f) {
  if (S.demo) return;
  try {
    try { await g(item(folder) + ':/children', { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: 'Removed', folder: {}, '@microsoft.graph.conflictBehavior': 'fail' }) }); }
    catch (e) { if (e.status !== 409) throw e; }
    const removed = await (await g(item(folder + '/Removed') + '?$select=id')).json();
    await g(item(folder + '/' + f.file), { method: 'PATCH', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ parentReference: { id: removed.id } }) });
  } catch (e) { console.warn('remove file', e); }
}

/* ---------------- gold prices ---------------- */
async function fetchJSON(url, ms = 9000) {
  const c = new AbortController(); const t = setTimeout(() => c.abort(), ms);
  try { const r = await fetch(url, { signal: c.signal, cache: 'no-store' }); if (!r.ok) throw new Error('HTTP ' + r.status); return await r.json(); }
  finally { clearTimeout(t); }
}
/** Gold price per troy ounce in USD, AED and INR for one day (fawazahmed0 currency-api, from Mar 2024). */
async function xauOn(day) {
  for (const u of [`https://cdn.jsdelivr.net/npm/@fawazahmed0/currency-api@${day}/v1/currencies/xau.min.json`,
                   `https://${day}.currency-api.pages.dev/v1/currencies/xau.min.json`]) {
    try { const j = await fetchJSON(u, 7000); const x = j.xau; if (x?.usd) return { date: j.date, usd: x.usd, aed: x.aed || x.usd * AED_PER_USD, inr: x.inr || null, src: 'currency-api' }; } catch {}
  }
  return null;
}
let longHistory = null;
async function usdHistoryOn(day) {          // freegoldapi.com — long USD history, used for dates before Mar 2024
  longHistory ||= fetchJSON('https://freegoldapi.com/data/latest.json', 20000).catch(e => { longHistory = null; throw e; });
  const rows = await longHistory; let best = null;
  for (const r of rows) { if (r.date <= day) best = r; else break; }
  return best ? { date: best.date, usd: +best.price, aed: +best.price * AED_PER_USD, inr: null, src: 'freegoldapi' } : null;
}
/** Market gold price on a past day. Kept forever in this browser and in the item itself. */
async function rateOn(day) {
  const cache = lsJSON('mg.hist', {});
  if (cache[day]) return cache[day];
  let r = null;
  if (day >= '2024-03-06') for (let i = 0; i < 4 && !r; i++) r = await xauOn(addDays(day, -i));
  if (!r) { try { r = await usdHistoryOn(day); } catch {} }
  if (r && day < todayStr()) { cache[day] = r; ls.set('mg.hist', JSON.stringify(cache)); }
  return r;
}
/** Live price: spot from gold-api.com, converted with today's currency rates. */
async function refreshRates(force) {
  if (S.ratesBusy) return;
  if (!force && S.rates && Date.now() - S.rates.at < 4 * 60 * 1000) return;
  S.ratesBusy = true; S.ratesErr = ''; paintRates();
  try {
    const [spot, base] = await Promise.all([
      fetchJSON('https://api.gold-api.com/price/XAU', 8000).catch(() => null),
      (async () => { for (let i = 0; i < 3; i++) { const r = await xauOn(addDays(todayStr(), -i)); if (r) return r; } return xauOn('latest'); })(),
    ]);
    let r = null;
    if (spot?.price && base) r = { usd: +spot.price, aed: +spot.price * (base.aed / base.usd), inr: base.inr ? +spot.price * (base.inr / base.usd) : null, at: Date.now(), quoteAt: spot.updatedAt || null, src: 'gold-api.com · live' };
    else if (spot?.price) r = { usd: +spot.price, aed: +spot.price * AED_PER_USD, inr: S.rates?.inr && S.rates?.usd ? +spot.price * (S.rates.inr / S.rates.usd) : null, at: Date.now(), quoteAt: spot.updatedAt || null, src: 'gold-api.com · live' };
    else if (base) r = { usd: base.usd, aed: base.aed, inr: base.inr, at: Date.now(), quoteAt: base.date, src: 'currency-api · daily' };
    if (!r) throw new Error('Could not get the gold price. Showing the last known price.');
    S.rates = r; ls.set('mg.rates', JSON.stringify(r));
  } catch (e) { S.ratesErr = e.message; }
  S.ratesBusy = false;
  if (S.phase === 'ready') render();
}

/* ---------------- calculations ---------------- */
const owners = () => S.data?.owners || [];
const ownerById = id => owners().find(o => o.id === id);
const purityOf = it => +it.purity || (KARATS.find(k => k.k === +it.karat)?.p) || (+it.karat ? +it.karat / 24 : 0.999);
const karatLabel = it => it.karat ? `${it.karat}K` : `${trimZ((purityOf(it) * 1000).toFixed(1))}`;
const netWeight = it => Math.max(0, (+it.weight || 0) - (it.type === 'jewellery' ? (+it.stoneWeight || 0) : 0));
const pureGrams = it => netWeight(it) * purityOf(it);
const cur = () => S.data?.settings?.currency || 'AED';
/** Price of one gram of pure gold now, in currency c (with the local adjustment from Settings). */
function pureGramNow(c = cur()) {
  const r = S.rates; if (!r) return null;
  const perOz = c === 'USD' ? r.usd : c === 'AED' ? r.aed : r.inr;
  if (!perOz) return null;
  return perOz / OZ * (1 + (+S.data?.settings?.premium?.[c] || 0) / 100);
}
const gramRate = (purity, c = cur()) => { const p = pureGramNow(c); return p == null ? null : p * purity; };
const valueNow = (it, c = cur()) => { const p = pureGramNow(c); return p == null ? null : pureGrams(it) * p; };
/** Converts an amount between currencies using the rates of a given day (or today's). */
function convert(amount, from, to, rs = S.rates) {
  if (from === to) return amount;
  const k = c => c === 'USD' ? rs?.usd : c === 'AED' ? rs?.aed : rs?.inr;
  const a = k(from), b = k(to);
  if (a && b) return amount * b / a;
  if (rs !== S.rates) return convert(amount, from, to, S.rates);
  return null;
}
/** What was paid, shown in the display currency (converted at the rate of the purchase day when known). */
const costIn = (it, c = cur()) => it.pricePaid ? convert(+it.pricePaid, it.currency || c, c, it.buyRate || S.rates) : null;
/** Market value of the gold on the day it was bought. */
function marketThen(it, c = cur()) {
  const r = it.buyRate; if (!r) return null;
  const perOz = c === 'USD' ? r.usd : c === 'AED' ? r.aed : r.inr;
  return perOz ? pureGrams(it) * perOz / OZ : null;
}
const held = () => (S.data?.items || []).filter(i => i.status === 'held');
function totals(list) {
  let value = 0, cost = 0, costKnown = 0, pure = 0, gross = 0, valueOfCosted = 0;
  for (const it of list) {
    const v = valueNow(it) || 0, c = costIn(it);
    value += v; pure += pureGrams(it); gross += +it.weight || 0;
    if (c != null) { cost += c; costKnown++; valueOfCosted += v; }
  }
  const gain = valueOfCosted - cost;
  return { value, cost, gain, gainPct: cost ? gain / cost * 100 : 0, pure, gross, count: list.length, costKnown };
}

/* ---------------- zakat ---------------- */
function hijriParts(dateStr) {
  try {
    const f = new Intl.DateTimeFormat('en-u-ca-islamic-umalqura', { day: 'numeric', month: 'numeric', year: 'numeric' });
    const p = Object.fromEntries(f.formatToParts(parseDay(dateStr)).map(x => [x.type, x.value]));
    return { d: +p.day, m: +p.month, y: parseInt(p.year) };
  } catch { return null; }
}
function hijriLabel(dateStr) {
  try { return new Intl.DateTimeFormat('en-u-ca-islamic-umalqura', { day: 'numeric', month: 'long', year: 'numeric' }).format(parseDay(dateStr)); }
  catch { return ''; }
}
/** Next 1 Ramadan on or after today (Umm al-Qura calendar). */
function nextRamadan(from = todayStr()) {
  for (let i = 0; i < 400; i++) { const d = addDays(from, i), h = hijriParts(d); if (h && h.m === 9 && h.d === 1) return d; }
  return addDays(from, 354);
}
const zakatDate = () => { const s = S.data?.settings?.zakatDate; return s && s >= todayStr() ? s : nextRamadan(); };
/** Why a piece is (or is not) counted for zakat. */
function zakatRule(it) {
  if (it.zakat === 'yes') return { due: true, why: 'Marked as zakatable' };
  if (it.zakat === 'no') return { due: false, why: 'Marked as not zakatable' };
  if (it.type !== 'jewellery') return { due: true, why: `${TYPES[it.type]?.label || 'Coin/bar'} — kept as savings` };
  if (it.purpose === 'investment') return { due: true, why: 'Jewellery kept as investment' };
  if (S.data?.settings?.zakatMethod === 'hanafi') return { due: true, why: 'Hanafi: all gold jewellery is zakatable' };
  return { due: false, why: 'Jewellery for personal wear — exempt (majority view)' };
}
function zakatFor(ownerId) {
  const list = held().filter(i => !ownerId || i.ownerId === ownerId);
  const zl = list.filter(i => zakatRule(i).due), ex = list.filter(i => !zakatRule(i).due);
  const pure = zl.reduce((a, i) => a + pureGrams(i), 0);
  const value = zl.reduce((a, i) => a + (valueNow(i) || 0), 0);
  const above = pure >= NISAB_G;
  return { list, zl, ex, pure, value, above, due: above ? value * 0.025 : 0, dueGrams: above ? pure * 0.025 : 0,
    exPure: ex.reduce((a, i) => a + pureGrams(i), 0) };
}

/* ---------------- UI shell ---------------- */
let toastTimer;
function toast(msg) {
  document.querySelector('.toast')?.remove();
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = msg; document.body.appendChild(t);
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.remove(), 3600);
}
function applyTheme() {
  const t = S.data?.settings?.theme || ls.get('mg.theme') || 'auto';
  if (t === 'auto') document.documentElement.removeAttribute('data-theme'); else document.documentElement.dataset.theme = t;
}
const ICON = {
  home: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/></svg>',
  items: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 17h16l-2.5-6h-11z"/><path d="M8 11 9.5 7h5L16 11"/></svg>',
  zakat: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z"/><path d="m17 4 .7 1.6 1.8.2-1.3 1.2.4 1.8-1.6-.9-1.6.9.4-1.8-1.3-1.2 1.8-.2z"/></svg>',
  settings: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
};
const TABS = [['home', 'Home'], ['items', 'My Gold'], ['zakat', 'Zakat'], ['settings', 'Settings']];

function render() {
  const app = $('#app');
  const sy = window.scrollY;
  if (S.phase === 'loading') { app.innerHTML = `<div class="welcome"><img src="icons/icon-512.png" alt=""><h1>My Gold</h1><p><span class="spin"></span></p></div>`; return; }
  if (S.phase === 'welcome' || S.phase === 'needsignin') { app.innerHTML = welcomeView(); return; }
  if (S.phase === 'folder') { app.innerHTML = folderView(); folderBrowser($('#fb')); return; }
  if (S.phase === 'setup') { app.innerHTML = setupView(); return; }
  const title = { home: 'My Gold', items: 'Collection', zakat: 'Zakat', settings: 'Settings' }[S.tab];
  const body = S.tab === 'items' ? itemsView() : S.tab === 'zakat' ? zakatView() : S.tab === 'settings' ? settingsView() : homeView();
  app.innerHTML = `
    <header class="top"><div class="topbar"><h1>${title}</h1>
      ${S.tab === 'home' ? `<button class="iconbtn" data-a="refresh" title="Refresh prices">${S.ratesBusy ? '<span class="spin" style="width:16px;height:16px"></span>' : '↻'}</button>` : ''}
    </div><div class="sync" id="syncline">${syncLine()}</div></header>
    <main>${body}</main>
    ${S.tab === 'settings' || S.tab === 'zakat' ? '' : `<button class="fab" data-a="add"><span style="font-size:19px;line-height:1">＋</span> Add gold</button>`}
    <nav class="tabs">${TABS.map(([k, l]) => `<button data-tab="${k}" class="${S.tab === k ? 'on' : ''}">${ICON[k]}<span>${l}</span></button>`).join('')}</nav>`;
  window.scrollTo(0, sy);
}
function paintRates() { const b = $('[data-a="refresh"]'); if (b) b.innerHTML = S.ratesBusy ? '<span class="spin" style="width:16px;height:16px"></span>' : '↻'; }

/* ---------------- welcome / setup ---------------- */
function welcomeView() {
  const again = S.phase === 'needsignin';
  return `<div class="welcome">
    <img src="icons/icon-512.png" alt="">
    <h1>My Gold</h1>
    <p>${again ? 'Your Microsoft sign-in has expired. Sign in again to continue — your data is safe in OneDrive.' :
      'Every coin, bar and piece of jewellery your family owns — with photos, bills, live value and zakat. Kept in your own OneDrive.'}</p>
    <div class="feat"><div>📸 Photos &amp; bills of each piece</div><div>📈 Live gold price</div><div>👨‍👩‍👧 Per person &amp; family total</div><div>🌙 Zakat for Ramadan</div></div>
    <div class="acts">
      <button class="btn" data-a="signin">Sign in with Microsoft</button>
      ${again ? `<button class="btn sec" data-a="signin-other">Use a different account</button>` : `<button class="btn sec" data-a="demo">Look around with sample data</button>`}
    </div>
    <p class="small" style="opacity:.7;margin-top:18px">Your data stays in your OneDrive. Nothing is stored anywhere else.</p>
  </div>`;
}
function folderView() {
  return `<header class="top"><div class="topbar"><h1>Where to keep it</h1></div></header>
    <main><div class="note">Choose (or create) a folder in your OneDrive — for example <b>Documents › My Gold</b>. Your gold record, photos and bills will all be kept there.</div>
    <div id="fb"></div></main>`;
}
function setupView() {
  const names = (S.setupNames ||= ['Me', 'Wife', 'Daughter']);
  return `<header class="top"><div class="topbar"><h1>Your family</h1></div></header>
    <main>
      <div class="card"><h3>Who owns gold?</h3>
        <div class="small muted">Each person's gold is kept under their own name. You can change these later.</div>
        <div id="setupNames">${names.map((n, i) => `<div class="field"><span>Person ${i + 1}</span><input data-setup="${i}" value="${esc(n)}" placeholder="Name"></div>`).join('')}</div>
        <button class="linkbtn" data-a="setup-more" style="margin-top:10px">＋ Add another person</button>
      </div>
      <div class="card"><h3>Show amounts in</h3>
        <div class="pills" id="setupCur">${Object.keys(CURS).map(c => `<button data-cur="${c}" class="${(S.setupCur || 'AED') === c ? 'on' : ''}">${c}</button>`).join('')}</div>
      </div>
      <button class="btn" data-a="setup-done" style="margin-top:18px">Start</button>
    </main>`;
}

/* ---------------- Home ---------------- */
function ratesStrip() {
  const r = S.rates, c = cur();
  const chips = KARATS.slice(0, 4).map(k => `<div class="rate"><div class="k">${k.label} · 1 g</div><div class="v num">${money(gramRate(k.p, c), c)}</div></div>`).join('');
  const when = r ? (r.quoteAt && /T/.test(r.quoteAt) ? new Date(r.quoteAt).toLocaleString([], { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }) : timeStr(r.at)) : '';
  return `<div class="rates">${chips}<div class="rate"><div class="k">1 troy oz</div><div class="v num">${money(r ? (c === 'USD' ? r.usd : c === 'AED' ? r.aed : r.inr) : null, c, 0)}</div></div></div>
    <div class="small faint" style="margin-top:6px">${r ? `International spot · ${esc(when)}${+S.data.settings.premium[c] ? ` · +${S.data.settings.premium[c]}% local adjustment` : ''}` : 'Getting today’s gold price…'}${S.ratesErr ? ` · <span style="color:var(--bad)">${esc(S.ratesErr)}</span>` : ''}</div>`;
}
function homeView() {
  const all = held(), t = totals(all), c = cur();
  if (!all.length) return `${ratesStrip()}
    <div class="card" style="text-align:center;padding:30px 18px">
      <div style="font-size:46px">🪙</div>
      <h2 class="serif" style="margin-top:8px">Add your first piece</h2>
      <p class="muted">Coins, bars or jewellery — add a photo and the bill, and My Gold works out what it's worth today.</p>
      <button class="btn" data-a="add" style="max-width:260px;margin:14px auto 0">＋ Add gold</button></div>`;
  const zDue = S.data.settings.zakatCombine ? zakatFor(null).due : owners().reduce((a, o) => a + zakatFor(o.id).due, 0), zd = zakatDate();
  const byType = Object.keys(TYPES).map(k => ({ k, v: totals(all.filter(i => i.type === k)) })).filter(x => x.v.count);
  const byKarat = [...new Set(all.map(karatLabel))].map(l => ({ l, v: totals(all.filter(i => karatLabel(i) === l)) })).sort((a, b) => b.v.value - a.v.value);
  const KC = ['#D4A017', '#B45309', '#EAB308', '#A16207', '#F59E0B', '#78350F'];
  const recent = [...all].sort((a, b) => cmpText(b.createdAt || '', a.createdAt || '')).slice(0, 4);
  return `
    <div class="hero">
      <div class="lbl">Family gold · today</div>
      <div class="big num">${money(t.value, c, 0)}</div>
      <div class="sub num">${grams(t.gross, 2)} · ${grams(t.pure, 2)} pure gold · ${t.count} piece${t.count === 1 ? '' : 's'}</div>
      <div class="stats">
        <div class="stat"><div class="k">Paid</div><div class="v num">${t.costKnown ? moneyK(t.cost, c) : '–'}</div></div>
        <div class="stat"><div class="k">Gain</div><div class="v num ${t.gain >= 0 ? 'up' : 'down'}">${t.costKnown ? (t.gain >= 0 ? '+' : '') + moneyK(t.gain, c) : '–'}</div></div>
        <div class="stat"><div class="k">Return</div><div class="v num ${t.gain >= 0 ? 'up' : 'down'}">${t.costKnown ? pct(t.gainPct) : '–'}</div></div>
      </div>
    </div>
    ${ratesStrip()}
    <div class="people">${owners().map(o => { const ot = totals(all.filter(i => i.ownerId === o.id));
      return `<button class="person" data-a="owner" data-id="${o.id}">
        <div class="ph"><span class="av" style="--c:${o.color}">${esc(initials(o.name))}</span><span class="name">${esc(o.name)}</span>
        <span style="flex:1"></span>${ot.costKnown ? `<span class="chip ${ot.gain >= 0 ? 'good' : 'bad'}">${pct(ot.gainPct)}</span>` : ''}</div>
        <div class="val num">${money(ot.value, c, 0)}</div>
        <div class="meta num"><span>${grams(ot.gross, 2)}</span><span>·</span><span>${ot.count} piece${ot.count === 1 ? '' : 's'}</span><span>·</span><span>${t.value ? Math.round(ot.value / t.value * 100) : 0}% of family</span></div>
      </button>`; }).join('')}</div>
    <div class="grid2">
      <div class="card"><h3>By type</h3>
        <div class="bar">${byType.map(x => `<i style="width:${x.v.value / t.value * 100}%;background:${TYPES[x.k].color}"></i>`).join('')}</div>
        <div class="legend">${byType.map(x => `<span><i class="sw" style="--c:${TYPES[x.k].color}"></i>${TYPES[x.k].plural} <b class="num">${money(x.v.value, c, 0)}</b> <span class="faint num">${grams(x.v.gross, 1)}</span></span>`).join('')}</div>
      </div>
      <div class="card"><h3>By purity</h3>
        <div class="bar">${byKarat.map((x, i) => `<i style="width:${x.v.value / t.value * 100}%;background:${KC[i % KC.length]}"></i>`).join('')}</div>
        <div class="legend">${byKarat.map((x, i) => `<span><i class="sw" style="--c:${KC[i % KC.length]}"></i>${esc(x.l)} <b class="num">${grams(x.v.gross, 2)}</b></span>`).join('')}</div>
      </div>
    </div>
    <button class="card" data-tab="zakat" style="width:100%;text-align:left;display:flex;gap:14px;align-items:center;font:inherit">
      <span class="ic" style="width:46px;height:46px;font-size:22px">🌙</span>
      <span style="flex:1;min-width:0"><span style="display:block;font-weight:700">Zakat ${zDue ? 'due' : ''} · ${esc(hijriLabel(zd) || shortDate(zd))}</span>
      <span class="small muted">${shortDate(zd)} · in ${daysBetween(todayStr(), zd)} days</span></span>
      <span class="serif num" style="font-size:20px;font-weight:700">${zDue ? money(zDue, c, 0) : 'None due'}</span>
    </button>
    <div class="card" style="padding:0;overflow:hidden"><h3 style="padding:14px 14px 0">Recently added <span class="grow"></span><button class="linkbtn" data-tab="items" style="text-transform:none;letter-spacing:0;font-size:14px">See all</button></h3>
      <div>${recent.map(itemRow).join('')}</div></div>`;
}

/* ---------------- Collection ---------------- */
function itemRow(it) {
  const o = ownerById(it.ownerId), v = valueNow(it), cst = costIn(it), gain = v != null && cst ? v - cst : null;
  const ph = it.photos?.[0]?.thumb;
  return `<button class="row" data-a="open" data-id="${it.id}">
    <span class="thumb" style="${ph ? `background-image:url('${ph}')` : ''}">${ph ? '' : TYPES[it.type]?.icon || '🪙'}</span>
    <span class="main"><span class="t">${esc(it.title || TYPES[it.type]?.label)}</span>
      <span class="s"><span class="chip gold">${esc(karatLabel(it))}</span> ${grams(+it.weight, 3)} · ${esc(o?.name || '–')}${it.status !== 'held' ? ` · <b>${it.status === 'sold' ? 'Sold' : 'Gifted'}</b>` : ''}</span></span>
    <span class="r"><div class="a num">${it.status === 'held' ? money(v, cur(), 0) : money(convert(+it.soldPrice || 0, it.currency || cur(), cur()), cur(), 0)}</div>
      ${gain != null && it.status === 'held' ? `<div class="b num ${gain >= 0 ? 'up' : 'down'}">${gain >= 0 ? '+' : ''}${money(gain, cur(), 0)}</div>` : `<div class="b faint">${shortDate(it.purchaseDate)}</div>`}</span>
  </button>`;
}
function filteredItems() {
  const f = S.f, q = f.q.trim().toLowerCase();
  let list = (S.data.items || []).filter(i => (f.sold ? i.status !== 'held' : i.status === 'held'))
    .filter(i => !f.owner || i.ownerId === f.owner).filter(i => !f.type || i.type === f.type)
    .filter(i => !q || [i.title, i.shop, i.invoiceNo, i.notes, karatLabel(i), ownerById(i.ownerId)?.name].join(' ').toLowerCase().includes(q));
  const by = { newest: (a, b) => cmpText(b.purchaseDate || '', a.purchaseDate || ''), oldest: (a, b) => cmpText(a.purchaseDate || '', b.purchaseDate || ''),
    value: (a, b) => (valueNow(b) || 0) - (valueNow(a) || 0), weight: (a, b) => (+b.weight || 0) - (+a.weight || 0), name: (a, b) => cmpText(a.title, b.title) }[f.sort];
  return list.sort(by);
}
function itemsView() {
  const f = S.f, list = filteredItems(), t = totals(list.filter(i => i.status === 'held'));
  return `<div class="filters">
      <div class="seg">${[['', 'Everyone'], ...owners().map(o => [o.id, o.name])].map(([id, n]) => `<button data-fo="${id}" class="${f.owner === id ? 'on' : ''}">${esc(n)}</button>`).join('')}</div>
      <div class="seg">${[['', 'All'], ...Object.entries(TYPES).map(([k, v]) => [k, v.plural])].map(([k, n]) => `<button data-ft="${k}" class="${f.type === k ? 'on' : ''}">${n}</button>`).join('')}</div>
      <div style="display:flex;gap:8px">
        <label class="search" style="flex:1">🔍<input id="q" type="search" placeholder="Search name, shop, bill no." value="${esc(f.q)}"></label>
        <select id="sort" class="search" style="padding:0 10px;flex:none;max-width:130px">${[['newest', 'Newest'], ['oldest', 'Oldest'], ['value', 'Value'], ['weight', 'Weight'], ['name', 'Name']].map(([k, l]) => `<option value="${k}" ${f.sort === k ? 'selected' : ''}>${l}</option>`).join('')}</select>
      </div>
    </div>
    ${!f.sold && list.length ? `<div class="split" style="margin-top:12px">
      <div class="box"><div class="k">Value today</div><div class="v num">${money(t.value, cur(), 0)}</div></div>
      <div class="box"><div class="k">Weight · pure</div><div class="v num">${grams(t.gross, 2)} <span class="small muted">${grams(t.pure, 2)}</span></div></div></div>` : ''}
    <div class="list">${list.map(itemRow).join('') || `<div class="empty">${f.sold ? 'Nothing sold or gifted.' : (S.data.items.length ? 'Nothing matches.' : 'No gold added yet.')}</div>`}</div>
    <div style="text-align:center;margin-top:12px"><button class="linkbtn" data-a="toggle-sold">${f.sold ? '‹ Back to gold we hold' : 'Show sold & gifted'}</button></div>`;
}

/* ---------------- Zakat ---------------- */
function zakatView() {
  const s = S.data.settings, zd = zakatDate(), c = cur(), nisabVal = pureGramNow(c) == null ? null : NISAB_G * pureGramNow(c);
  const per = owners().map(o => ({ o, z: zakatFor(o.id) }));
  const fam = zakatFor(null), combine = !!s.zakatCombine;
  const total = combine ? fam.due : per.reduce((a, x) => a + x.z.due, 0), totalG = combine ? fam.dueGrams : per.reduce((a, x) => a + x.z.dueGrams, 0);
  const young = held().filter(i => i.purchaseDate && daysBetween(i.purchaseDate, zd) < HAWL_DAYS && zakatRule(i).due);
  return `
    <div class="hero">
      <div class="lbl">Zakat on gold · ${esc(hijriLabel(zd))}</div>
      <div class="big num">${money(total, c, 0)}</div>
      <div class="sub">${total ? `or ${grams(totalG, 2)} of pure gold · ` : ''}${shortDate(zd)} · in ${daysBetween(todayStr(), zd)} days</div>
      <div class="stats">
        <div class="stat"><div class="k">Nisab</div><div class="v num">${NISAB_G} g pure</div></div>
        <div class="stat" style="grid-column:span 2"><div class="k">Nisab value today</div><div class="v num">${money(nisabVal, c, 0)}</div></div>
      </div>
    </div>
    <div class="card"><h3>How jewellery is treated</h3>
      <div class="seg"><button data-zm="majority" class="${s.zakatMethod !== 'hanafi' ? 'on' : ''}">Shafi'i · Maliki · Hanbali</button><button data-zm="hanafi" class="${s.zakatMethod === 'hanafi' ? 'on' : ''}">Hanafi</button></div>
      <div class="small muted" style="margin-top:8px">${s.zakatMethod === 'hanafi' ? 'All gold, including jewellery worn, is counted.' : 'Coins, bars and investment jewellery are counted. Jewellery worn for personal use is exempt.'} You can override any piece when editing it.</div>
      <div class="field"><span>Count gold</span><div class="seg"><button data-zc="0" class="${!combine ? 'on' : ''}">Per person</button><button data-zc="1" class="${combine ? 'on' : ''}">Family together</button></div>
        <span class="hint">${combine ? `Family zakatable gold: ${grams(fam.pure, 2)} pure — ${fam.above ? 'above' : 'below'} the ${NISAB_G} g nisab.` : 'Zakat is each owner’s duty, so each person is checked against nisab on their own.'}</span></div>
      <div class="field"><span>Zakat date</span><input type="date" id="zdate" value="${esc(s.zakatDate || '')}"><span class="hint">Leave empty to use 1 Ramadan (${shortDate(nextRamadan())}).</span></div>
    </div>
    ${per.map(({ o, z }) => `<div class="card">
      <div class="ph" style="margin-bottom:12px"><span class="av" style="--c:${o.color}">${esc(initials(o.name))}</span><span style="font-weight:700;font-size:16px;flex:1">${esc(o.name)}</span>
        ${z.zl.length ? `<span class="chip ${z.above ? 'gold' : ''}">${z.above ? 'Above nisab' : 'Below nisab'}</span>` : ''}</div>
      <div class="split">
        <div class="box"><div class="k">Zakatable gold</div><div class="v num">${grams(z.pure, 2)}</div><div class="small muted num">${money(z.value, c, 0)}</div></div>
        <div class="box"><div class="k">Zakat due</div><div class="v num" style="color:var(--goldInk)">${money(z.due, c, 0)}</div><div class="small muted num">${z.above ? grams(z.dueGrams, 2) + ' pure gold' : 'Nothing due'}</div></div>
      </div>
      <div class="bar" style="margin-top:12px"><i style="width:${Math.min(100, z.pure / NISAB_G * 100)}%;background:linear-gradient(90deg,var(--gold2),var(--gold))"></i></div>
      <div class="small muted" style="margin-top:5px">${grams(z.pure, 1)} of ${NISAB_G} g nisab${z.ex.length ? ` · ${grams(z.exPure, 1)} pure in exempt jewellery` : ''}</div>
      ${z.list.length ? `<div class="list" style="margin-top:12px">${z.list.map(i => { const r = zakatRule(i);
        return `<button class="row" data-a="open" data-id="${i.id}"><span class="ic">${r.due ? '✅' : '➖'}</span><span class="main"><span class="t">${esc(i.title)}</span><span class="s">${esc(r.why)}</span></span><span class="r num small">${grams(pureGrams(i), 2)}</span></button>`; }).join('')}</div>` : '<div class="small faint" style="margin-top:8px">No gold yet.</div>'}
    </div>`).join('')}
    ${young.length ? `<div class="note info">ℹ️ ${young.length} zakatable piece${young.length === 1 ? ' was' : 's were'} bought less than a lunar year before the zakat date (${young.map(i => esc(i.title)).join(', ')}). Most scholars add new gold to your yearly total once you are above nisab; ask your scholar if unsure.</div>` : ''}
    <div class="card"><h3>Zakat paid <span class="grow"></span><button class="linkbtn" data-a="zpay" style="text-transform:none;letter-spacing:0;font-size:14px">＋ Record</button></h3>
      ${(S.data.zakatLog || []).length ? `<div class="list" style="margin-top:0">${[...S.data.zakatLog].sort((a, b) => cmpText(b.date, a.date)).map(z => `<button class="row" data-a="zdel" data-id="${z.id}"><span class="ic">🌙</span><span class="main"><span class="t">${esc(ownerById(z.ownerId)?.name || 'Family')} · ${esc(hijriLabel(z.date))}</span><span class="s">${shortDate(z.date)}${z.note ? ' · ' + esc(z.note) : ''}</span></span><span class="r num"><b>${money(+z.amount, z.currency)}</b></span></button>`).join('')}</div>`
        : '<div class="small faint">Keep a record of zakat you have given each year.</div>'}
    </div>
    <div class="note">This is a guide using international gold prices and the common nisab of 85 g of pure gold at 2.5%. Please confirm with a scholar you trust, especially for jewellery and for gold owned by children.</div>`;
}

/* ---------------- Settings ---------------- */
function settingsView() {
  const s = S.data.settings;
  return `
    <div class="card"><h3>Family <span class="grow"></span><button class="linkbtn" data-a="owner-add" style="text-transform:none;letter-spacing:0;font-size:14px">＋ Add</button></h3>
      <div class="list" style="margin-top:0">${owners().map(o => `<button class="row" data-a="owner-edit" data-id="${o.id}"><span class="av" style="--c:${o.color};width:34px;height:34px;font-size:13px">${esc(initials(o.name))}</span><span class="main"><span class="t">${esc(o.name)}</span><span class="s">${held().filter(i => i.ownerId === o.id).length} pieces</span></span><span class="faint">›</span></button>`).join('')}</div>
    </div>
    <div class="card"><h3>Show amounts in</h3>
      <div class="seg">${Object.keys(CURS).map(c => `<button data-setcur="${c}" class="${s.currency === c ? 'on' : ''}">${c}</button>`).join('')}</div>
      <div class="field"><span>Local price adjustment for ${s.currency} (%)</span>
        <input id="prem" type="number" inputmode="decimal" step="0.1" value="${+s.premium[s.currency] || 0}">
        <span class="hint">Values use the international spot price. Add a percentage if you want to match your local jeweller's rate (India's rates include import duty and GST, for example).</span></div>
    </div>
    <div class="card"><h3>Appearance</h3>
      <div class="seg">${[['auto', 'Automatic'], ['light', 'Light'], ['dark', 'Dark']].map(([k, l]) => `<button data-theme="${k}" class="${(s.theme || 'auto') === k ? 'on' : ''}">${l}</button>`).join('')}</div>
    </div>
    <div class="card"><h3>Storage</h3>
      ${S.demo ? `<div class="small muted">You are looking at sample data on this device. Sign in to keep your real gold record in OneDrive.</div><button class="btn" data-a="leave-demo" style="margin-top:12px">Sign in with Microsoft</button>` : `
      <dl class="kv"><dt>Account</dt><dd>${esc(S.account || '–')}</dd><dt>Folder</dt><dd>${esc(locFull())}</dd></dl>
      <div style="display:flex;gap:8px;margin-top:12px"><button class="btn sec" data-a="sync">Sync now</button><button class="btn sec" data-a="change-folder">Change folder</button></div>`}
    </div>
    <div class="card"><h3>Export</h3>
      <div style="display:flex;gap:8px"><button class="btn sec" data-a="csv">Spreadsheet (CSV)</button><button class="btn sec" data-a="json">Full backup (JSON)</button></div>
    </div>
    <div class="card"><h3>Gold prices</h3><div class="small muted">Live spot price from gold-api.com. Purchase-day prices from the free currency-api (since March 2024) and freegoldapi.com for older dates. Rates for each piece are saved with it, so history never changes.</div></div>
    <button class="btn danger" data-a="signout" style="margin-top:16px">${S.demo ? 'Leave sample data' : 'Sign out'}</button>
    <div class="small faint" style="text-align:center;margin-top:14px">My Gold ${CFG.version}</div>`;
}

/* ---------------- sheets ---------------- */
let sheetEl, scrimEl, sheetOnClose;
function openSheet(title, html, { onClose, right = '' } = {}) {
  closeSheet(true);
  scrimEl = document.createElement('div'); scrimEl.className = 'scrim';
  sheetEl = document.createElement('div'); sheetEl.className = 'sheet';
  sheetEl.innerHTML = `<div class="sh"><button class="linkbtn" data-a="close" style="font-weight:600">Close</button><h2>${esc(title)}</h2>${right}</div><div class="sb">${html}</div>`;
  document.body.append(scrimEl, sheetEl);
  scrimEl.addEventListener('click', () => closeSheet());
  sheetEl.addEventListener('click', onClick);
  sheetEl.addEventListener('change', onChange);
  sheetEl.addEventListener('input', onInput);
  requestAnimationFrame(() => { scrimEl.classList.add('on'); sheetEl.classList.add('on'); });
  sheetOnClose = onClose; document.body.style.overflow = 'hidden';
  return sheetEl;
}
function closeSheet(instant) {
  if (!sheetEl) return;
  const s = sheetEl, c = scrimEl; sheetEl = scrimEl = null; document.body.style.overflow = '';
  if (instant) { s.remove(); c.remove(); } else { s.classList.remove('on'); c.classList.remove('on'); setTimeout(() => { s.remove(); c.remove(); }, 260); }
  const f = sheetOnClose; sheetOnClose = null; f?.();
}

/* ---- piece detail ---- */
function openItem(id) {
  const it = S.data.items.find(i => i.id === id); if (!it) return;
  const o = ownerById(it.ownerId), c = cur(), v = valueNow(it), cst = costIn(it), then = marketThen(it), gain = v != null && cst ? v - cst : null;
  const zr = zakatRule(it), cover = it.photos[0];
  const rateThen = it.buyRate ? (c === 'USD' ? it.buyRate.usd : c === 'AED' ? it.buyRate.aed : it.buyRate.inr) : null;
  openSheet(it.title || 'Gold', `
    ${cover ? `<div class="photo-hero" data-a="photo" data-pid="${cover.id}" style="background-image:url('${cover.local || cover.thumb}')"></div>`
      : `<label class="photo-hero empty"><span style="font-size:40px">${TYPES[it.type]?.icon}</span><span class="small" style="font-weight:700;color:var(--goldInk)">📷 Add a photo</span><input type="file" accept="image/*" hidden data-up="photo" data-id="${it.id}"></label>`}
    ${cover ? `<div class="gallery" style="margin-top:8px">${it.photos.slice(1).map(p => `<div class="g" data-a="photo" data-pid="${p.id}" style="background-image:url('${p.local || p.thumb}')"></div>`).join('')}
      <label class="g add">＋ Photo<input type="file" accept="image/*" hidden data-up="photo" data-id="${it.id}"></label></div>` : ''}
    <div style="display:flex;gap:6px;flex-wrap:wrap;margin-top:10px">
      <span class="chip" style="background:${o?.color};color:#fff">${esc(o?.name || '–')}</span>
      <span class="chip gold">${esc(karatLabel(it))}</span><span class="chip">${TYPES[it.type]?.label}</span>
      <span class="chip ${zr.due ? 'gold' : ''}">${zr.due ? 'Zakatable' : 'Zakat exempt'}</span>
      ${it.status !== 'held' ? `<span class="chip bad">${it.status === 'sold' ? 'Sold' : 'Gifted'} ${shortDate(it.soldDate)}</span>` : ''}
    </div>
    <div class="split" style="margin-top:12px">
      <div class="box"><div class="k">Worth today</div><div class="v num">${money(v, c, 0)}</div></div>
      <div class="box"><div class="k">You paid</div><div class="v num">${cst != null ? money(cst, c, 0) : '–'}</div>${cst != null && it.currency && it.currency !== c ? `<div class="small muted num">${money(+it.pricePaid, it.currency, 0)}</div>` : ''}</div>
      <div class="box"><div class="k">Gain</div><div class="v num ${gain >= 0 ? 'up' : 'down'}">${gain != null ? (gain >= 0 ? '+' : '') + money(gain, c, 0) : '–'}</div>${gain != null && cst ? `<div class="small num ${gain >= 0 ? 'up' : 'down'}">${pct(gain / cst * 100)}</div>` : ''}</div>
      <div class="box"><div class="k">Gold value on purchase day</div><div class="v num">${then != null ? money(then, c, 0) : '–'}</div>${then != null && cst ? `<div class="small muted num">Paid ${pct((cst - then) / then * 100)} over gold value</div>` : ''}</div>
    </div>
    <div class="card"><dl class="kv">
      <dt>Bought</dt><dd>${shortDate(it.purchaseDate)}</dd>
      ${it.shop ? `<dt>From</dt><dd>${esc(it.shop)}</dd>` : ''}
      ${it.invoiceNo ? `<dt>Bill no.</dt><dd>${esc(it.invoiceNo)}</dd>` : ''}
      <dt>Weight</dt><dd class="num">${grams(+it.weight)}</dd>
      ${it.type === 'jewellery' && +it.stoneWeight ? `<dt>Stones</dt><dd class="num">− ${grams(+it.stoneWeight)}</dd><dt>Gold weight</dt><dd class="num">${grams(netWeight(it))}</dd>` : ''}
      <dt>Purity</dt><dd class="num">${esc(karatLabel(it))} · ${trimZ((purityOf(it) * 100).toFixed(2))}%</dd>
      <dt>Pure gold</dt><dd class="num">${grams(pureGrams(it))}</dd>
      ${+it.billRate ? `<dt>Rate on bill</dt><dd class="num">${money(+it.billRate, it.currency)}/g</dd>` : ''}
      ${rateThen ? `<dt>Market rate then</dt><dd class="num">${money(rateThen / OZ * purityOf(it), c)}/g</dd>` : ''}
      <dt>Market rate today</dt><dd class="num">${money(gramRate(purityOf(it), c), c)}/g</dd>
      ${+it.making ? `<dt>Making charges</dt><dd class="num">${money(+it.making, it.currency)}</dd>` : ''}
      ${+it.vat ? `<dt>VAT / GST</dt><dd class="num">${money(+it.vat, it.currency)}</dd>` : ''}
      <dt>Zakat</dt><dd style="font-weight:500">${esc(zr.why)}</dd>
      ${it.status === 'sold' && +it.soldPrice ? `<dt>Sold for</dt><dd class="num">${money(+it.soldPrice, it.currency)}</dd>` : ''}
    </dl>${it.buyRate ? `<div class="small faint" style="margin-top:8px">Purchase-day price from ${esc(it.buyRate.src)} (${shortDate(it.buyRate.date)}).</div>` : ''}</div>
    <div class="card"><h3>Bills &amp; receipts</h3>
      <div class="list" style="margin-top:0">${it.docs.map(d => `<button class="row" data-a="doc" data-did="${d.id}"><span class="thumb" style="width:42px;height:42px;font-size:18px;${d.thumb ? `background-image:url('${d.thumb}')` : ''}">${d.thumb ? '' : '📄'}</span><span class="main"><span class="t">${esc(d.kind)}</span><span class="s">${esc(d.name)}</span></span><span class="faint">›</span></button>`).join('') || '<div class="empty small">No bill added yet.</div>'}</div>
      <div style="display:flex;gap:8px;margin-top:10px">
        <label class="btn sec" style="flex:1">📷 Photo of bill<input type="file" accept="image/*" capture="environment" hidden data-up="bill" data-id="${it.id}"></label>
        <label class="btn sec" style="flex:1">📎 File / PDF<input type="file" accept="image/*,application/pdf" hidden data-up="bill" data-id="${it.id}"></label>
      </div>
    </div>
    ${it.notes ? `<div class="card"><h3>Notes</h3><div style="white-space:pre-wrap">${esc(it.notes)}</div></div>` : ''}
    <div style="display:flex;gap:8px;margin-top:16px">
      <button class="btn" data-a="edit" data-id="${it.id}" style="flex:1.3">Edit</button>
      ${it.status === 'held' ? `<button class="btn sec" data-a="dispose" data-id="${it.id}" style="flex:1">Sold / gifted</button>` : `<button class="btn sec" data-a="unsell" data-id="${it.id}" style="flex:1">Still have it</button>`}
    </div>
    <button class="btn danger" data-a="del" data-id="${it.id}" style="margin-top:8px">Delete</button>`);
  sheetEl.dataset.item = id;
}

/* ---- add / edit form ---- */
let draft = null;
function openForm(id, presetOwner) {
  const it = id ? S.data.items.find(i => i.id === id) : null;
  draft = it ? JSON.parse(JSON.stringify(it)) : {
    id: uuid(), ownerId: presetOwner || S.f.owner || owners()[0]?.id, type: 'jewellery', karat: 22, title: '', weight: '', stoneWeight: '',
    purchaseDate: todayStr(), pricePaid: '', currency: cur(), billRate: '', making: '', vat: '', shop: '', invoiceNo: '',
    purpose: 'personal', zakat: 'auto', notes: '', photos: [], docs: [], status: 'held',
  };
  draft._newPhotos = []; draft._newDocs = []; draft._isNew = !it;
  openSheet(it ? 'Edit' : 'Add gold', formHtml(), { right: `<button class="linkbtn" data-a="save">Save</button>` });
}
function formHtml() {
  const d = draft, karatOther = !KARATS.some(k => k.k === +d.karat);
  return `<form id="gform" onsubmit="return false">
    <div class="field"><span>Whose gold?</span><div class="pills">${owners().map(o => `<button type="button" data-fo2="${o.id}" class="${d.ownerId === o.id ? 'on' : ''}">${esc(o.name)}</button>`).join('')}</div></div>
    <div class="field"><span>Type</span><div class="pills">${Object.entries(TYPES).map(([k, v]) => `<button type="button" data-ty="${k}" class="${d.type === k ? 'on' : ''}">${v.icon} ${v.label}</button>`).join('')}</div></div>
    <div class="field"><span>Name</span><input name="title" value="${esc(d.title)}" placeholder="${d.type === 'coin' ? 'e.g. 10 g Federal Bank coin' : d.type === 'bar' ? 'e.g. 1 tola biscuit' : 'e.g. Wedding necklace'}" autocomplete="off"></div>
    <div class="field"><span>Purity</span><div class="pills">${KARATS.map(k => `<button type="button" data-ka="${k.k}" class="${+d.karat === k.k ? 'on' : ''}">${k.label}</button>`).join('')}<button type="button" data-ka="0" class="${karatOther ? 'on' : ''}">Other</button></div>
      ${karatOther ? `<input name="purity" type="number" inputmode="decimal" step="0.001" min="0.3" max="1" value="${esc(d.purity || '')}" placeholder="Fineness, e.g. 0.995" style="margin-top:6px">` : ''}</div>
    <div class="two">
      <div class="field"><span>Weight (g)</span><input name="weight" type="number" inputmode="decimal" step="0.001" min="0" value="${esc(d.weight)}" placeholder="0.000"></div>
      ${d.type === 'jewellery' ? `<div class="field"><span>Stones / beads (g)</span><input name="stoneWeight" type="number" inputmode="decimal" step="0.001" min="0" value="${esc(d.stoneWeight)}" placeholder="0"></div>`
        : `<div class="field"><span>Bought on</span><input name="purchaseDate" type="date" value="${esc(d.purchaseDate)}" max="${todayStr()}"></div>`}
    </div>
    ${d.type === 'jewellery' ? `<div class="field"><span>Bought on</span><input name="purchaseDate" type="date" value="${esc(d.purchaseDate)}" max="${todayStr()}"></div>` : ''}
    <div class="field"><span>Total paid (including making &amp; VAT)</span><div class="inwrap"><input name="pricePaid" type="number" inputmode="decimal" step="0.01" min="0" value="${esc(d.pricePaid)}" placeholder="0">
      <select name="currency">${Object.keys(CURS).map(c => `<option ${d.currency === c ? 'selected' : ''}>${c}</option>`).join('')}</select></div></div>
    <div class="two">
      <div class="field"><span>Rate on bill (per g)</span><input name="billRate" type="number" inputmode="decimal" step="0.01" min="0" value="${esc(d.billRate)}" placeholder="optional"></div>
      <div class="field"><span>Making charges</span><input name="making" type="number" inputmode="decimal" step="0.01" min="0" value="${esc(d.making)}" placeholder="optional"></div>
    </div>
    <div class="two">
      <div class="field"><span>VAT / GST</span><input name="vat" type="number" inputmode="decimal" step="0.01" min="0" value="${esc(d.vat)}" placeholder="optional"></div>
      <div class="field"><span>Bill no.</span><input name="invoiceNo" value="${esc(d.invoiceNo)}" placeholder="optional"></div>
    </div>
    <div class="field"><span>Shop / bank</span><input name="shop" value="${esc(d.shop)}" placeholder="e.g. Malabar Gold, Federal Bank" list="shops"><datalist id="shops">${[...new Set(S.data.items.map(i => i.shop).filter(Boolean))].map(s => `<option value="${esc(s)}">`).join('')}</datalist></div>
    ${d.type === 'jewellery' ? `<div class="field"><span>Kept for</span><div class="pills"><button type="button" data-pu="personal" class="${d.purpose !== 'investment' ? 'on' : ''}">Wearing</button><button type="button" data-pu="investment" class="${d.purpose === 'investment' ? 'on' : ''}">Investment</button></div></div>` : ''}
    <div class="field"><span>Zakat</span><select name="zakat"><option value="auto" ${d.zakat === 'auto' ? 'selected' : ''}>Automatic (${zakatRule({ ...d, zakat: 'auto' }).due ? 'zakatable' : 'exempt'})</option><option value="yes" ${d.zakat === 'yes' ? 'selected' : ''}>Always count for zakat</option><option value="no" ${d.zakat === 'no' ? 'selected' : ''}>Never count for zakat</option></select></div>
    <div class="field"><span>Photos</span><div class="gallery">
      ${[...d.photos, ...d._newPhotos].map(p => `<div class="g" style="background-image:url('${p.thumb || p.preview}')"><button type="button" class="x" data-rmp="${p.id}">✕</button></div>`).join('')}
      <label class="g add">📷<br>Camera<input type="file" accept="image/*" capture="environment" hidden data-fp="1"></label>
      <label class="g add">🖼️<br>Library<input type="file" accept="image/*" multiple hidden data-fp="1"></label></div></div>
    ${d._isNew ? `<div class="field"><span>Bill / receipt</span><div class="gallery">
      ${d._newDocs.map(p => `<div class="g" style="${p.preview ? `background-image:url('${p.preview}')` : ''};display:grid;place-items:center;font-size:12px;text-align:center">${p.preview ? '' : '📄<br>' + esc(p.file.name.slice(0, 14))}<button type="button" class="x" data-rmd="${p.id}">✕</button></div>`).join('')}
      <label class="g add">🧾<br>Photo of bill<input type="file" accept="image/*" capture="environment" hidden data-fd="1"></label>
      <label class="g add">📎<br>PDF / file<input type="file" accept="image/*,application/pdf" hidden data-fd="1"></label></div></div>` : ''}
    <div class="field"><span>Notes</span><textarea name="notes" rows="3" placeholder="Anything to remember — design, occasion, where it is kept…">${esc(d.notes)}</textarea></div>
    <button type="button" class="btn" data-a="save" style="margin-top:18px">Save</button>
  </form>`;
}
function readForm() {
  const f = $('#gform', sheetEl); if (!f) return;
  for (const el of f.elements) if (el.name) draft[el.name] = el.value;
}
function redrawForm() { readForm(); const sb = $('.sb', sheetEl), y = sb.scrollTop; sb.innerHTML = formHtml(); sb.scrollTop = y; }

async function saveForm(btn) {
  readForm();
  const d = draft;
  const w = num(d.weight);
  if (!d.ownerId) return toast('Choose whose gold this is.');
  if (!(w > 0)) return toast('Enter the weight in grams.');
  if (!+d.karat && !(num(d.purity) > 0.3 && num(d.purity) <= 1)) return toast('Enter the purity (e.g. 0.995).');
  if (d.type === 'jewellery' && num(d.stoneWeight) >= w) return toast('Stone weight must be less than the total weight.');
  document.querySelectorAll('[data-a="save"]').forEach(b => { b.disabled = true; b.innerHTML = '<span class="spin" style="width:16px;height:16px"></span> Saving'; });
  try {
    const owner = ownerById(d.ownerId)?.name || '';
    const title = d.title.trim() || `${karatLabel(d)} ${TYPES[d.type].label.toLowerCase()}`;
    let buyRate = d.buyRate;
    if (d.purchaseDate && (!buyRate || d._rateFor !== d.purchaseDate) && (!buyRate || buyRate.forDay !== d.purchaseDate)) {
      const r = await rateOn(d.purchaseDate).catch(() => null);
      buyRate = r ? { ...r, forDay: d.purchaseDate } : null;
    }
    const photos = [];
    for (const p of d._newPhotos) photos.push(await savePhoto(p.file, owner, title));
    const docs = [];
    for (const p of d._newDocs) docs.push(await saveDoc(p.file, owner, title, 'Bill'));
    const clean = {
      id: d.id, ownerId: d.ownerId, type: d.type, title, karat: +d.karat || 0, purity: +d.karat ? undefined : num(d.purity),
      weight: round2(w * 10) / 10, stoneWeight: d.type === 'jewellery' && num(d.stoneWeight) > 0 ? num(d.stoneWeight) : 0,
      purchaseDate: d.purchaseDate || '', pricePaid: num(d.pricePaid) > 0 ? num(d.pricePaid) : 0, currency: d.currency,
      billRate: num(d.billRate) > 0 ? num(d.billRate) : 0, making: num(d.making) > 0 ? num(d.making) : 0, vat: num(d.vat) > 0 ? num(d.vat) : 0,
      shop: d.shop.trim(), invoiceNo: d.invoiceNo.trim(), purpose: d.type === 'jewellery' ? d.purpose : 'investment', zakat: d.zakat,
      notes: d.notes.trim(), buyRate: buyRate || null, status: d.status || 'held', soldDate: d.soldDate, soldPrice: d.soldPrice,
    };
    clean.weight = Math.round(w * 1000) / 1000;
    const keepPhotos = new Set(d.photos.map(p => p.id));
    const ok = await mutate(data => {
      const ex = data.items.find(i => i.id === d.id);
      if (ex) {
        const removed = ex.photos.filter(p => !keepPhotos.has(p.id));
        removed.forEach(p => removeFile('Photos', p));
        Object.assign(ex, clean, { photos: [...ex.photos.filter(p => keepPhotos.has(p.id)), ...photos], docs: [...ex.docs, ...docs], updatedAt: isoNow() });
      } else data.items.push({ ...clean, photos, docs, createdAt: isoNow(), updatedAt: isoNow() });
    });
    if (ok) { closeSheet(); toast(d._isNew ? 'Added ✓' : 'Saved ✓'); if (!buyRate && d.purchaseDate) toast('Saved. The gold price for that day could not be found — check your internet.'); openItem(d.id); }
  } catch (e) { toast(e.message); }
  document.querySelectorAll('[data-a="save"]').forEach(b => { b.disabled = false; b.textContent = 'Save'; });
}

/* ---- people ---- */
function openOwner(id) {
  const o = id ? ownerById(id) : { id: uuid(), name: '', color: OWNER_COLORS[owners().length % OWNER_COLORS.length] };
  const count = (S.data.items || []).filter(i => i.ownerId === o.id).length;
  openSheet(id ? 'Edit person' : 'Add person', `
    <div class="field"><span>Name</span><input id="oname" value="${esc(o.name)}" placeholder="Name"></div>
    <div class="field"><span>Colour</span><div class="pills" id="ocol">${OWNER_COLORS.map(c => `<button type="button" data-oc="${c}" class="${o.color === c ? 'on' : ''}" style="width:40px;height:40px;padding:0;background:${c};border-width:${o.color === c ? 3 : 1}px"></button>`).join('')}</div></div>
    <button class="btn" data-a="owner-save" data-id="${o.id}" style="margin-top:18px">Save</button>
    ${id ? `<button class="btn danger" data-a="owner-del" data-id="${o.id}" style="margin-top:8px" ${count ? 'disabled' : ''}>Remove</button>${count ? `<div class="small faint" style="text-align:center;margin-top:6px">${esc(o.name)} has ${count} piece${count === 1 ? '' : 's'}. Move or delete them first.</div>` : ''}` : ''}`);
  sheetEl.dataset.color = o.color;
}

/* ---- OneDrive folder browser ---- */
function folderBrowser(box) {
  const st = { stack: [{ name: 'My files', driveId: null, id: null, path: '' }], items: [], loading: true, err: '' };
  const here = () => st.stack[st.stack.length - 1];
  async function load() {
    st.loading = true; st.err = ''; draw();
    try {
      const h = here();
      let next = (h.id ? `/drives/${h.driveId}/items/${h.id}/children` : '/me/drive/root/children') + '?$select=id,name,folder,file,parentReference,root&$top=200', all = [];
      while (next) { const j = await (await g(next)).json(); all = all.concat(j.value || []); next = j['@odata.nextLink']; }
      if (!h.id) { try { const r = await (await g('/me/drive/root?$select=id,parentReference')).json(); h.id = r.id; h.driveId = r.parentReference?.driveId; } catch {} }
      st.items = all.filter(x => x.folder || x.name === CFG.dataFile).sort((a, b) => (!!b.folder - !!a.folder) || cmpText(a.name, b.name));
      if (!h.driveId && all[0]) h.driveId = all[0].parentReference?.driveId;
    } catch (e) { st.err = e.message; }
    st.loading = false; draw();
  }
  function draw() {
    const hasData = st.items.some(x => x.name === CFG.dataFile), atRoot = st.stack.length === 1;
    box.innerHTML = `<div class="list">
      <div class="gh" style="text-transform:none;letter-spacing:0;align-items:center;gap:8px">${atRoot ? '' : `<button class="linkbtn" data-b="up" style="padding:0">‹ Back</button>`}
        <span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">📁 ${esc(st.stack.map(x => x.name).join(' › '))}</span></div>
      ${st.loading ? '<div class="empty"><span class="spin"></span></div>' : st.err ? `<div class="empty" style="color:var(--bad)">${esc(st.err)}</div>` :
        (st.items.map((x, i) => x.folder
          ? `<button class="row" data-b="open" data-i="${i}"><span class="ic">📁</span><span class="main"><span class="t">${esc(x.name)}</span><span class="s">${x.folder.childCount} item${x.folder.childCount === 1 ? '' : 's'}</span></span><span class="faint">›</span></button>`
          : `<div class="row"><span class="ic">✅</span><span class="main"><span class="t">My Gold data found here</span><span class="s">${CFG.dataFile}</span></span></div>`).join('') || '<div class="empty">No folders here.</div>')}
    </div>
    <div style="display:flex;gap:8px;margin-top:10px"><button class="btn sec" data-b="new" style="flex:1">＋ New folder</button>
      <button class="btn" data-b="use" style="flex:1.4" ${atRoot || st.loading ? 'disabled' : ''}>${hasData ? 'Use this data' : 'Use this folder'}</button></div>
    <div class="small faint" style="margin-top:6px">${atRoot ? 'Open a folder, or tap “New folder” to make one called “My Gold”.' : `Data will be kept in <b>${esc(locFull({ path: here().path }))}</b>.`}</div>`;
  }
  box.addEventListener('click', async e => {
    const b = e.target.closest('[data-b]'); if (!b || st.loading) return;
    if (b.dataset.b === 'up') { st.stack.pop(); load(); }
    else if (b.dataset.b === 'open') { const x = st.items[+b.dataset.i]; st.stack.push({ name: x.name, id: x.id, driveId: x.parentReference?.driveId || here().driveId, path: [here().path, x.name].filter(Boolean).join('/') }); load(); }
    else if (b.dataset.b === 'new') {
      const name = (prompt('Name for the new folder:', 'My Gold') || '').trim().replace(/[\\/:*?"<>|]/g, ' '); if (!name) return;
      try {
        const h = here(), url = h.id ? `/drives/${h.driveId}/items/${h.id}/children` : '/me/drive/root/children';
        const f = await (await g(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' }) })).json();
        st.stack.push({ name: f.name, id: f.id, driveId: f.parentReference?.driveId || h.driveId, path: [h.path, f.name].filter(Boolean).join('/') }); load();
      } catch (err) { toast(err.message); }
    } else if (b.dataset.b === 'use') {
      const h = here(); b.disabled = true;
      try { saveLoc({ driveId: h.driveId, id: h.id, path: h.path }); await writePointer(LOC).catch(() => {}); closeSheet(); S.phase = 'loading'; render(); await pull(true); }
      catch (err) { toast(err.message); b.disabled = false; }
    }
  });
  load();
}

/* ---------------- sample data ---------------- */
function sampleData() {
  const me = uuid(), wife = uuid(), dau = uuid(), y = new Date().getFullYear();
  const it = (o, type, title, karat, weight, date, paid, currency, extra = {}) => ({ id: uuid(), ownerId: o, type, title, karat, weight, stoneWeight: 0,
    purchaseDate: date, pricePaid: paid, currency, billRate: 0, making: 0, vat: 0, shop: '', invoiceNo: '', purpose: type === 'jewellery' ? 'personal' : 'investment',
    zakat: 'auto', notes: '', photos: [], docs: [], status: 'held', createdAt: date + 'T10:00:00Z', ...extra });
  return normalize({
    owners: [{ id: me, name: 'Me', color: OWNER_COLORS[0] }, { id: wife, name: 'Wife', color: OWNER_COLORS[1] }, { id: dau, name: 'Daughter', color: OWNER_COLORS[2] }],
    items: [
      it(me, 'coin', '10 g Federal Bank coin', 24, 10, '2024-05-12', 2820, 'AED', { shop: 'Federal Bank', invoiceNo: 'FB-22841' }),
      it(me, 'bar', '1 tola biscuit', 24, 11.664, '2024-11-03', 3420, 'AED', { shop: 'Joyalukkas, Dubai' }),
      it(me, 'coin', '8 g sovereign', 22, 8, '2025-02-20', 61500, 'INR', { shop: 'Malabar Gold, Kozhikode' }),
      it(wife, 'jewellery', 'Wedding necklace', 22, 48.2, '2024-04-10', 14200, 'AED', { stoneWeight: 2.1, making: 1500, shop: 'Malabar Gold, Dubai' }),
      it(wife, 'jewellery', 'Pair of bangles', 22, 32, '2024-08-15', 9700, 'AED', { shop: 'Kalyan Jewellers' }),
      it(wife, 'coin', '20 g investment coin', 24, 20, '2025-01-08', 6100, 'AED', { shop: 'Federal Bank' }),
      it(wife, 'jewellery', 'Daily chain', 18, 6.5, '2024-12-01', 1650, 'AED'),
      it(dau, 'jewellery', 'Earrings', 22, 4, '2025-03-30', 1420, 'AED'),
      it(dau, 'coin', 'Birthday coin', 24, 5, `${y - 1}-06-14`, 0, 'AED', { notes: 'Gift from grandparents' }),
    ],
    zakatLog: [], settings: { currency: 'AED' },
  });
}
async function fillSampleRates(d) {
  await Promise.all(d.items.map(async i => { const r = await rateOn(i.purchaseDate).catch(() => null); if (r) i.buyRate = { ...r, forDay: i.purchaseDate }; }));
}

/* ---------------- export ---------------- */
function download(name, text, type) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([text], { type })); a.download = name; a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
}
function exportCSV() {
  const c = cur(), q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const head = ['Owner', 'Name', 'Type', 'Purity', 'Weight g', 'Stones g', 'Pure gold g', 'Bought on', 'Shop', 'Bill no.', 'Paid', 'Paid currency', `Paid in ${c}`, `Value today ${c}`, `Gain ${c}`, 'Zakat', 'Status'];
  const rows = S.data.items.map(i => { const v = valueNow(i), cs = costIn(i);
    return [ownerById(i.ownerId)?.name, i.title, TYPES[i.type]?.label, karatLabel(i), i.weight, i.stoneWeight || 0, pureGrams(i).toFixed(3), i.purchaseDate, i.shop, i.invoiceNo,
      i.pricePaid || '', i.currency, cs != null ? cs.toFixed(2) : '', v != null ? v.toFixed(2) : '', v != null && cs ? (v - cs).toFixed(2) : '', zakatRule(i).due ? 'Zakatable' : 'Exempt', i.status]; });
  download(`My Gold ${todayStr()}.csv`, [head, ...rows].map(r => r.map(q).join(',')).join('\n'), 'text/csv');
}

/* ---------------- events ---------------- */
async function onClick(e) {
  const t = e.target;
  const tab = t.closest('[data-tab]');
  if (tab && !t.closest('.sheet')) { S.tab = tab.dataset.tab; ls.set('mg.tab', S.tab); window.scrollTo(0, 0); render(); return; }
  if (tab) { closeSheet(); S.tab = tab.dataset.tab; render(); return; }

  // filters
  const fo = t.closest('[data-fo]'); if (fo) { S.f.owner = fo.dataset.fo; render(); return; }
  const ft = t.closest('[data-ft]'); if (ft) { S.f.type = ft.dataset.ft; render(); return; }
  // form pills
  const fo2 = t.closest('[data-fo2]'); if (fo2) { readForm(); draft.ownerId = fo2.dataset.fo2; redrawForm(); return; }
  const ty = t.closest('[data-ty]'); if (ty) { readForm(); draft.type = ty.dataset.ty; if (ty.dataset.ty !== 'jewellery' && +draft.karat < 22) draft.karat = 24; redrawForm(); return; }
  const ka = t.closest('[data-ka]'); if (ka) { readForm(); draft.karat = +ka.dataset.ka; redrawForm(); return; }
  const pu = t.closest('[data-pu]'); if (pu) { readForm(); draft.purpose = pu.dataset.pu; redrawForm(); return; }
  const rmp = t.closest('[data-rmp]'); if (rmp) { readForm(); draft.photos = draft.photos.filter(p => p.id !== rmp.dataset.rmp); draft._newPhotos = draft._newPhotos.filter(p => p.id !== rmp.dataset.rmp); redrawForm(); return; }
  const rmd = t.closest('[data-rmd]'); if (rmd) { readForm(); draft._newDocs = draft._newDocs.filter(p => p.id !== rmd.dataset.rmd); redrawForm(); return; }
  const oc = t.closest('[data-oc]'); if (oc) { sheetEl.dataset.color = oc.dataset.oc; sheetEl.querySelectorAll('[data-oc]').forEach(b => { b.classList.toggle('on', b === oc); b.style.borderWidth = b === oc ? '3px' : '1px'; }); return; }
  // settings toggles
  const sc = t.closest('[data-setcur]'); if (sc) { mutate(d => { d.settings.currency = sc.dataset.setcur; }); return; }
  const th = t.closest('[data-theme]'); if (th) { ls.set('mg.theme', th.dataset.theme); mutate(d => { d.settings.theme = th.dataset.theme; }); return; }
  const zc = t.closest('[data-zc]'); if (zc) { mutate(d => { d.settings.zakatCombine = zc.dataset.zc === '1'; }); return; }
  const zm = t.closest('[data-zm]'); if (zm) { mutate(d => { d.settings.zakatMethod = zm.dataset.zm; }); return; }
  const cu = t.closest('[data-cur]'); if (cu) { S.setupCur = cu.dataset.cur; render(); return; }

  const a = t.closest('[data-a]'); if (!a) return;
  const id = a.dataset.id;
  switch (a.dataset.a) {
    case 'close': closeSheet(); break;
    case 'signin': signIn(false); break;
    case 'signin-other': signIn(true); break;
    case 'demo': {
      S.demo = true; ls.set('mg.demo', '1'); S.phase = 'loading'; render();
      const d = sampleData(); await fillSampleRates(d); ls.set('mg.demoData', JSON.stringify(d)); applyData(d, null); refreshRates(true); break;
    }
    case 'leave-demo': ls.del('mg.demo'); ls.del('mg.demoData'); S.demo = false; signIn(false); break;
    case 'add': if (!owners().length) return toast('Add a family member first.'); openForm(null); break;
    case 'open': closeSheet(true); openItem(id); break;
    case 'edit': closeSheet(true); openForm(id); break;
    case 'save': saveForm(a); break;
    case 'owner': S.f.owner = id; S.tab = 'items'; ls.set('mg.tab', 'items'); window.scrollTo(0, 0); render(); break;
    case 'refresh': refreshRates(true); pull(); break;
    case 'toggle-sold': S.f.sold = !S.f.sold; render(); break;
    case 'photo': { const it = S.data.items.find(i => i.id === sheetEl?.dataset.item); const p = it?.photos.find(x => x.id === a.dataset.pid); if (p) showPhoto(p); break; }
    case 'doc': { const it = S.data.items.find(i => i.id === sheetEl?.dataset.item); const d = it?.docs.find(x => x.id === a.dataset.did); if (d) openDoc(d); break; }
    case 'dispose': {
      const it = S.data.items.find(i => i.id === id);
      openSheet('Sold or gifted', `
        <div class="field"><span>What happened?</span><div class="pills" id="dk"><button type="button" data-dk="sold" class="on">Sold</button><button type="button" data-dk="gifted">Gifted</button></div></div>
        <div class="field"><span>Date</span><input type="date" id="sdate" value="${todayStr()}" max="${todayStr()}"></div>
        <div class="field" id="spwrap"><span>Sold for (${esc(it.currency || cur())})</span><input type="number" inputmode="decimal" id="sprice" placeholder="0"></div>
        <button class="btn" data-a="dispose-save" data-id="${id}" style="margin-top:18px">Save</button>
        <div class="small faint" style="margin-top:8px;text-align:center">It moves to “Sold & gifted” and stops counting in totals and zakat. Photos and bills are kept.</div>`);
      sheetEl.dataset.dk = 'sold';
      sheetEl.addEventListener('click', ev => { const b = ev.target.closest('[data-dk]'); if (!b) return; sheetEl.dataset.dk = b.dataset.dk;
        sheetEl.querySelectorAll('[data-dk]').forEach(x => x.classList.toggle('on', x === b)); $('#spwrap', sheetEl).style.display = b.dataset.dk === 'sold' ? '' : 'none'; });
      break;
    }
    case 'dispose-save': {
      const kind = sheetEl.dataset.dk, date = $('#sdate', sheetEl).value, price = num($('#sprice', sheetEl)?.value);
      if (await mutate(d => { const it = d.items.find(i => i.id === id); if (it) { it.status = kind; it.soldDate = date; it.soldPrice = kind === 'sold' && price > 0 ? price : 0; it.updatedAt = isoNow(); } })) { closeSheet(); toast('Updated ✓'); }
      break;
    }
    case 'unsell': if (await mutate(d => { const it = d.items.find(i => i.id === id); if (it) { it.status = 'held'; delete it.soldDate; delete it.soldPrice; } })) { closeSheet(true); openItem(id); } break;
    case 'del': {
      const it = S.data.items.find(i => i.id === id);
      if (!confirm(`Delete “${it.title}”? Its photos and bills are moved to a Removed folder in OneDrive.`)) return;
      if (await mutate(d => { d.items = d.items.filter(i => i.id !== id); })) { it.photos.forEach(p => removeFile('Photos', p)); it.docs.forEach(p => removeFile('Bills', p)); closeSheet(); toast('Deleted'); }
      break;
    }
    case 'owner-add': openOwner(null); break;
    case 'owner-edit': openOwner(id); break;
    case 'owner-save': {
      const name = $('#oname', sheetEl).value.trim(), color = sheetEl.dataset.color; if (!name) return toast('Enter a name.');
      if (await mutate(d => { const o = d.owners.find(x => x.id === id); if (o) Object.assign(o, { name, color }); else d.owners.push({ id, name, color }); })) closeSheet();
      break;
    }
    case 'owner-del': if (confirm('Remove this person?') && await mutate(d => { d.owners = d.owners.filter(o => o.id !== id); })) closeSheet(); break;
    case 'zpay': {
      const zd = zakatDate();
      openSheet('Record zakat paid', `
        <div class="field"><span>For</span><select id="zo"><option value="">Whole family</option>${owners().map(o => `<option value="${o.id}">${esc(o.name)}</option>`).join('')}</select></div>
        <div class="two"><div class="field"><span>Amount</span><input id="za" type="number" inputmode="decimal" value="${Math.round(S.data.settings.zakatCombine ? zakatFor(null).due : owners().reduce((a, o) => a + zakatFor(o.id).due, 0)) || ''}"></div>
        <div class="field"><span>Currency</span><select id="zc">${Object.keys(CURS).map(c => `<option ${c === cur() ? 'selected' : ''}>${c}</option>`).join('')}</select></div></div>
        <div class="field"><span>Date</span><input id="zdt" type="date" value="${todayStr() < zd ? todayStr() : zd}"></div>
        <div class="field"><span>Note</span><input id="zn" placeholder="e.g. Given through mosque"></div>
        <button class="btn" data-a="zpay-save" style="margin-top:18px">Save</button>`);
      break;
    }
    case 'zpay-save': {
      const amount = num($('#za', sheetEl).value); if (!(amount > 0)) return toast('Enter the amount.');
      const rec = { id: uuid(), ownerId: $('#zo', sheetEl).value, amount, currency: $('#zc', sheetEl).value, date: $('#zdt', sheetEl).value || todayStr(), note: $('#zn', sheetEl).value.trim() };
      if (await mutate(d => { d.zakatLog.push(rec); })) closeSheet();
      break;
    }
    case 'zdel': if (confirm('Delete this zakat record?')) mutate(d => { d.zakatLog = d.zakatLog.filter(z => z.id !== id); }); break;
    case 'sync': pull(true); break;
    case 'change-folder': { openSheet('Choose folder', '<div id="fb2"></div>'); folderBrowser($('#fb2', sheetEl)); break; }
    case 'csv': exportCSV(); break;
    case 'json': download(`My Gold backup ${todayStr()}.json`, stableJSON(S.data), 'application/json'); break;
    case 'signout': if (confirm(S.demo ? 'Leave the sample data?' : 'Sign out on this device? Your data stays in OneDrive.')) signOut(); break;
    case 'setup-more': collectSetup(); S.setupNames.push(''); render(); break;
    case 'setup-done': await finishSetup(a); break;
  }
}
function collectSetup() { document.querySelectorAll('[data-setup]').forEach(i => { S.setupNames[+i.dataset.setup] = i.value; }); }
async function finishSetup(btn) {
  collectSetup();
  const names = S.setupNames.map(s => s.trim()).filter(Boolean); if (!names.length) return toast('Add at least one name.');
  btn.disabled = true; btn.innerHTML = '<span class="spin" style="width:16px;height:16px"></span>';
  const d = normalize({ owners: names.map((n, i) => ({ id: uuid(), name: n, color: OWNER_COLORS[i % OWNER_COLORS.length] })), items: [], zakatLog: [],
    settings: { currency: S.setupCur || 'AED' }, createdAt: isoNow() });
  try { const etag = await putData(d, null, true); applyData(d, etag); setSync('ok'); }
  catch (e) { if (e.status === 412) { await pull(true); return; } toast(e.message); btn.disabled = false; btn.textContent = 'Start'; }
}
async function onChange(e) {
  const t = e.target;
  if (t.id === 'sort') { S.f.sort = t.value; ls.set('mg.sort', t.value); render(); return; }
  if (t.id === 'zdate') { mutate(d => { d.settings.zakatDate = t.value; }); return; }
  if (t.id === 'prem') { const v = num(t.value); mutate(d => { d.settings.premium[d.settings.currency] = isFinite(v) ? v : 0; }); return; }
  if (t.name === 'karat' || t.name === 'type') return;
  if (t.matches('[data-fp],[data-fd]') && t.files?.length) {
    readForm();
    for (const file of t.files) {
      const entry = { id: uuid(), file };
      if (/^image\//.test(file.type)) { try { entry.preview = await makeThumb(file); } catch {} }
      (t.dataset.fp ? draft._newPhotos : draft._newDocs).push(entry);
    }
    redrawForm(); return;
  }
  if (t.dataset.up && t.files?.length) {   // add photo / bill straight from the detail sheet
    const it = S.data.items.find(i => i.id === t.dataset.id); if (!it) return;
    const owner = ownerById(it.ownerId)?.name || '';
    toast('Uploading…');
    try {
      const added = [];
      for (const f of t.files) added.push(t.dataset.up === 'photo' ? await savePhoto(f, owner, it.title) : await saveDoc(f, owner, it.title, 'Bill'));
      if (await mutate(d => { const x = d.items.find(i => i.id === it.id); if (x) { (t.dataset.up === 'photo' ? x.photos : x.docs).push(...added); x.updatedAt = isoNow(); } })) { closeSheet(true); openItem(it.id); toast('Added ✓'); }
    } catch (err) { toast(err.message); }
  }
}
let qTimer;
function onInput(e) {
  if (e.target.id === 'q') { clearTimeout(qTimer); const v = e.target.value; qTimer = setTimeout(() => { S.f.q = v; const pos = e.target.selectionStart; render(); const q = $('#q'); if (q) { q.focus(); q.setSelectionRange(pos, pos); } }, 200); }
}
document.addEventListener('click', e => { if (!e.target.closest('.sheet')) onClick(e); });
document.addEventListener('change', e => { if (!e.target.closest('.sheet')) onChange(e); });
document.addEventListener('input', e => { if (!e.target.closest('.sheet')) onInput(e); });

/* ---------------- start ---------------- */
async function boot() {
  applyTheme();
  if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});
  if (S.demo) {
    const d = lsJSON('mg.demoData'); if (d) { applyData(normalize(d), null); refreshRates(); return; }
    S.demo = false; ls.del('mg.demo');
  }
  try { await handleRedirect(); } catch (e) { toast(e.message); }
  if (!tokens) { S.phase = 'welcome'; render(); return; }
  const cached = lsJSON('mg.cache');
  if (cached && LOC) { S.etag = ls.get('mg.etag') || null; applyData(normalize(cached), S.etag); }
  else { S.phase = 'loading'; render(); }
  refreshRates();
  try {
    if (!(await resolveLocation())) { S.phase = 'folder'; render(); return; }
    await pull(!cached);
    if (!S.data && S.phase !== 'setup' && S.phase !== 'needsignin') S.phase = 'setup';
    render();
  } catch (e) {
    if (e instanceof AuthError) { if (trySilent()) return; S.phase = cached ? 'ready' : 'needsignin'; render(); return; }
    if (cached) setSync('bad', e.status === 0 ? 'Offline — showing the last synced copy' : e.message);
    else { S.phase = 'welcome'; render(); toast(e.message); }
  }
}
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.phase === 'ready') { pull(); refreshRates(); } });
setInterval(() => { if (document.visibilityState === 'visible' && S.phase === 'ready') refreshRates(true); }, 5 * 60 * 1000);
window.addEventListener('scroll', () => $('header.top')?.classList.toggle('scrolled', window.scrollY > 4), { passive: true });
boot();
