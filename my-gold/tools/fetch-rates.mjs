// Reads today's retail gold rates for My Gold and writes them to rates.json.
//   UAE    — https://igold.ae/gold-rate                       (AED per gram: 24K, 22K, 21K, 18K)
//   Kerala — https://www.goodreturns.in/gold-rates/kerala.html (INR per gram: 24K, 22K, 18K)
// Runs in GitHub Actions every 30 minutes. If a site can't be read, its last good rates are kept.
// Usage: node fetch-rates.mjs path/to/rates.json
import { readFileSync, writeFileSync, existsSync } from 'node:fs';

const OUT = process.argv[2] || 'rates.json';
const UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Safari/605.1.15';
const nowIso = () => new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
const num = s => +((String(s).replace(/&#?x?\w+;/gi, ' ').match(/\d[\d,]*(?:\.\d+)?/) || ['NaN'])[0].replace(/,/g, ''));
const strip = h => h.replace(/<[^>]+>/g, ' ').replace(/&#?x?\w+;/gi, ' ').replace(/\s+/g, ' ').trim();

async function get(url) {
  const r = await fetch(url, { headers: { 'User-Agent': UA, 'Accept': 'text/html,application/xhtml+xml', 'Accept-Language': 'en-GB,en;q=0.9' } });
  if (!r.ok) throw new Error(`${url} → HTTP ${r.status}`);
  return r.text();
}
const cellsOf = row => [...row.matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)].map(m => strip(m[1]));

async function uae() {
  const html = await get('https://igold.ae/gold-rate');
  const table = [...html.matchAll(/<table[\s\S]*?<\/table>/gi)].map(m => m[0]).find(t => /24K/i.test(t) && /22K/i.test(t) && /AED/i.test(t));
  if (!table) throw new Error('igold: rate table not found');
  const rows = [...table.matchAll(/<tr[\s\S]*?<\/tr>/gi)].map(m => cellsOf(m[0]));
  const head = rows.find(r => r.some(c => /^24K$/i.test(c)));
  const vals = rows.find(r => r.length === head.length && r.every(c => /AED/i.test(c)));
  if (!head || !vals) throw new Error('igold: rows not found');
  const rates = {};
  head.forEach((h, i) => { const k = parseInt(h); if (k) rates[k] = num(vals[i]); });
  const upd = /Prices updated:\s*(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/i.exec(strip(table));
  // igold shows UAE time (UTC+4)
  const at = upd ? new Date(Date.UTC(+upd[3], +upd[2] - 1, +upd[1], +upd[4] - 4, +upd[5], +(upd[6] || 0))).toISOString().replace(/\.\d{3}Z$/, 'Z') : nowIso();
  check('igold', rates, [24, 22, 21, 18], 100, 2000);
  return { ...rates, at, source: 'igold.ae' };
}

const MONTHS = ['january','february','march','april','may','june','july','august','september','october','november','december'];
async function kerala() {
  const html = await get('https://www.goodreturns.in/gold-rates/kerala.html');
  const table = [...html.matchAll(/<table[\s\S]*?<\/table>/gi)].map(m => m[0]).find(t => /24K/.test(t) && /22K/.test(t) && /Gram/i.test(t));
  if (!table) throw new Error('goodreturns: rate table not found');
  const rows = [...table.matchAll(/<tr[\s\S]*?<\/tr>/gi)].map(m => cellsOf(m[0]));
  const head = rows.find(r => r.some(c => /^24K$/i.test(c)));
  const one = rows.find(r => r[0] === '1');
  if (!head || !one) throw new Error('goodreturns: rows not found');
  const rates = {};
  head.forEach((h, i) => { const k = parseInt(h); if (k) rates[k] = num(one[i].replace(/\(.*\)/, '')); });
  const t = /Price on (\d{1,2}) (\w+) (\d{4})/i.exec(/<title>([\s\S]*?)<\/title>/i.exec(html)?.[1] || '');
  const date = t ? `${t[3]}-${String(MONTHS.indexOf(t[2].toLowerCase()) + 1).padStart(2, '0')}-${t[1].padStart(2, '0')}` : null;
  check('goodreturns', rates, [24, 22, 18], 2000, 60000);
  return { ...rates, date, at: nowIso(), source: 'goodreturns.in' };
}
function check(name, r, ks, lo, hi) {
  for (const k of ks) if (!(r[k] > lo && r[k] < hi)) throw new Error(`${name}: ${k}K rate looks wrong (${r[k]})`);
  if (!(r[22] < r[24] && r[18] < r[22])) throw new Error(`${name}: karat rates out of order`);
}

const prev = existsSync(OUT) ? JSON.parse(readFileSync(OUT, 'utf8')) : {};
const out = { updatedAt: nowIso(), uae: prev.uae || null, kerala: prev.kerala || null, errors: {}, history: prev.history || {} };
const [u, k] = await Promise.allSettled([uae(), kerala()]);
if (u.status === 'fulfilled') out.uae = u.value; else out.errors.uae = String(u.reason?.message || u.reason);
if (k.status === 'fulfilled') out.kerala = k.value; else out.errors.kerala = String(k.reason?.message || k.reason);

// One entry per day (UAE date), so purchases can show the real local rate of that day.
const day = new Date(Date.now() + 4 * 3600e3).toISOString().slice(0, 10);
const pick = (o, ks) => o ? Object.fromEntries(ks.filter(x => o[x]).map(x => [x, o[x]])) : undefined;
out.history[day] = { ...(out.history[day] || {}), ...(u.status === 'fulfilled' ? { uae: pick(out.uae, [24, 22, 21, 18]) } : {}),
  ...(k.status === 'fulfilled' ? { kerala: pick(out.kerala, [24, 22, 18]) } : {}) };
writeFileSync(OUT, JSON.stringify(out, null, 1));
console.log(JSON.stringify({ uae: out.uae, kerala: out.kerala, errors: out.errors }, null, 1));
if (u.status === 'rejected' && k.status === 'rejected') process.exit(1);
