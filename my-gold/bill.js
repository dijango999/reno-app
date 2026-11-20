/* My Gold — reads a gold bill (photo or PDF) and turns it into form fields.
 *
 * Text recognition runs on the phone (Tesseract.js, loaded only when a bill is scanned), so the
 * bill never leaves the device. PDFs with real text are read directly (pdf.js); scanned PDFs are
 * rendered and recognised like photos. parseBill() then pulls out the shop, bill number, date,
 * metal rates, each item (purity, weights, stones, amounts, VAT/GST) and any old gold exchanged.
 */
'use strict';

const BILL = (() => {
  const CDN = {
    tesseract: 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js',
    pdf: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js',
    pdfWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js',
  };
  const loaded = {};
  function loadScript(src) {
    return loaded[src] ||= new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.async = true;
      s.onload = res; s.onerror = () => { delete loaded[src]; rej(new Error('Could not load the bill reader. Check your internet connection.')); };
      document.head.appendChild(s);
    });
  }

  /* ---------- images ---------- */
  async function bitmap(file) {
    try { return await createImageBitmap(file, { imageOrientation: 'from-image' }); }
    catch { const u = URL.createObjectURL(file); const i = new Image(); i.src = u; await i.decode(); return i; }
  }
  /** Grey, contrast-stretched canvas about 2200px wide — what text recognition reads best. */
  function prepare(src, w0, h0) {
    const sc = Math.min(3, 2200 / w0);
    const c = document.createElement('canvas'); c.width = Math.round(w0 * sc); c.height = Math.round(h0 * sc);
    const x = c.getContext('2d', { willReadFrequently: true });
    x.drawImage(src, 0, 0, c.width, c.height);
    const d = x.getImageData(0, 0, c.width, c.height), p = d.data;
    let lo = 255, hi = 0;
    for (let i = 0; i < p.length; i += 16) { const v = (p[i] * 299 + p[i + 1] * 587 + p[i + 2] * 114) / 1000; if (v < lo) lo = v; if (v > hi) hi = v; }
    const span = Math.max(40, hi - lo);
    for (let i = 0; i < p.length; i += 4) {
      const v = (p[i] * 299 + p[i + 1] * 587 + p[i + 2] * 114) / 1000, g = Math.max(0, Math.min(255, (v - lo) * 255 / span));
      p[i] = p[i + 1] = p[i + 2] = g;
    }
    x.putImageData(d, 0, 0);
    return c;
  }
  let worker = null;
  async function ocr(canvas, onProgress, psm = '6') {
    await loadScript(CDN.tesseract);
    if (!worker) {
      worker = await Tesseract.createWorker('eng', 1, {
        logger: m => { if (m.status === 'recognizing text') onProgress?.(Math.round(m.progress * 100)); },
      });
    }
    await worker.setParameters({ tessedit_pageseg_mode: psm, preserve_interword_spaces: '1' });
    const { data } = await worker.recognize(canvas);
    return data.text || '';
  }

  /* ---------- PDFs ---------- */
  async function pdfText(file, onProgress) {
    await loadScript(CDN.pdf);
    pdfjsLib.GlobalWorkerOptions.workerSrc = CDN.pdfWorker;
    const doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const page = await doc.getPage(1);
    const tc = await page.getTextContent();
    // Rebuild lines from text pieces by their vertical position.
    const rows = [];
    for (const it of tc.items) {
      if (!it.str?.trim()) continue;
      const y = Math.round(it.transform[5]), x = it.transform[4];
      let r = rows.find(r => Math.abs(r.y - y) <= 3); if (!r) rows.push(r = { y, parts: [] });
      r.parts.push({ x, s: it.str });
    }
    const text = rows.sort((a, b) => b.y - a.y).map(r => r.parts.sort((a, b) => a.x - b.x).map(p => p.s).join(' ')).join('\n');
    if (text.replace(/\s/g, '').length > 80) return { text, image: null };
    // Scanned PDF: draw it and recognise it like a photo.
    const vp = page.getViewport({ scale: 2.5 });
    const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height;
    await page.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
    return { text: await ocr(prepare(c, c.width, c.height), onProgress), image: c };
  }

  async function read(file, onProgress) {
    let text;
    if (/pdf/i.test(file.type) || /\.pdf$/i.test(file.name || '')) text = (await pdfText(file, onProgress)).text;
    else {
      const b = await bitmap(file), w = b.width || b.naturalWidth, h = b.height || b.naturalHeight, cv = prepare(b, w, h);
      text = await ocr(cv, onProgress, '6');
      let res = parseBill(text);
      if (!res.lines.length) {   // photos of curved/skewed paper read better in automatic layout mode
        try {
          const t2 = await ocr(cv, onProgress, '4'), r2 = parseBill(t2);
          if (score(r2) > score(res)) { res = r2; text = t2; }
          else if (!res.lines.length && r2.lines.length === 0 && !res.fallback && r2.fallback) { res = r2; text = t2; }
        } catch {}
      }
      return { text, ...res };
    }
    return { text, ...parseBill(text) };
  }

  const score = r => r.lines.length * 10 + (r.fallback ? 4 : 0) + (r.date ? 1 : 0) + (r.invoiceNo ? 1 : 0) + (r.totals.net != null ? 1 : 0);


  /* ---------- optional: read the bill with Claude (needs the person's own API key) ---------- */
  const AI_MODEL = 'claude-sonnet-5-5';
  const AI_PROMPT = `You are reading a gold purchase bill / tax invoice (UAE or India). Extract EVERYTHING relevant and answer with ONE JSON object only — no prose, no markdown.
Schema:
{"shop":string,"invoiceNo":string,"date":"YYYY-MM-DD","currency":"AED"|"INR"|"USD"|"",
 "rates":{"24":number,"22":number,"21":number,"18":number},   // gold rates per gram printed on the bill, if any
 "lines":[{"title":string,"type":"coin"|"bar"|"jewellery","karat":number,"pcs":number,"gross":number,"net":number,"stones":number,"amount":number,"tax":number,"taxPct":number,"total":number}],
 "exchange":[{"title":string,"karat":number,"gross":number,"net":number,"total":number}],  // old gold / scrap given in exchange (total as positive number)
 "totals":{"sale":number|null,"tax":number|null,"net":number|null,"old":number|null},
 "making":number|null}
Rules: weights in grams (gross = with stones, net = gold weight); amounts are plain numbers without currency symbols or thousands separators; "total" of a line = amount including tax; "totals.sale" = total of the items sold, "totals.net" = final amount actually payable after any old-gold exchange. Use 0 or null when a value is not on the bill — never guess. Read the numbers digit by digit; keep decimal points exactly. Date must be the invoice date.`;
  const b64 = buf => { let s = '', a = new Uint8Array(buf); for (let i = 0; i < a.length; i += 0x8000) s += String.fromCharCode.apply(null, a.subarray(i, i + 0x8000)); return btoa(s); };
  async function aiBlock(file) {
    if (/pdf/i.test(file.type) || /\.pdf$/i.test(file.name || ''))
      return { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b64(await file.arrayBuffer()) } };
    const b = await bitmap(file), w = b.width || b.naturalWidth, h = b.height || b.naturalHeight, sc = Math.min(1, 1800 / Math.max(w, h));
    const c = document.createElement('canvas'); c.width = Math.round(w * sc); c.height = Math.round(h * sc);
    c.getContext('2d').drawImage(b, 0, 0, c.width, c.height);
    return { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: c.toDataURL('image/jpeg', 0.88).split(',')[1] } };
  }
  const num = v => { const n = typeof v === 'string' ? parseFloat(v.replace(/[^\d.\-]/g, '')) : +v; return isFinite(n) ? n : 0; };
  /** Free option: Google Gemini (key from aistudio.google.com). Tries a few model names since Google renames them. */
  async function callGemini(file, key) {
    const blk = await aiBlock(file), part = { inline_data: { mime_type: blk.source.media_type, data: blk.source.data } };
    let last = 'AI could not read the bill.';
    for (const model of ['gemini-flash-latest', 'gemini-2.5-flash', 'gemini-2.5-flash-lite']) {
      const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: 'POST', headers: { 'content-type': 'application/json', 'x-goog-api-key': key },
        body: JSON.stringify({ contents: [{ parts: [part, { text: AI_PROMPT }] }], generationConfig: { responseMimeType: 'application/json', temperature: 0 } }),
      });
      const j = await res.json().catch(() => ({}));
      if (res.ok) { const t = (j.candidates?.[0]?.content?.parts || []).map(p => p.text || '').join(''); if (t) return t; last = 'AI gave no answer.'; continue; }
      if (res.status === 400 && /API key/i.test(j.error?.message || '')) throw new Error('The AI key was not accepted — check it in Settings.');
      if (res.status === 429) { last = 'Free AI limit reached for now (try again in a minute or tomorrow).'; continue; }
      last = j.error?.message || last;
    }
    throw new Error(last);
  }
  async function callClaude(file, key) {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-api-key': key, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
      body: JSON.stringify({ model: AI_MODEL, max_tokens: 2000, messages: [{ role: 'user', content: [await aiBlock(file), { type: 'text', text: AI_PROMPT }] }] }),
    });
    const j = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(res.status === 401 ? 'The AI key was not accepted — check it in Settings.' : res.status === 429 ? 'AI is busy or out of credit.' : (j.error?.message || 'AI could not read the bill.'));
    return (j.content || []).map(c => c.text || '').join('');
  }
  async function readAI(file, key) {
    const txt = /^sk-ant-/.test(key) ? await callClaude(file, key) : await callGemini(file, key);
    const m = /\{[\s\S]*\}/.exec(txt); if (!m) throw new Error('AI gave no answer.');
    const a = JSON.parse(m[0]);
    const row = l => ({ title: String(l.title || '').trim(), karat: num(l.karat), type: ['coin', 'bar', 'jewellery'].includes(l.type) ? l.type : 'jewellery', pcs: num(l.pcs) || 1,
      gross: num(l.gross) || num(l.net), net: num(l.net) || num(l.gross), stones: num(l.stones), amount: num(l.amount), tax: num(l.tax), taxPct: num(l.taxPct), total: Math.abs(num(l.total)) || num(l.amount) });
    const t = a.totals || {}, nn = v => v == null || v === '' ? null : num(v);
    const rates = {}; for (const [k, v] of Object.entries(a.rates || {})) if (num(v) > 50) rates[+k] = num(v);
    return { text: txt, ai: true, shop: String(a.shop || ''), invoiceNo: String(a.invoiceNo || '').toUpperCase(), date: /^\d{4}-\d{2}-\d{2}$/.test(a.date || '') ? a.date : '',
      currency: ['AED', 'INR', 'USD'].includes(a.currency) ? a.currency : '', rates,
      lines: (a.lines || []).map(row).filter(l => l.gross > 0), exchange: (a.exchange || []).map(row).filter(l => l.gross > 0 || l.total > 0).map(e => ({ ...e, total: -Math.abs(e.total) })),
      totals: { sale: nn(t.sale), tax: nn(t.tax), net: nn(t.net), old: nn(t.old) }, making: nn(a.making) || undefined };
  }

  /* ---------- understanding the text ---------- */
  const SHOPS = [
    [/malabar/i, 'Malabar Gold & Diamonds'], [/joy\s*al+uk+as/i, 'Joyalukkas'], [/jos\s*al+uk+as/i, 'Jos Alukkas'],
    [/kalyan/i, 'Kalyan Jewellers'], [/tanishq/i, 'Tanishq'], [/bhima/i, 'Bhima Jewellers'], [/chemmanur/i, 'Chemmanur'],
    [/federal\s*bank/i, 'Federal Bank'], [/state\s*bank|\bsbi\b/i, 'SBI'], [/south\s*indian\s*bank/i, 'South Indian Bank'],
    [/damas/i, 'Damas'], [/pure\s*gold/i, 'Pure Gold'], [/sky\s*jewel/i, 'Sky Jewellery'], [/al\s*romaizan/i, 'Al Romaizan'],
    [/emirates\s*gold/i, 'Emirates Gold'], [/igold/i, 'iGold'], [/senco/i, 'Senco Gold'], [/caratlane/i, 'CaratLane'],
    [/muthoot/i, 'Muthoot'], [/mmtc/i, 'MMTC-PAMP'], [/lulu/i, 'Lulu'], [/thangamayil/i, 'Thangamayil'], [/grt\b/i, 'GRT Jewellers'],
  ];
  const ITEM_WORDS = /\b(ring|chain|bangle|bracelet|necklace|neck\s*lace|ear\s*rings?|studs?|jhumka|pendant|anklet|kada|mala|haram|coin|bar|biscuit|ingot|nose\s*pin|mangalsutra|thali|locket|kammal|vala|maala|minnu|tops)\b/i;
  const OLD = /\b(old\s*gold|scrap|exchange|adjust|own_|buy\s*back|purchase\s*return|urd)\b/i;

  /** Number tokens on a line, in order: "2,014.00", "(742.00)" = -742, "3.620." = 3.620, "1,00,456" = 100456. */
  function numbers(line) {
    const out = [];
    for (const raw of line.split(/\s+/)) {
      const neg = /^[-~(]|\)$/.test(raw);
      const t = raw.replace(/^[^\d-]+|[^\d]+$/g, '').replace(/^-/, '');
      if (!/^\d[\d,]*(\.\d+)?$/.test(t)) continue;
      const dec = (t.split('.')[1] || '').length, v = +t.replace(/,/g, '');
      if (!isFinite(v)) continue;
      out.push({ v: neg ? -v : v, dec, raw });
    }
    return out;
  }
  const near = (a, b, tol = 0.6) => Math.abs(a - b) <= tol;

  function parseDate(t) {
    let m = /\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](20\d{2})\b/.exec(t);
    if (m) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    m = /\b(\d{1,2})[\s\-]*(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s\-,]*(20\d{2})\b/i.exec(t);
    if (m) return `${m[3]}-${String('janfebmaraprmayjunjulaugsepoctnovdec'.indexOf(m[2].toLowerCase()) / 3 + 1).padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    m = /\b(20\d{2})[\/.\-](\d{1,2})[\/.\-](\d{1,2})\b/.exec(t);
    if (m) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
    return '';
  }
  const karatIn = s => {
    const m = /\b(24|23|22|21|20|18|14|9)\s*(?:k|kt|ct|karat|carat)\b/i.exec(s) || /\b(?:k|kt)\s*(24|22|21|18|14)\b/i.exec(s);
    if (m) return +m[1];
    if (/\b916\b/.test(s)) return 22; if (/\b(999|995|24\s*carat)\b/.test(s)) return 24; if (/\b750\b/.test(s)) return 18; if (/\b875\b/.test(s)) return 21;
    return 0;
  };
  const typeIn = s => /\b(coin|sovereign|pavan|guinea)\b/i.test(s) ? 'coin' : /\b(bar|biscuit|ingot|tola\s*bar)\b/i.test(s) ? 'bar' : 'jewellery';
  const hasStones = s => /\b(diamond|stone|ruby|emerald|pearl|polki|kundan|gem|cz|solitaire|vvs|ct\b)/i.test(s);
  function titleFrom(s) {
    const words = s.replace(/\b\d+(\.\d+)?\b|@|[|:;_=*#"'“”‘’()\[\]{}]/g, ' ').replace(/\b(\d{2}\s*k|kt|ct|pcs?|nos?|gms?|grams?|wt)\b/gi, ' ')
      .split(/\s+/).filter(w => /^[A-Za-z][A-Za-z&\-]{1,}$/.test(w) && !/^(rnd|vvs|vs|gh|dis|ref|no|sl|description|item|hsn|qty)$/i.test(w));
    const t = words.slice(0, 5).join(' ').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    return t.length >= 3 ? t : '';
  }

  function parseBill(text) {
    const lines = text.split(/\n/).map(l => l.replace(/\s+/g, ' ').trim()).filter(Boolean);
    const all = lines.join('\n');
    const r = { shop: '', invoiceNo: '', date: '', currency: '', rates: {}, lines: [], exchange: [], totals: {} };

    r.shop = (SHOPS.find(([re]) => re.test(all)) || [0, ''])[1];
    const aed = (all.match(/\bAED\b|dirham|\bUAE\b|dubai|sharjah|abu\s*dhabi|ajman|\bTRN\b|tax\s*registration|\bVAT\b/gi) || []).length;
    const inr = (all.match(/₹|\bRs\.?\b|\bINR\b|GSTIN|CGST|SGST|IGST|kerala|india|rupees/gi) || []).length;
    r.currency = aed > inr ? 'AED' : inr > aed ? 'INR' : '';

    // Bill number: first code-like token after "Invoice/Bill No"
    for (const l of lines) {
      const m = /(invoice|bill|receipt|voucher)\s*(no|number|#)\.?/i.exec(l); if (!m) continue;
      const tok = l.slice(m.index + m[0].length).split(/\s+/).map(t => t.replace(/^[^A-Z0-9]+|[^A-Z0-9]+$/gi, ''))
        .find(t => /^[A-Z0-9][A-Z0-9\-\/]{4,}$/i.test(t) && (t.match(/\d/g) || []).length >= 3);
      if (tok) { r.invoiceNo = tok.toUpperCase(); break; }
    }
    // Date: prefer the line that says date
    r.date = parseDate(lines.find(l => /date/i.test(l) && parseDate(l)) || '') || parseDate(all);

    // Metal / gold rates printed on the bill: "22K : 263.25", "Gold rate 22KT 7,150"
    for (const l of lines.filter(l => /rate/i.test(l))) {
      for (const m of l.matchAll(/\b(24|22|21|18|14)\s*(?:k|kt|ct)?\s*[:\-=]?\s*(?:rs\.?|₹|aed)?\s*(\d{1,3}(?:,\d{2,3})*(?:\.\d{1,2})?)/gi)) {
        const v = +m[2].replace(/,/g, ''); if (v > 50) r.rates[+m[1]] = v;
      }
    }

    // Item lines: weights (3 decimals) and money (2 decimals) on the same line.
    for (let i = 0; i < lines.length; i++) {
      const l = lines[i];
      if (/^(total|grand|net\s*total|sub\s*total|rounded)/i.test(l)) continue;
      let nums = numbers(l);
      const lost = n => n && n.dec === 0 && /^\d{4}$/.test(String(Math.abs(n.v)));
      const w3ish = n => n && (n.dec === 3 || lost(n));
      nums = nums.map((n, k) => lost(n) && (w3ish(nums[k - 1]) || w3ish(nums[k + 1])) ? { v: n.v / 1000, dec: 3, raw: n.raw } : n);
      const w3 = nums.filter(n => n.dec === 3 && Math.abs(n.v) < 5000);
      const m2 = nums.filter(n => n.dec === 2);
      if (!w3.length || !m2.length) continue;
      // context: this line and the 2 lines above (descriptions often sit above the numbers)
      const ctx = [lines[i - 2], lines[i - 1], l, lines[i + 1], lines[i + 2]].filter(Boolean).join(' ');
      const before = [lines[i - 2], lines[i - 1], l].filter(Boolean).join(' ');
      const named = [l, lines[i - 1], lines[i - 2]].find(x => x && ITEM_WORDS.test(x)) || '';
      if (m2.length < 2 && !named) continue;   // spec lines such as diamond details
      const neg = w3.some(n => n.v < 0) || m2.some(n => n.v < 0) || OLD.test([lines[i - 1], l].join(' '));
      // drop the rate written as "218.50 @ code"
      const at = /(\d[\d,]*\.\d{2})\s*@/.exec(l); let money = m2.map(n => Math.abs(n.v));
      if (at) { const rv = +at[1].replace(/,/g, ''); const k = money.indexOf(rv); if (k >= 0) money.splice(k, 1); }
      let ws = w3.map(n => Math.abs(n.v));
      let pcs = 1;
      if (ws.length >= 3 && Number.isInteger(ws[0]) && ws[0] <= 50) pcs = ws.shift();
      const gross = Math.max(...ws), net = ws.length > 1 ? Math.min(...ws.slice(0, 2)) : gross;
      const total = Math.max(...money);
      let tax = 0, taxPct = 0, amount = total;
      const pctN = money.find(v => [3, 5, 1.5, 0].includes(v) && v < total);
      if (pctN != null) taxPct = pctN;
      for (const a of money) for (const b of money) if (a !== b && a !== total && b !== total && b < a && near(a + b, total)) { amount = a; tax = b; }
      if (!tax && taxPct) { const t = money.find(v => v !== total && v !== taxPct && near(v, total * taxPct / (100 + taxPct), 1)); if (t) { tax = t; amount = total - t; } }
      const karat = karatIn(l) || karatIn(lines[i + 1] || '') || karatIn(named) || karatIn(lines[i - 1] || '') || karatIn(lines[i - 2] || '') || karatIn(all);
      const oldName = /diamond/i.test([l, lines[i + 1]].join(' ')) ? 'Old diamond jewellery' : 'Old gold';
      const row = { title: neg ? oldName : titleFrom(named.replace(/.*\bdescription\b/i, '')) || '', karat, type: typeIn(named || before), pcs, gross, net,
        stones: gross > net && hasStones(ctx) ? +(gross - net).toFixed(3) : 0, amount, tax, taxPct, total };
      if (neg) r.exchange.push({ ...row, total: -Math.abs(total) }); else if (gross > 0.05 && total >= 20) r.lines.push(row);
    }

    // Fallback: no tidy item row found — look for weights written near "wt / gm / grams" and the bill total.
    const tot0 = re => { const l = lines.filter(l => re.test(l) && numbers(l).length).pop(); if (!l) return null; const n = numbers(l).filter(n => n.dec === 2 || n.v >= 100); return n.length ? n[n.length - 1].v : null; };
    if (!r.lines.length) {
      let G = 0, N = 0; const any = [];
      lines.forEach(l => {
        if (OLD.test(l) && !ITEM_WORDS.test(l)) return;
        const hasG = /\b(gross|gr|g\.?\s*wt)\b/i.test(l), hasN = /\b(net|nt|n\.?\s*wt)\b/i.test(l);
        const unit = /\b(gms?|grams?|g)\b/i.test(l) || /\d\s*g\b/i.test(l), lab = hasG || hasN || /\b(wt|weight|wgt)\b/i.test(l);
        if (!lab && !unit) return;
        const ns = numbers(l).filter(n => !/\d\s*(k|kt|ct|karat|carat)$/i.test(n.raw) && n.v >= 0.1 && n.v <= 2000 && (n.dec >= 1 && n.dec <= 3 || (unit && n.dec === 0 && n.v < 500))).map(n => n.v);
        if (!ns.length) return;
        if (hasG && hasN && ns.length >= 2) { G = G || ns[0]; N = N || ns[1]; }
        else if (hasN && !hasG) N = N || ns[0];
        else if (hasG && !hasN) G = G || ns[0];
        else any.push(ns[0]);
      });
      const g0 = G || N || any[0] || 0, n0 = N || G || any[0] || 0;
      if (g0) {
        const g = g0, nt = n0;
        const total = r.totals0 = tot0(/^(net\s*(total|amount|payable|value)|grand\s*total|total\s*amount|amount\s*(payable|paid)|total\s*value|total)\b/i) ?? tot0(/\b(total|amount|price|value|paid)\b/i);
        const named = lines.find(l => ITEM_WORDS.test(l)) || '';
        r.lines.push({ title: titleFrom(named), karat: karatIn(all), type: typeIn(named || all), pcs: 1, gross: Math.max(g, nt), net: Math.min(g, nt), stones: 0, amount: total || 0, tax: 0, taxPct: 0, total: total || 0, guess: true });
        r.fallback = true;
      } else {
        // no weight at all — still offer the bill total
        const total = tot0(/^(net\s*(total|amount|payable)|grand\s*total|total\s*amount|amount\s*(payable|paid))/i);
        if (total) { r.totals.net = total; r.fallback = true; }
      }
    }

    const tot = re => {
      const l = lines.filter(l => re.test(l) && numbers(l).length).pop(); if (!l) return null;
      const all = numbers(l), last = all[all.length - 1], m2 = all.filter(n => n.dec === 2);
      if (last && last.dec === 0 && !/,/.test(last.raw) && String(Math.abs(last.v)).length >= 4 && m2.length) return last.v / 100;
      return m2.length ? m2[m2.length - 1].v : null;
    };
    const keepNet = r.totals.net;
    r.totals = { sale: tot(/^total\s*sale/i), tax: tot(/total\s*(tax|vat|gst)\s*amount|^(vat|gst)\b/i), net: tot(/^(net\s*total|grand\s*total|net\s*amount|amount\s*payable|total\s*amount)/i), old: tot(/total\s*old\s*gold/i) };
    if (r.totals.net == null && keepNet != null) r.totals.net = keepNet;
    if (r.totals.tax == null) {   // India: CGST + SGST (or IGST) on separate lines
      const g = lines.filter(l => /\b(cgst|sgst|igst|vat)\b/i.test(l) && !/gstin/i.test(l)).map(l => numbers(l).filter(n => n.dec === 2 && n.v > 0).pop()?.v || 0);
      if (g.some(Boolean)) r.totals.tax = Math.round(g.reduce((a, b) => a + b, 0) * 100) / 100;
    }
    // Making charges if listed separately
    const mk = lines.find(l => /\b(making|wastage|labour|va\s*charges?)\b/i.test(l));
    if (mk) { const n = numbers(mk).filter(n => n.dec === 2 && n.v > 0); if (n.length) r.making = n[n.length - 1].v; }
    return r;
  }

  return { read, readAI, parseBill };
})();
if (typeof module !== 'undefined') module.exports = BILL;
