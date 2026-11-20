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
  async function ocr(canvas, onProgress) {
    await loadScript(CDN.tesseract);
    if (!worker) {
      worker = await Tesseract.createWorker('eng', 1, {
        logger: m => { if (m.status === 'recognizing text') onProgress?.(Math.round(m.progress * 100)); },
      });
      await worker.setParameters({ tessedit_pageseg_mode: '6', preserve_interword_spaces: '1' });
    }
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
    else { const b = await bitmap(file); const w = b.width || b.naturalWidth, h = b.height || b.naturalHeight; text = await ocr(prepare(b, w, h), onProgress); }
    return { text, ...parseBill(text) };
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

    const tot = re => {
      const l = lines.filter(l => re.test(l) && numbers(l).length).pop(); if (!l) return null;
      const all = numbers(l), last = all[all.length - 1], m2 = all.filter(n => n.dec === 2);
      if (last && last.dec === 0 && !/,/.test(last.raw) && String(Math.abs(last.v)).length >= 4 && m2.length) return last.v / 100;
      return m2.length ? m2[m2.length - 1].v : null;
    };
    r.totals = { sale: tot(/^total\s*sale/i), tax: tot(/total\s*(tax|vat|gst)\s*amount|^(vat|gst)\b/i), net: tot(/^(net\s*total|grand\s*total|net\s*amount|amount\s*payable|total\s*amount)/i), old: tot(/total\s*old\s*gold/i) };
    if (r.totals.tax == null) {   // India: CGST + SGST (or IGST) on separate lines
      const g = lines.filter(l => /\b(cgst|sgst|igst|vat)\b/i.test(l) && !/gstin/i.test(l)).map(l => numbers(l).filter(n => n.dec === 2 && n.v > 0).pop()?.v || 0);
      if (g.some(Boolean)) r.totals.tax = Math.round(g.reduce((a, b) => a + b, 0) * 100) / 100;
    }
    // Making charges if listed separately
    const mk = lines.find(l => /\b(making|wastage|labour|va\s*charges?)\b/i.test(l));
    if (mk) { const n = numbers(mk).filter(n => n.dec === 2 && n.v > 0); if (n.length) r.making = n[n.length - 1].v; }
    return r;
  }

  return { read, parseBill };
})();
if (typeof module !== 'undefined') module.exports = BILL;
