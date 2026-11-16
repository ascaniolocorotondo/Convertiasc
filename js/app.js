/* Convertiasc — convertitore di file che gira tutto nel browser. */
'use strict';

/* =========================================================
   Utilità
   ========================================================= */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const sleep = (ms = 0) => new Promise(r => setTimeout(r, ms));
const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const isObj = v => v !== null && typeof v === 'object' && !Array.isArray(v);
const baseName = n => n.replace(/\.[^.]+$/, '') || 'file';
const fmtSize = b => b < 1024 ? b + ' B' : b < 1048576 ? (b / 1024).toFixed(1) + ' KB' : (b / 1048576).toFixed(1) + ' MB';

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 60000);
}

/* Librerie esterne caricate solo quando servono */
const LIBS = {
  xlsx: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
  jspdf: 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js',
  pdfjs: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
  pdfjsWorker: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',
  mammoth: 'https://cdn.jsdelivr.net/npm/mammoth@1.6.0/mammoth.browser.min.js',
  marked: 'https://cdn.jsdelivr.net/npm/marked@9.1.6/marked.min.js',
  turndown: 'https://cdn.jsdelivr.net/npm/turndown@7.1.2/dist/turndown.js',
  yaml: 'https://cdnjs.cloudflare.com/ajax/libs/js-yaml/4.1.0/js-yaml.min.js',
  lame: 'https://cdn.jsdelivr.net/npm/lamejs@1.2.1/lame.min.js',
  jszip: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
};
const libPromises = {};
function lib(name) {
  if (!libPromises[name]) {
    libPromises[name] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = LIBS[name];
      s.onload = resolve;
      s.onerror = () => {
        delete libPromises[name];
        s.remove();
        reject(new Error(`Impossibile caricare il modulo «${name}». Serve una connessione a internet.`));
      };
      document.head.append(s);
    });
  }
  return libPromises[name];
}

/* =========================================================
   Formati
   ========================================================= */
const EXT_ALIAS = { jpeg: 'jpg', jpe: 'jpg', htm: 'html', yml: 'yaml', markdown: 'md', tif: 'tiff', oga: 'ogg', heif: 'heic' };

const CATS = {
  image: { label: 'Immagine', ext: ['png', 'jpg', 'webp', 'bmp', 'gif', 'ico', 'svg', 'avif', 'heic', 'tiff'], out: ['png', 'jpg', 'webp', 'bmp', 'ico', 'pdf'] },
  pdf:   { label: 'PDF', ext: ['pdf'], out: ['png', 'jpg', 'txt', 'md', 'html'] },
  doc:   { label: 'Documento', ext: ['docx', 'md', 'html', 'txt'], out: ['pdf', 'doc', 'html', 'md', 'txt'] },
  data:  { label: 'Dati / Foglio', ext: ['csv', 'tsv', 'xlsx', 'xls', 'ods', 'json', 'xml', 'yaml'], out: ['xlsx', 'csv', 'json', 'tsv', 'ods', 'xml', 'yaml', 'html', 'md'] },
  audio: { label: 'Audio', ext: ['mp3', 'wav', 'ogg', 'm4a', 'aac', 'flac', 'opus', 'weba'], out: ['mp3', 'wav', 'rec'] },
  video: { label: 'Video', ext: ['mp4', 'webm', 'mov', 'mkv', 'm4v', 'ogv'], out: ['mp3', 'wav', 'rec', 'png', 'jpg'] },
};

const FMT_LABEL = {
  png: 'PNG', jpg: 'JPG', webp: 'WEBP', bmp: 'BMP', ico: 'ICO (icona)', pdf: 'PDF',
  txt: 'TXT', md: 'Markdown', html: 'HTML', doc: 'DOC (Word)',
  xlsx: 'XLSX (Excel)', csv: 'CSV', tsv: 'TSV', ods: 'ODS (LibreOffice)', json: 'JSON', xml: 'XML', yaml: 'YAML',
  mp3: 'MP3', wav: 'WAV',
};

function recFormat(kind) {
  if (!window.MediaRecorder) return null;
  const list = kind === 'video'
    ? [['video/webm;codecs=vp9,opus', 'webm'], ['video/webm;codecs=vp8,opus', 'webm'], ['video/webm', 'webm'], ['video/mp4', 'mp4']]
    : [['audio/webm;codecs=opus', 'webm'], ['audio/ogg;codecs=opus', 'ogg'], ['audio/mp4', 'm4a']];
  for (const [mime, ext] of list) if (MediaRecorder.isTypeSupported(mime)) return { mime, ext };
  return null;
}
const REC = { audio: recFormat('audio'), video: recFormat('video') };

function fmtLabel(fmt, cat) {
  if (fmt === 'rec') return `${REC[cat].ext.toUpperCase()} (in tempo reale)`;
  if (cat === 'video' && (fmt === 'png' || fmt === 'jpg')) return `${fmt.toUpperCase()} (fotogramma)`;
  if (cat === 'pdf' && (fmt === 'png' || fmt === 'jpg')) return `${fmt.toUpperCase()} (una per pagina)`;
  if (cat === 'video' && (fmt === 'mp3' || fmt === 'wav')) return `${fmt.toUpperCase()} (solo audio)`;
  return FMT_LABEL[fmt] || fmt.toUpperCase();
}

function detect(file) {
  let ext = (file.name.split('.').pop() || '').toLowerCase();
  ext = EXT_ALIAS[ext] || ext;
  for (const [cat, def] of Object.entries(CATS)) if (def.ext.includes(ext)) return { cat, ext };
  // fallback sul MIME type
  const t = file.type || '';
  if (t.startsWith('image/')) return { cat: 'image', ext };
  if (t.startsWith('audio/')) return { cat: 'audio', ext };
  if (t.startsWith('video/')) return { cat: 'video', ext };
  if (t === 'text/plain') return { cat: 'doc', ext: 'txt' };
  return { cat: null, ext };
}

function outputsFor(cat, ext) {
  if (!cat) return [];
  return CATS[cat].out.filter(o => {
    if (o === 'rec') return !!REC[cat];
    if (cat === 'image') return true;      // png→png è utile per ridimensionare / rimuovere bianco
    if (cat === 'data' && ext === 'xls' && o === 'xlsx') return true;
    return o !== ext;
  });
}

/* =========================================================
   Immagini
   ========================================================= */
async function fileToCanvas(file, ext, maxW) {
  const blob = ext === 'svg' && !file.type ? new Blob([file], { type: 'image/svg+xml' }) : file;
  const url = URL.createObjectURL(blob);
  const img = new Image();
  img.src = url;
  try {
    await img.decode();
  } catch {
    URL.revokeObjectURL(url);
    const tip = ['heic', 'tiff'].includes(ext) ? ' Prova con Safari, che supporta HEIC e TIFF.' : '';
    throw new Error('Questo browser non riesce ad aprire l\'immagine.' + tip);
  }
  let w = img.naturalWidth || 1024, h = img.naturalHeight || 1024;
  if (maxW && w > maxW) { h = Math.max(1, Math.round(h * maxW / w)); w = maxW; }
  const c = document.createElement('canvas');
  c.width = w; c.height = h;
  c.getContext('2d').drawImage(img, 0, 0, w, h);
  URL.revokeObjectURL(url);
  return c;
}

function flatten(c, color = '#ffffff') {
  const o = document.createElement('canvas');
  o.width = c.width; o.height = c.height;
  const g = o.getContext('2d');
  g.fillStyle = color;
  g.fillRect(0, 0, o.width, o.height);
  g.drawImage(c, 0, 0);
  return o;
}

function canvasBlob(c, type, q) {
  return new Promise((res, rej) => c.toBlob(b => {
    if (!b) return rej(new Error('Impossibile creare l\'immagine.'));
    if (b.type !== type) return rej(new Error(`Questo browser non sa creare file ${type.split('/')[1].toUpperCase()}.`));
    res(b);
  }, type, q));
}

function canvasToBMP(c) {
  const w = c.width, h = c.height;
  const d = flatten(c).getContext('2d').getImageData(0, 0, w, h).data;
  const row = Math.ceil(w * 3 / 4) * 4;
  const size = 54 + row * h;
  const buf = new ArrayBuffer(size);
  const v = new DataView(buf);
  v.setUint16(0, 0x4D42, true);          // "BM"
  v.setUint32(2, size, true);
  v.setUint32(10, 54, true);
  v.setUint32(14, 40, true);
  v.setInt32(18, w, true);
  v.setInt32(22, h, true);
  v.setUint16(26, 1, true);
  v.setUint16(28, 24, true);
  v.setUint32(34, row * h, true);
  v.setInt32(38, 2835, true);
  v.setInt32(42, 2835, true);
  const u = new Uint8Array(buf);
  for (let y = 0; y < h; y++) {
    let o = 54 + (h - 1 - y) * row;
    for (let x = 0; x < w; x++) {
      const i = (y * w + x) * 4;
      u[o++] = d[i + 2]; u[o++] = d[i + 1]; u[o++] = d[i];
    }
  }
  return new Blob([buf], { type: 'image/bmp' });
}

async function canvasToICO(c) {
  const sizes = [16, 32, 48, 64, 128, 256].filter(s => s <= Math.max(c.width, c.height) || s <= 32);
  const pngs = [];
  for (const s of sizes) {
    const k = document.createElement('canvas');
    k.width = k.height = s;
    const g = k.getContext('2d');
    g.imageSmoothingQuality = 'high';
    const r = Math.min(s / c.width, s / c.height);
    const w = c.width * r, h = c.height * r;
    g.drawImage(c, (s - w) / 2, (s - h) / 2, w, h);
    pngs.push(new Uint8Array(await (await canvasBlob(k, 'image/png')).arrayBuffer()));
  }
  const head = new DataView(new ArrayBuffer(6 + 16 * sizes.length));
  head.setUint16(2, 1, true);
  head.setUint16(4, sizes.length, true);
  let offset = 6 + 16 * sizes.length;
  sizes.forEach((s, i) => {
    const e = 6 + i * 16;
    head.setUint8(e, s >= 256 ? 0 : s);
    head.setUint8(e + 1, s >= 256 ? 0 : s);
    head.setUint16(e + 4, 1, true);
    head.setUint16(e + 6, 32, true);
    head.setUint32(e + 8, pngs[i].length, true);
    head.setUint32(e + 12, offset, true);
    offset += pngs[i].length;
  });
  return new Blob([head.buffer, ...pngs], { type: 'image/x-icon' });
}

async function canvasesToPdf(canvases, q, keepAlpha) {
  await lib('jspdf');
  const { jsPDF } = window.jspdf;
  let doc;
  for (const c of canvases) {
    const w = c.width * 0.75, h = c.height * 0.75;   // px → pt (96 dpi)
    const ori = w > h ? 'l' : 'p';
    if (!doc) doc = new jsPDF({ unit: 'pt', format: [w, h], orientation: ori, compress: true });
    else doc.addPage([w, h], ori);
    if (keepAlpha) doc.addImage(c.toDataURL('image/png'), 'PNG', 0, 0, w, h);
    else doc.addImage(flatten(c).toDataURL('image/jpeg', q), 'JPEG', 0, 0, w, h);
  }
  return doc.output('blob');
}

async function imageCanvas(it, o) {
  let c = await fileToCanvas(it.file, it.ext, o.maxW);
  if (o.rw) c = applyRemoveWhite(c);
  return c;
}

async function convImage(it, o) {
  const c = await imageCanvas(it, o);
  const base = baseName(it.file.name);
  switch (it.out) {
    case 'png':  return [{ name: base + '.png', blob: await canvasBlob(c, 'image/png') }];
    case 'jpg':  return [{ name: base + '.jpg', blob: await canvasBlob(flatten(c), 'image/jpeg', o.q) }];
    case 'webp': return [{ name: base + '.webp', blob: await canvasBlob(c, 'image/webp', o.q) }];
    case 'bmp':  return [{ name: base + '.bmp', blob: canvasToBMP(c) }];
    case 'ico':  return [{ name: base + '.ico', blob: await canvasToICO(c) }];
    case 'pdf':  return [{ name: base + '.pdf', blob: await canvasesToPdf([c], o.q, o.rw) }];
  }
}

/* =========================================================
   Rimuovi bianco
   ========================================================= */
function hexToRgb(hex) {
  const n = parseInt(hex.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

/**
 * Rende trasparente il colore scelto (di default il bianco).
 * tol, soft: 0–100. mode: "edges" (solo lo sfondo collegato ai bordi) o "all".
 */
function removeColor(src, { tol, soft, mode, defringe, color }) {
  const w = src.width, h = src.height, s = src.data;
  const out = new ImageData(new Uint8ClampedArray(s), w, h);
  const d = out.data;
  const [tr, tg, tb] = color;
  const T = tol * 2.55, S = soft * 2.55, lim = T + S;
  const dist = i => Math.max(Math.abs(s[i] - tr), Math.abs(s[i + 1] - tg), Math.abs(s[i + 2] - tb));

  let mask = null;
  if (mode === 'edges') {
    // riempimento a partire dai bordi: tocca solo lo sfondo, non il bianco interno
    mask = new Uint8Array(w * h);
    const stack = [];
    const push = p => {
      if (!mask[p] && s[p * 4 + 3] > 0 && dist(p * 4) <= lim) { mask[p] = 1; stack.push(p); }
      else if (!mask[p] && s[p * 4 + 3] === 0) { mask[p] = 2; stack.push(p); } // già trasparente: attraversabile
    };
    for (let x = 0; x < w; x++) { push(x); push((h - 1) * w + x); }
    for (let y = 0; y < h; y++) { push(y * w); push(y * w + w - 1); }
    while (stack.length) {
      const p = stack.pop(), x = p % w;
      if (x > 0) push(p - 1);
      if (x < w - 1) push(p + 1);
      if (p >= w) push(p - w);
      if (p < w * (h - 1)) push(p + w);
    }
  }

  for (let p = 0, i = 0; p < w * h; p++, i += 4) {
    if (mask && mask[p] !== 1) continue;
    const dd = dist(i);
    if (dd > lim) continue;
    const a = dd <= T ? 0 : (dd - T) / S;      // quanto del pixel resta visibile (0–1)
    if (a <= 0) { d[i + 3] = 0; continue; }
    if (defringe) {
      // toglie il "velo" del colore rimosso dai pixel semitrasparenti
      d[i]     = (s[i]     - tr * (1 - a)) / a;
      d[i + 1] = (s[i + 1] - tg * (1 - a)) / a;
      d[i + 2] = (s[i + 2] - tb * (1 - a)) / a;
    }
    d[i + 3] = s[i + 3] * a;
  }
  return out;
}

function rwSettings() {
  return {
    tol: +$('#tol').value,
    soft: +$('#soft').value,
    mode: $('input[name=rwMode]:checked').value,
    defringe: $('#defringe').checked,
    color: hexToRgb($('#rwColor').value),
  };
}

function applyRemoveWhite(c) {
  const g = c.getContext('2d');
  const data = removeColor(g.getImageData(0, 0, c.width, c.height), rwSettings());
  g.putImageData(data, 0, 0);
  return c;
}

/* =========================================================
   PDF in ingresso
   ========================================================= */
async function openPdf(file) {
  await lib('pdfjs');
  pdfjsLib.GlobalWorkerOptions.workerSrc = LIBS.pdfjsWorker;
  return pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
}

async function pdfText(pdf, progress) {
  const pages = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const tc = await page.getTextContent();
    pages.push(tc.items.map(t => t.str + (t.hasEOL ? '\n' : '')).join('').replace(/[ \t]+\n/g, '\n').trim());
    progress(n / pdf.numPages);
  }
  return pages;
}

async function convPdf(it, o, progress) {
  const pdf = await openPdf(it.file);
  const base = baseName(it.file.name);
  if (it.out === 'png' || it.out === 'jpg') {
    const res = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const page = await pdf.getPage(n);
      const vp = page.getViewport({ scale: 2 });
      const c = document.createElement('canvas');
      c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
      const g = c.getContext('2d');
      g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: g, viewport: vp }).promise;
      const blob = it.out === 'png' ? await canvasBlob(c, 'image/png') : await canvasBlob(c, 'image/jpeg', o.q);
      const suffix = pdf.numPages > 1 ? `-pagina-${String(n).padStart(String(pdf.numPages).length, '0')}` : '';
      res.push({ name: `${base}${suffix}.${it.out}`, blob });
      progress(n / pdf.numPages);
    }
    return res;
  }
  const pages = await pdfText(pdf, progress);
  if (it.out === 'txt') return [textFile(base + '.txt', pages.join('\n\n'))];
  if (it.out === 'md') return [textFile(base + '.md', pages.map((p, i) => `## Pagina ${i + 1}\n\n${p}`).join('\n\n'))];
  const html = pages.map((p, i) => `<section><h2>Pagina ${i + 1}</h2>\n${textToHtml(p)}</section>`).join('\n');
  return [textFile(base + '.html', htmlPage(base, html), 'text/html')];
}

/* =========================================================
   Documenti (DOCX, Markdown, HTML, TXT)
   ========================================================= */
const textFile = (name, text, type = 'text/plain') => ({ name, blob: new Blob([text], { type: type + ';charset=utf-8' }) });
const textToHtml = t => t.split(/\n{2,}/).filter(p => p.trim()).map(p => `<p>${esc(p).replace(/\n/g, '<br>')}</p>`).join('\n');
const htmlPage = (title, body) => `<!doctype html>
<html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(title)}</title>
<style>body{font:16px/1.6 system-ui,sans-serif;max-width:800px;margin:40px auto;padding:0 20px;color:#222}
table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:6px 10px}img{max-width:100%}</style>
</head><body>
${body}
</body></html>`;

async function docToHtml(file, ext) {
  if (ext === 'docx') {
    await lib('mammoth');
    const r = await mammoth.convertToHtml({ arrayBuffer: await file.arrayBuffer() });
    return r.value;
  }
  const text = await file.text();
  if (ext === 'md') { await lib('marked'); return marked.parse(text); }
  if (ext === 'html') {
    const d = new DOMParser().parseFromString(text, 'text/html');
    d.querySelectorAll('script, style, noscript').forEach(e => e.remove());
    return d.body.innerHTML;
  }
  return textToHtml(text);
}

function htmlToText(html) {
  const d = new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html');
  d.querySelectorAll('br').forEach(b => b.replaceWith('\n'));
  d.querySelectorAll('li').forEach(li => { li.prepend('• '); li.append('\n'); });
  d.querySelectorAll('td, th').forEach(c => c.append('\t'));
  d.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, tr, pre, blockquote, ul, ol, table, section').forEach(e => e.append('\n\n'));
  return d.body.textContent.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

async function htmlToMd(html) {
  await lib('turndown');
  return new TurndownService({ headingStyle: 'atx', codeBlockStyle: 'fenced', bulletListMarker: '-' }).turndown(html);
}

/* PDF "vero" con titoli, elenchi e paragrafi */
async function htmlToPdf(html) {
  await lib('jspdf');
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const M = 56, W = doc.internal.pageSize.getWidth() - M * 2, H = doc.internal.pageSize.getHeight();
  let y = M;
  const SIZES = { H1: 22, H2: 17, H3: 14, H4: 12.5, H5: 12, H6: 12 };

  const write = (text, { size = 11, bold = false, indent = 0, gap = 8, mono = false } = {}) => {
    text = text.replace(/\s+\n/g, '\n').trim();
    if (!text) return;
    doc.setFont(mono ? 'courier' : 'helvetica', bold ? 'bold' : 'normal');
    doc.setFontSize(size);
    const lh = size * 1.4;
    for (const line of doc.splitTextToSize(text, W - indent)) {
      if (y + lh > H - M) { doc.addPage(); y = M; }
      doc.text(line, M + indent, y + size);
      y += lh;
    }
    y += gap;
  };

  const walk = el => {
    for (const n of el.childNodes) {
      if (n.nodeType === 3) { write(n.textContent); continue; }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (SIZES[tag]) { y += 6; write(n.textContent, { size: SIZES[tag], bold: true, gap: 6 }); }
      else if (tag === 'UL' || tag === 'OL') {
        [...n.children].forEach((li, i) => write((tag === 'OL' ? `${i + 1}. ` : '•  ') + htmlToText(li.innerHTML), { indent: 14, gap: 3 }));
        y += 6;
      }
      else if (tag === 'PRE') write(n.textContent, { mono: true, size: 9.5 });
      else if (tag === 'TABLE') {
        for (const tr of n.querySelectorAll('tr')) write([...tr.children].map(c => c.textContent.trim()).join('   |   '), { size: 10, gap: 2, bold: !!tr.querySelector('th') });
        y += 8;
      }
      else if (tag === 'BLOCKQUOTE') write(n.textContent, { indent: 18 });
      else if (['DIV', 'SECTION', 'ARTICLE', 'MAIN', 'HEADER', 'FOOTER', 'BODY'].includes(tag)) walk(n);
      else if (tag === 'HR') y += 12;
      else write(n.textContent, { bold: tag === 'STRONG' || tag === 'B' });
    }
  };
  walk(new DOMParser().parseFromString(`<body>${html}</body>`, 'text/html').body);
  return doc.output('blob');
}

function htmlToDoc(title, html) {
  const page = `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="utf-8"><title>${esc(title)}</title></head><body>${html}</body></html>`;
  return new Blob(['\ufeff', page], { type: 'application/msword' });
}

async function convDoc(it) {
  const base = baseName(it.file.name);
  const html = await docToHtml(it.file, it.ext);
  switch (it.out) {
    case 'html': return [textFile(base + '.html', htmlPage(base, html), 'text/html')];
    case 'txt':  return [textFile(base + '.txt', htmlToText(html))];
    case 'md':   return [textFile(base + '.md', await htmlToMd(html), 'text/markdown')];
    case 'pdf':  return [{ name: base + '.pdf', blob: await htmlToPdf(html) }];
    case 'doc':  return [{ name: base + '.doc', blob: htmlToDoc(base, html) }];
  }
}

/* =========================================================
   Dati (CSV, Excel, JSON, XML, YAML…)
   ========================================================= */
function xmlToObj(text) {
  const d = new DOMParser().parseFromString(text, 'application/xml');
  if (d.querySelector('parsererror')) throw new Error('Il file XML non è valido.');
  const node = el => {
    const obj = {};
    for (const a of el.attributes) obj['@' + a.name] = a.value;
    const kids = [...el.children];
    if (!kids.length) {
      const t = el.textContent.trim();
      if (!el.attributes.length) return t;
      if (t) obj['#text'] = t;
      return obj;
    }
    for (const k of kids) {
      const v = node(k);
      if (k.tagName in obj) {
        if (!Array.isArray(obj[k.tagName])) obj[k.tagName] = [obj[k.tagName]];
        obj[k.tagName].push(v);
      } else obj[k.tagName] = v;
    }
    return obj;
  };
  return { [d.documentElement.tagName]: node(d.documentElement) };
}

function objToXml(tree) {
  const tag = n => { let t = String(n).replace(/[^\w.-]/g, '_'); if (!/^[A-Za-z_]/.test(t)) t = '_' + t; return t; };
  const build = (name, val, ind) => {
    const t = tag(name);
    if (Array.isArray(val)) return val.map(v => build(name, v, ind)).join('');
    if (!isObj(val)) return `${ind}<${t}>${esc(val ?? '')}</${t}>\n`;
    let attrs = '', inner = '', text = '';
    for (const [k, v] of Object.entries(val)) {
      if (k.startsWith('@')) attrs += ` ${tag(k.slice(1))}="${esc(v)}"`;
      else if (k === '#text') text = esc(v);
      else inner += build(k, v, ind + '  ');
    }
    if (!inner) return `${ind}<${t}${attrs}>${text}</${t}>\n`;
    return `${ind}<${t}${attrs}>${text}\n${inner}${ind}</${t}>\n`;
  };
  let body;
  if (isObj(tree) && Object.keys(tree).length === 1 && !Array.isArray(Object.values(tree)[0])) body = build(Object.keys(tree)[0], Object.values(tree)[0], '');
  else if (Array.isArray(tree)) body = `<root>\n${tree.map(v => build('item', v, '  ')).join('')}</root>\n`;
  else body = build('root', tree, '');
  return `<?xml version="1.0" encoding="UTF-8"?>\n${body}`;
}

function flattenObj(o, prefix = '', out = {}) {
  for (const [k, v] of Object.entries(o)) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (isObj(v)) flattenObj(v, key, out);
    else if (Array.isArray(v)) out[key] = v.every(x => !isObj(x) && !Array.isArray(x)) ? v.join(', ') : JSON.stringify(v);
    else out[key] = v;
  }
  return out;
}

function toRows(t) {
  if (Array.isArray(t)) return t.map(v => isObj(v) ? flattenObj(v) : { valore: v });
  if (isObj(t)) {
    const keys = Object.keys(t);
    if (keys.length === 1 && (Array.isArray(t[keys[0]]) || isObj(t[keys[0]]))) return toRows(t[keys[0]]);
    const arrKey = keys.find(k => Array.isArray(t[k]) && t[k].some(isObj));
    if (arrKey) return toRows(t[arrKey]);
    return [flattenObj(t)];
  }
  return [{ valore: t }];
}

async function parseData(file, ext) {
  if (ext === 'json') {
    try { return { tree: JSON.parse(await file.text()) }; }
    catch { throw new Error('Il file JSON non è valido.'); }
  }
  if (ext === 'yaml') { await lib('yaml'); return { tree: jsyaml.load(await file.text()) }; }
  if (ext === 'xml') return { tree: xmlToObj(await file.text()) };
  await lib('xlsx');
  const wb = ext === 'csv' || ext === 'tsv'
    ? XLSX.read(await file.text(), { type: 'string', FS: ext === 'tsv' ? '\t' : undefined })
    : XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellDates: true });
  return { wb };
}

async function convData(it) {
  const base = baseName(it.file.name);
  const parsed = await parseData(it.file, it.ext);
  const sheetRows = (wb, name) => XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });

  const getTree = () => {
    if (parsed.tree !== undefined) return parsed.tree;
    const names = parsed.wb.SheetNames;
    if (names.length === 1) return sheetRows(parsed.wb, names[0]);
    return Object.fromEntries(names.map(n => [n, sheetRows(parsed.wb, n)]));
  };
  const getWb = async () => {
    await lib('xlsx');
    if (parsed.wb) return parsed.wb;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(toRows(parsed.tree)), 'Foglio1');
    return wb;
  };
  const perSheet = async (fn, ext, type) => {
    const wb = await getWb();
    return wb.SheetNames.map(n => textFile(
      wb.SheetNames.length > 1 ? `${base}-${n}.${ext}` : `${base}.${ext}`,
      fn(wb.Sheets[n], n), type));
  };

  switch (it.out) {
    case 'json': return [textFile(base + '.json', JSON.stringify(getTree(), null, 2), 'application/json')];
    case 'yaml': await lib('yaml'); return [textFile(base + '.yaml', jsyaml.dump(getTree(), { lineWidth: -1 }), 'text/yaml')];
    case 'xml':  return [textFile(base + '.xml', objToXml(getTree()), 'application/xml')];
    case 'csv':  return perSheet(s => '\ufeff' + XLSX.utils.sheet_to_csv(s), 'csv', 'text/csv');
    case 'tsv':  return perSheet(s => '\ufeff' + XLSX.utils.sheet_to_csv(s, { FS: '\t' }), 'tsv', 'text/tab-separated-values');
    case 'html': return perSheet((s, n) => htmlPage(n, XLSX.utils.sheet_to_html(s, { header: '', footer: '' })), 'html', 'text/html');
    case 'md':   return perSheet(s => rowsToMd(XLSX.utils.sheet_to_json(s, { header: 1, defval: '' })), 'md', 'text/markdown');
    case 'xlsx':
    case 'ods': {
      const wb = await getWb();
      const data = XLSX.write(wb, { bookType: it.out, type: 'array' });
      return [{ name: `${base}.${it.out}`, blob: new Blob([data], { type: 'application/octet-stream' }) }];
    }
  }
}

function rowsToMd(rows) {
  if (!rows.length) return '';
  const cols = Math.max(...rows.map(r => r.length));
  const cell = v => String(v ?? '').replace(/\|/g, '\\|').replace(/\n/g, ' ');
  const line = r => '| ' + Array.from({ length: cols }, (_, i) => cell(r[i])).join(' | ') + ' |';
  return [line(rows[0]), '| ' + Array(cols).fill('---').join(' | ') + ' |', ...rows.slice(1).map(line)].join('\n') + '\n';
}

/* =========================================================
   Audio e video
   ========================================================= */
async function decodeAudio(file) {
  const Ctx = window.AudioContext || window.webkitAudioContext;
  const ctx = new Ctx();
  try {
    return await ctx.decodeAudioData(await file.arrayBuffer());
  } catch {
    throw new Error('Questo browser non riesce a leggere l\'audio di questo file.');
  } finally {
    ctx.close();
  }
}

function encodeWAV(buf) {
  const ch = buf.numberOfChannels, sr = buf.sampleRate, n = buf.length;
  const size = n * ch * 2;
  const ab = new ArrayBuffer(44 + size);
  const v = new DataView(ab);
  const str = (o, s) => { for (let i = 0; i < s.length; i++) v.setUint8(o + i, s.charCodeAt(i)); };
  str(0, 'RIFF'); v.setUint32(4, 36 + size, true); str(8, 'WAVE');
  str(12, 'fmt '); v.setUint32(16, 16, true); v.setUint16(20, 1, true); v.setUint16(22, ch, true);
  v.setUint32(24, sr, true); v.setUint32(28, sr * ch * 2, true); v.setUint16(32, ch * 2, true); v.setUint16(34, 16, true);
  str(36, 'data'); v.setUint32(40, size, true);
  const chans = Array.from({ length: ch }, (_, i) => buf.getChannelData(i));
  let o = 44;
  for (let i = 0; i < n; i++) {
    for (let c = 0; c < ch; c++) {
      const s = Math.max(-1, Math.min(1, chans[c][i]));
      v.setInt16(o, s < 0 ? s * 0x8000 : s * 0x7FFF, true);
      o += 2;
    }
  }
  return new Blob([ab], { type: 'audio/wav' });
}

async function encodeMP3(buf, kbps, progress) {
  await lib('lame');
  const ch = Math.min(2, buf.numberOfChannels);
  const enc = new lamejs.Mp3Encoder(ch, buf.sampleRate, kbps);
  const toI16 = f => {
    const o = new Int16Array(f.length);
    for (let i = 0; i < f.length; i++) { const s = Math.max(-1, Math.min(1, f[i])); o[i] = s < 0 ? s * 0x8000 : s * 0x7FFF; }
    return o;
  };
  const L = toI16(buf.getChannelData(0));
  const R = ch > 1 ? toI16(buf.getChannelData(1)) : null;
  const parts = [];
  const B = 1152;
  for (let i = 0, k = 0; i < L.length; i += B, k++) {
    const mp3 = R ? enc.encodeBuffer(L.subarray(i, i + B), R.subarray(i, i + B)) : enc.encodeBuffer(L.subarray(i, i + B));
    if (mp3.length) parts.push(new Uint8Array(mp3));
    if (k % 400 === 0) { progress(i / L.length); await sleep(); }
  }
  const end = enc.flush();
  if (end.length) parts.push(new Uint8Array(end));
  return new Blob(parts, { type: 'audio/mpeg' });
}

async function recordAudioBuffer(buf, rec, progress) {
  const ctx = new (window.AudioContext || window.webkitAudioContext)();
  const src = ctx.createBufferSource();
  src.buffer = buf;
  const dest = ctx.createMediaStreamDestination();
  src.connect(dest);
  const mr = new MediaRecorder(dest.stream, { mimeType: rec.mime });
  const chunks = [];
  mr.ondataavailable = e => e.data.size && chunks.push(e.data);
  const stopped = new Promise(r => (mr.onstop = r));
  const t0 = ctx.currentTime;
  const timer = setInterval(() => progress((ctx.currentTime - t0) / buf.duration), 250);
  mr.start(1000);
  src.start();
  await new Promise(r => (src.onended = r));
  clearInterval(timer);
  mr.stop();
  await stopped;
  ctx.close();
  return new Blob(chunks, { type: rec.mime.split(';')[0] });
}

function loadVideo(file) {
  const v = document.createElement('video');
  v.src = URL.createObjectURL(file);
  v.playsInline = true;
  v.preload = 'auto';
  Object.assign(v.style, { position: 'fixed', left: '-9999px', width: '2px', height: '2px', opacity: '0' });
  document.body.append(v);
  return new Promise((res, rej) => {
    v.onloadeddata = () => res(v);
    v.onerror = () => { cleanupVideo(v); rej(new Error('Questo browser non riesce a leggere questo video.')); };
  });
}
function cleanupVideo(v) { URL.revokeObjectURL(v.src); v.remove(); }

async function videoFrame(it, o) {
  const v = await loadVideo(it.file);
  try {
    v.currentTime = Math.min(1, (v.duration || 0) * 0.1);
    await new Promise(r => (v.onseeked = r));
    const c = document.createElement('canvas');
    c.width = v.videoWidth; c.height = v.videoHeight;
    c.getContext('2d').drawImage(v, 0, 0);
    const base = baseName(it.file.name);
    return it.out === 'png'
      ? [{ name: base + '.png', blob: await canvasBlob(c, 'image/png') }]
      : [{ name: base + '.jpg', blob: await canvasBlob(c, 'image/jpeg', o.q) }];
  } finally { cleanupVideo(v); }
}

async function recordVideo(it, progress) {
  const rec = REC.video;
  const v = await loadVideo(it.file);
  try {
    const cap = v.captureStream ? v.captureStream() : v.mozCaptureStream ? v.mozCaptureStream() : null;
    if (!cap) throw new Error('Questo browser non supporta la conversione video: usa Chrome o Firefox.');
    // L'audio passa da un AudioContext così non esce dalle casse durante la registrazione
    const tracks = [...cap.getVideoTracks()];
    let ctx = null;
    try {
      ctx = new AudioContext();
      const dest = ctx.createMediaStreamDestination();
      ctx.createMediaElementSource(v).connect(dest);
      tracks.push(...dest.stream.getAudioTracks());
    } catch { tracks.push(...cap.getAudioTracks()); }
    const mr = new MediaRecorder(new MediaStream(tracks), { mimeType: rec.mime });
    const chunks = [];
    mr.ondataavailable = e => e.data.size && chunks.push(e.data);
    const stopped = new Promise(r => (mr.onstop = r));
    v.ontimeupdate = () => v.duration && progress(v.currentTime / v.duration);
    mr.start(1000);
    await v.play();
    await new Promise(r => (v.onended = r));
    mr.stop();
    await stopped;
    ctx?.close();
    return [{ name: `${baseName(it.file.name)}.${rec.ext}`, blob: new Blob(chunks, { type: rec.mime.split(';')[0] }) }];
  } finally { cleanupVideo(v); }
}

async function convMedia(it, o, progress) {
  const base = baseName(it.file.name);
  if (it.cat === 'video' && (it.out === 'png' || it.out === 'jpg')) return videoFrame(it, o);
  if (it.cat === 'video' && it.out === 'rec') return recordVideo(it, progress);
  const buf = await decodeAudio(it.file);
  if (it.out === 'wav') return [{ name: base + '.wav', blob: encodeWAV(buf) }];
  if (it.out === 'mp3') return [{ name: base + '.mp3', blob: await encodeMP3(buf, o.kbps, progress) }];
  if (it.out === 'rec') return [{ name: `${base}.${REC.audio.ext}`, blob: await recordAudioBuffer(buf, REC.audio, progress) }];
}

/* =========================================================
   Scheda «Converti»
   ========================================================= */
const items = [];
let uid = 0;
let mergedResult = null;

function readOpts() {
  const maxW = parseInt($('#maxW').value, 10);
  return {
    q: +$('#q').value / 100,
    maxW: maxW > 0 ? maxW : 0,
    kbps: +$('#kbps').value,
    rw: $('#convRw').checked,
    mergePdf: $('#mergePdf').checked,
  };
}

function convert(it, o, progress) {
  switch (it.cat) {
    case 'image': return convImage(it, o);
    case 'pdf':   return convPdf(it, o, progress);
    case 'doc':   return convDoc(it);
    case 'data':  return convData(it);
    case 'audio':
    case 'video': return convMedia(it, o, progress);
  }
  throw new Error('Formato non supportato.');
}

function addFiles(files) {
  for (const file of files) {
    const { cat, ext } = detect(file);
    const outs = outputsFor(cat, ext);
    const it = { id: ++uid, file, cat, ext, outs, out: outs[0], results: [], busy: false };
    it.el = renderItem(it);
    items.push(it);
    $('#list').append(it.el);
  }
  refreshBar();
}

function renderItem(it) {
  const li = document.createElement('li');
  li.className = 'item';
  li.innerHTML = `
    <div class="thumb"></div>
    <div class="meta">
      <div class="name"></div>
      <div class="sub"></div>
      <div class="status"></div>
      <div class="results"></div>
    </div>
    <div class="act">
      <select aria-label="Formato di destinazione"></select>
      <button class="primary go">Converti</button>
      <button class="x" title="Rimuovi" aria-label="Rimuovi">×</button>
    </div>`;
  $('.name', li).textContent = it.file.name;
  $('.sub', li).textContent = `${it.cat ? CATS[it.cat].label : 'Sconosciuto'} · ${it.ext.toUpperCase() || '?'} · ${fmtSize(it.file.size)}`;

  const thumb = $('.thumb', li);
  if (it.cat === 'image') {
    const img = new Image();
    img.src = URL.createObjectURL(it.file);
    img.onload = () => URL.revokeObjectURL(img.src);
    img.onerror = () => { URL.revokeObjectURL(img.src); thumb.textContent = it.ext; };
    thumb.append(img);
  } else thumb.textContent = it.ext || '?';

  const sel = $('select', li);
  if (!it.outs.length) {
    sel.hidden = true;
    $('.go', li).hidden = true;
    setStatus(it, 'Formato non supportato', 'err', li);
  } else {
    sel.innerHTML = it.outs.map(o => `<option value="${o}">${esc(fmtLabel(o, it.cat))}</option>`).join('');
    sel.onchange = () => (it.out = sel.value);
  }
  $('.go', li).onclick = () => runItem(it);
  $('.x', li).onclick = () => {
    items.splice(items.indexOf(it), 1);
    it.results.forEach(r => r.url && URL.revokeObjectURL(r.url));
    li.remove();
    refreshBar();
  };
  return li;
}

function setStatus(it, text, cls = '', li = it.el) {
  const s = $('.status', li);
  s.textContent = text;
  s.className = 'status ' + cls;
}

function showResults(box, results) {
  box.innerHTML = '';
  for (const r of results) {
    r.url = URL.createObjectURL(r.blob);
    const a = document.createElement('a');
    a.href = r.url;
    a.download = r.name;
    a.textContent = `⬇ ${r.name} (${fmtSize(r.blob.size)})`;
    box.append(a);
  }
}

async function runItem(it, o = readOpts()) {
  if (it.busy || !it.outs.length) return;
  it.busy = true;
  $('.go', it.el).disabled = true;
  it.results.forEach(r => r.url && URL.revokeObjectURL(r.url));
  it.results = [];
  $('.results', it.el).innerHTML = '';
  const live = it.out === 'rec' ? ' (dura quanto il file)' : '';
  setStatus(it, `Conversione in corso…${live}`, 'busy');
  const t0 = performance.now();
  try {
    const res = await convert(it, o, p => setStatus(it, `Conversione in corso… ${Math.min(100, Math.round(p * 100))}%${live}`, 'busy'));
    it.results = res;
    showResults($('.results', it.el), res);
    setStatus(it, `Fatto in ${((performance.now() - t0) / 1000).toFixed(1)} s`, 'ok');
  } catch (e) {
    console.error(e);
    setStatus(it, e.message || String(e), 'err');
  } finally {
    it.busy = false;
    $('.go', it.el).disabled = false;
    refreshBar();
  }
}

async function convertAll() {
  const o = readOpts();
  $('#convAll').disabled = true;
  const imgsToPdf = items.filter(it => it.cat === 'image' && it.out === 'pdf');
  const rest = o.mergePdf && imgsToPdf.length > 1 ? items.filter(it => !imgsToPdf.includes(it)) : items;
  try {
    if (rest !== items) await mergeImagesToPdf(imgsToPdf, o);
    for (const it of rest) await runItem(it, o);
  } finally {
    $('#convAll').disabled = false;
    refreshBar();
  }
}

async function mergeImagesToPdf(list, o) {
  const box = $('#extra');
  box.innerHTML = `<div class="extra-card"><div class="thumb">PDF</div><div class="meta"><div class="name">Immagini unite.pdf</div><div class="sub">${list.length} immagini in un unico PDF</div><div class="status busy">Creazione del PDF…</div><div class="results"></div></div></div>`;
  try {
    const canvases = [];
    for (const it of list) canvases.push(await imageCanvas(it, o));
    const blob = await canvasesToPdf(canvases, o.q, o.rw);
    mergedResult = { name: 'Immagini unite.pdf', blob };
    showResults($('.results', box), [mergedResult]);
    $('.status', box).textContent = 'Fatto';
    $('.status', box).className = 'status ok';
  } catch (e) {
    console.error(e);
    $('.status', box).textContent = e.message;
    $('.status', box).className = 'status err';
  }
}

function refreshBar() {
  const has = items.length > 0;
  $('#convBar').hidden = !has;
  $('#opts').hidden = !has;
  $('#drop').classList.toggle('compact', has);
  if (!has) { $('#extra').innerHTML = ''; mergedResult = null; }

  const all = [...new Set(items.flatMap(it => it.outs))];
  const sel = $('#bulkFmt');
  const prev = sel.value;
  sel.innerHTML = '<option value="">— scegli —</option>' + all.map(o => `<option value="${o}">${esc(o === 'rec' ? 'Registrazione (WEBM/OGG…)' : FMT_LABEL[o] || o.toUpperCase())}</option>`).join('');
  if (all.includes(prev)) sel.value = prev;

  const anyResult = items.some(it => it.results.length) || mergedResult;
  $('#zipAll').disabled = !anyResult;
}

async function zipAll() {
  const btn = $('#zipAll');
  btn.disabled = true;
  btn.textContent = 'Preparo lo ZIP…';
  try {
    const files = items.flatMap(it => it.results);
    if (mergedResult) files.unshift(mergedResult);
    await zipAndDownload(files, 'Convertiasc.zip');
  } catch (e) { alertInline(e.message); }
  finally { btn.textContent = 'Scarica tutto (ZIP)'; refreshBar(); }
}

async function zipAndDownload(files, zipName) {
  await lib('jszip');
  const zip = new JSZip();
  const used = new Set();
  for (const f of files) {
    let name = f.name, n = 1;
    while (used.has(name)) name = f.name.replace(/(\.[^.]+)?$/, ` (${++n})$1`);
    used.add(name);
    zip.file(name, f.blob);
  }
  download(await zip.generateAsync({ type: 'blob' }), zipName);
}

function alertInline(msg) {
  const box = $('#extra');
  box.innerHTML = `<div class="extra-card"><div class="thumb">!</div><div class="meta"><div class="status err"></div></div></div>`;
  $('.status', box).textContent = msg;
}

/* =========================================================
   Scheda «Rimuovi bianco»
   ========================================================= */
const RW_PREVIEW_MAX = 1400;
const rw = { list: [], cur: -1, showOriginal: false, raf: 0 };

async function rwAdd(files) {
  for (const file of files) {
    const { cat, ext } = detect(file);
    if (cat !== 'image') continue;
    try {
      const full = await fileToCanvas(file, ext, 0);
      const scale = Math.min(1, RW_PREVIEW_MAX / Math.max(full.width, full.height));
      const pc = document.createElement('canvas');
      pc.width = Math.max(1, Math.round(full.width * scale));
      pc.height = Math.max(1, Math.round(full.height * scale));
      const g = pc.getContext('2d', { willReadFrequently: true });
      g.drawImage(full, 0, 0, pc.width, pc.height);
      const entry = { file, full, preview: g.getImageData(0, 0, pc.width, pc.height), thumb: URL.createObjectURL(file) };
      rw.list.push(entry);
    } catch (e) {
      $('#rwInfo').textContent = `${file.name}: ${e.message}`;
    }
  }
  if (rw.cur < 0 && rw.list.length) rw.cur = 0;
  rwRenderThumbs();
  rwRender();
}

function rwRenderThumbs() {
  const box = $('#rwThumbs');
  box.innerHTML = '';
  rw.list.forEach((e, i) => {
    const b = document.createElement('button');
    b.className = i === rw.cur ? 'on' : '';
    b.title = e.file.name;
    b.innerHTML = `<img alt="">`;
    $('img', b).src = e.thumb;
    b.onclick = () => { rw.cur = i; rwRenderThumbs(); rwRender(); };
    box.append(b);
  });
  const has = rw.list.length > 0;
  $('#stage').hidden = !has;
  $('#rwDrop').classList.toggle('compact', has);
  $('#rwDl').disabled = !has;
  $('#rwDlAll').disabled = rw.list.length < 2;
}

function rwRender() {
  cancelAnimationFrame(rw.raf);
  rw.raf = requestAnimationFrame(() => {
    const e = rw.list[rw.cur];
    if (!e) return;
    const c = $('#rwCanvas');
    c.width = e.preview.width;
    c.height = e.preview.height;
    const data = rw.showOriginal ? e.preview : removeColor(e.preview, rwSettings());
    c.getContext('2d').putImageData(data, 0, 0);
    $('#rwInfo').textContent = `${e.file.name} · ${e.full.width}×${e.full.height} px`;
  });
}

function rwFullResult(e) {
  const c = document.createElement('canvas');
  c.width = e.full.width;
  c.height = e.full.height;
  const g = c.getContext('2d');
  g.drawImage(e.full, 0, 0);
  g.putImageData(removeColor(g.getImageData(0, 0, c.width, c.height), rwSettings()), 0, 0);
  return c;
}

const rwName = e => `${baseName(e.file.name)}-senza-sfondo.png`;

async function rwDownloadCurrent() {
  const e = rw.list[rw.cur];
  if (!e) return;
  download(await canvasBlob(rwFullResult(e), 'image/png'), rwName(e));
}

async function rwDownloadAll() {
  const btn = $('#rwDlAll');
  btn.disabled = true;
  btn.textContent = 'Preparo lo ZIP…';
  try {
    const files = [];
    for (const e of rw.list) { files.push({ name: rwName(e), blob: await canvasBlob(rwFullResult(e), 'image/png') }); await sleep(); }
    await zipAndDownload(files, 'Convertiasc-senza-sfondo.zip');
  } catch (err) { $('#rwInfo').textContent = err.message; }
  finally { btn.textContent = 'Scarica tutte (ZIP)'; btn.disabled = rw.list.length < 2; }
}

function rwPickColor(ev) {
  const e = rw.list[rw.cur];
  if (!e) return;
  const c = ev.currentTarget;
  const r = c.getBoundingClientRect();
  const x = Math.floor((ev.clientX - r.left) * c.width / r.width);
  const y = Math.floor((ev.clientY - r.top) * c.height / r.height);
  const i = (y * e.preview.width + x) * 4;
  const d = e.preview.data;
  $('#rwColor').value = '#' + [d[i], d[i + 1], d[i + 2]].map(v => v.toString(16).padStart(2, '0')).join('');
  rwRender();
}

/* =========================================================
   Scheda «Formati»
   ========================================================= */
function renderFormats() {
  const rows = Object.entries(CATS).map(([cat, def]) => {
    const outs = def.out.filter(o => o !== 'rec' || REC[cat]).map(o => fmtLabel(o, cat));
    return `<tr><th>${def.label}</th><td>${def.ext.map(e => `<code>${e}</code>`).join('')}</td><td>${outs.map(o => `<code>${esc(o)}</code>`).join('')}</td></tr>`;
  });
  $('#fmtTable').innerHTML = `<thead><tr><th>Tipo</th><th>Da</th><th>A</th></tr></thead><tbody>${rows.join('')}
    <tr><th>Rimuovi bianco</th><td><code>tutte le immagini</code></td><td><code>PNG trasparente</code></td></tr></tbody>`;
}

/* =========================================================
   Collegamenti all'interfaccia
   ========================================================= */
function setupDrop(zone, input, onFiles) {
  input.onchange = () => { onFiles([...input.files]); input.value = ''; };
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('over');
    onFiles([...e.dataTransfer.files]);
  });
}

function init() {
  // schede
  $$('.tabs button').forEach(b => b.onclick = () => {
    $$('.tabs button').forEach(x => x.classList.toggle('on', x === b));
    $$('.tab').forEach(t => (t.hidden = t.id !== 'tab-' + b.dataset.tab));
  });

  // trascinare un file fuori dalle zone non deve aprirlo nel browser
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => {
    e.preventDefault();
    const files = [...e.dataTransfer.files];
    if (!files.length) return;
    if (!$('#tab-rw').hidden) rwAdd(files);
    else if (!$('#tab-conv').hidden) addFiles(files);
  });

  // converti
  setupDrop($('#drop'), $('#fileIn'), addFiles);
  $('#convAll').onclick = convertAll;
  $('#zipAll').onclick = zipAll;
  $('#clearAll').onclick = () => {
    items.splice(0).forEach(it => { it.results.forEach(r => r.url && URL.revokeObjectURL(r.url)); it.el.remove(); });
    refreshBar();
  };
  $('#bulkFmt').onchange = e => {
    const f = e.target.value;
    if (!f) return;
    for (const it of items) if (it.outs.includes(f)) { it.out = f; $('select', it.el).value = f; }
  };
  $('#q').oninput = e => ($('#qV').textContent = e.target.value);

  // rimuovi bianco
  setupDrop($('#rwDrop'), $('#rwIn'), rwAdd);
  setupDrop($('#stage'), $('#rwIn'), rwAdd);
  $('#tol').oninput = e => { $('#tolV').textContent = e.target.value; rwRender(); };
  $('#soft').oninput = e => { $('#softV').textContent = e.target.value; rwRender(); };
  $$('input[name=rwMode]').forEach(r => (r.onchange = rwRender));
  $('#defringe').onchange = rwRender;
  $('#rwColor').oninput = rwRender;
  $('#rwReset').onclick = () => { $('#rwColor').value = '#ffffff'; rwRender(); };
  $('#rwCanvas').onclick = rwPickColor;
  $$('#bgSeg button').forEach(b => b.onclick = () => {
    $$('#bgSeg button').forEach(x => x.classList.toggle('on', x === b));
    const st = $('#stage');
    st.classList.remove('checker', 'dark', 'light');
    st.classList.add(b.dataset.bg);
  });
  const hold = $('#rwHold');
  const setOrig = v => { rw.showOriginal = v; rwRender(); };
  hold.addEventListener('pointerdown', () => setOrig(true));
  ['pointerup', 'pointerleave', 'pointercancel'].forEach(ev => hold.addEventListener(ev, () => rw.showOriginal && setOrig(false)));
  $('#rwDl').onclick = rwDownloadCurrent;
  $('#rwDlAll').onclick = rwDownloadAll;

  renderFormats();
}

init();
