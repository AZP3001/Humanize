// app.js — UI wiring: live text cleanup, toggles, presets, file processing,
// X-ray diff, scoreboard, and export.
'use strict';

const $ = (id) => document.getElementById(id);
const FP = window.FileProcessor;
const SAN = window.Sanitizer;

const TRANSFORM_IDS = ['zeroWidth', 'bidiControl', 'variationTags', 'homoglyph',
  'whitespace', 'smartQuotes', 'dashes', 'ellipsisLigature'];
const PROTECT_MAP = { protectUrls: 'urls', protectEmails: 'emails', protectCode: 'code',
  protectCitations: 'citations', protectMath: 'math', protectKeywords: 'keywords' };

let fileEntries = []; // {file, kind, name, blob, outName, status}

// ---- Options from UI -------------------------------------------------------

function readOptions() {
  const transforms = {};
  for (const id of TRANSFORM_IDS) transforms[id] = $(id).checked;
  transforms.stripFormatting = $('stripFormatting').checked;
  const protect = {};
  for (const [id, key] of Object.entries(PROTECT_MAP)) protect[key] = $(id).checked;
  protect.keywordList = $('keywordInput').value.split(',').map((s) => s.trim()).filter(Boolean);
  return {
    transforms,
    homoglyphNFKC: $('homoglyphNFKC').checked,
    dashReplacement: $('dashReplacement').value,
    protect,
  };
}

function metaOptions() {
  return {
    docxMetadata: $('docxMetadata').checked,
    docxRsid: $('docxRsid').checked,
    pdfMeta: $('pdfMeta').checked,
    imageMeta: $('imageMeta').checked,
  };
}

// ---- Live text processing --------------------------------------------------

let lastSegments = [];

function processLive() {
  const text = $('input').value;
  const opts = readOptions();
  const { output, segments, counts } = SAN.sanitizeText(text, opts);
  lastSegments = segments;
  $('output').value = output;

  $('scHidden').textContent = counts.hidden;
  $('scConf').textContent = counts.confusable;
  $('scTypo').textContent = counts.typography;
  const residual = SAN.countHidden(output);
  $('scRisk').textContent = residual === 0 ? 'clear' : residual + ' left';
  $('scRisk').style.color = residual === 0 ? 'var(--accent)' : 'var(--hidden)';

  $('inStats').textContent = `${text.length} chars · ${SAN.countHidden(text)} hidden`;
  $('outStats').textContent = `${output.length} chars`;

  if ($('xrayToggle').checked) renderXray();
}

function renderXray() {
  const box = $('xray');
  box.innerHTML = '';
  for (const seg of lastSegments) {
    if (seg.type === null) {
      box.appendChild(document.createTextNode(seg.text));
    } else if (seg.type === 'hidden') {
      const span = document.createElement('span');
      span.className = 'x-hidden';
      const cp = seg.orig.codePointAt(0).toString(16).toUpperCase().padStart(4, '0');
      span.textContent = `[U+${cp}]`;
      box.appendChild(span);
    } else {
      const span = document.createElement('span');
      span.className = seg.type === 'confusable' ? 'x-conf' : 'x-typo';
      span.textContent = seg.text || seg.orig;
      span.title = `${seg.orig} → ${seg.text || '∅'}`;
      box.appendChild(span);
    }
  }
}

function toggleXray() {
  const on = $('xrayToggle').checked;
  $('xray').classList.toggle('hidden', !on);
  $('output').classList.toggle('hidden', on);
  $('xrayLegend').classList.toggle('hidden', !on);
  if (on) renderXray();
}

// ---- Presets ---------------------------------------------------------------

const PRESETS = {
  safeClean: { on: [...TRANSFORM_IDS, 'docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'],
    off: ['homoglyphNFKC'], dash: ' - ' },
  watermarkOnly: { on: ['zeroWidth', 'bidiControl', 'variationTags', 'homoglyph',
    'docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'],
    off: ['whitespace', 'smartQuotes', 'dashes', 'ellipsisLigature', 'homoglyphNFKC'], dash: ' - ' },
  fullNormalize: { on: [...TRANSFORM_IDS, 'homoglyphNFKC',
    'docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'], off: [], dash: ' - ' },
  minimal: { on: ['zeroWidth', 'bidiControl', 'variationTags', 'docxMetadata', 'pdfMeta', 'imageMeta'],
    off: ['homoglyph', 'homoglyphNFKC', 'whitespace', 'smartQuotes', 'dashes', 'ellipsisLigature', 'docxRsid'],
    dash: ' - ' },
};
const ALL_TOGGLES = [...TRANSFORM_IDS, 'homoglyphNFKC', 'stripFormatting', 'docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'];

function applyPreset(name) {
  const p = PRESETS[name];
  if (!p) return;
  for (const id of ALL_TOGGLES) $(id).checked = p.on.includes(id);
  $('dashReplacement').value = p.dash;
  processLive();
}

// ---- Master switches -------------------------------------------------------

const CATEGORY = {
  protect: ['protectUrls', 'protectEmails', 'protectCode', 'protectCitations', 'protectMath', 'protectKeywords'],
  strip: ['zeroWidth', 'bidiControl', 'variationTags', 'homoglyph'],
  typo: ['whitespace', 'smartQuotes', 'dashes', 'ellipsisLigature'],
  format: ['stripFormatting'],
  meta: ['docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'],
};

// ---- File processing -------------------------------------------------------

async function processEntry(entry) {
  const buf = await FP.readAsArrayBuffer(entry.file);
  const opts = readOptions();
  const m = metaOptions();
  const ext = FP.extOf(entry.name);
  const outName = cleanName(entry.name);
  // "Strip all formatting" on a document means: give back its raw text only.
  if (opts.transforms.stripFormatting && (entry.kind === 'docx' || entry.kind === 'pdf')) {
    const raw = entry.raw != null ? entry.raw : await (entry.kind === 'docx' ? FP.docxToText(buf) : FP.pdfToText(buf));
    return { blob: new Blob([cleanTextOf(raw, opts)], { type: 'text/plain' }),
      outName: entry.name.replace(/\.[^.]+$/, '.clean.txt'), raw };
  }
  if (entry.kind === 'docx') {
    // Formatting-strip is for plain text; inside a .docx only character cleanup applies.
    const inner = { ...opts, transforms: { ...opts.transforms, stripFormatting: false } };
    const { blob } = await FP.cleanDocx(buf, { metadata: m.docxMetadata, rsid: m.docxRsid,
      textFn: (t) => SAN.sanitizeText(t, inner).output });
    return { blob, outName, raw: entry.raw != null ? entry.raw : await FP.docxToText(buf) };
  }
  if (entry.kind === 'pdf') {
    const blob = m.pdfMeta ? (await FP.cleanPdf(buf)).blob : new Blob([buf], { type: 'application/pdf' });
    return { blob, outName, raw: entry.raw != null ? entry.raw : await FP.pdfToText(buf) };
  }
  if (entry.kind === 'image') {
    if (!m.imageMeta) return { blob: new Blob([buf], { type: entry.file.type }), outName: entry.name };
    const { blob } = await FP.cleanImage(buf, entry.file.type, ext === 'jpeg' ? 'jpg' : ext);
    return { blob, outName };
  }
  // ignoreBOM keeps a leading BOM visible so it is counted and stripped like any hidden char.
  const raw = new TextDecoder('utf-8', { ignoreBOM: true }).decode(buf);
  return { blob: new Blob([cleanTextOf(raw, opts)], { type: 'text/plain' }), outName, raw };
}

function cleanTextOf(text, opts) {
  return SAN.sanitizeText(text, opts).output;
}
function cleanName(name) {
  return name.replace(/(\.[^.]+)$/, '.clean$1');
}

async function runEntry(entry) {
  try { Object.assign(entry, await processEntry(entry), { status: 'done' }); }
  catch (e) { entry.status = 'error'; entry.error = e.message || String(e); }
}

async function handleFiles(files) {
  const arr = Array.from(files);
  for (const file of arr) {
    const entry = { file, kind: FP.kindOf(file.name), name: file.name, status: 'processing', blob: null, outName: null };
    fileEntries.push(entry);
    renderFileList();
    await runEntry(entry);
    // A single dropped document also loads into the editor for preview.
    if (arr.length === 1 && entry.raw != null) { $('input').value = entry.raw; processLive(); }
    renderFileList();
  }
}

function renderFileList() {
  const ul = $('fileList');
  ul.innerHTML = '';
  fileEntries.forEach((e, idx) => {
    const li = document.createElement('li');
    const name = document.createElement('span');
    name.className = 'fname'; name.textContent = e.name;
    const stat = document.createElement('span');
    stat.className = 'fstat ' + (e.status === 'done' ? 'ok' : e.status === 'error' ? 'err' : '');
    stat.textContent = e.status === 'done' ? '✓ cleaned' : e.status === 'error' ? ('✗ ' + e.error) : '…';
    li.append(name, stat);
    if (e.status === 'done') {
      const dl = document.createElement('button');
      dl.textContent = 'Download';
      dl.onclick = () => saveBlob(e.blob, e.outName);
      li.appendChild(dl);
    }
    const rm = document.createElement('button');
    rm.textContent = '✕';
    rm.onclick = () => { fileEntries.splice(idx, 1); renderFileList(); };
    li.appendChild(rm);
    ul.appendChild(li);
  });
}

// ---- Export ----------------------------------------------------------------

function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function flash(msg) {
  $('exportMsg').textContent = msg;
  setTimeout(() => { $('exportMsg').textContent = ''; }, 2500);
}

async function copyOutput() {
  try { await navigator.clipboard.writeText($('output').value); flash('Copied'); }
  catch (e) { flash('Copy failed'); }
}

function downloadOutput() {
  const done = fileEntries.filter((e) => e.status === 'done');
  if (done.length === 1) { saveBlob(done[0].blob, done[0].outName); return; }
  saveBlob(new Blob([$('output').value], { type: 'text/plain' }), 'output.clean.txt');
}

async function batchZip() {
  const done = fileEntries.filter((e) => e.status === 'done');
  if (!done.length) {
    saveBlob(new Blob([$('output').value], { type: 'text/plain' }), 'output.clean.txt');
    return;
  }
  const zip = new JSZip();
  for (const e of done) zip.file(e.outName, e.blob);
  const blob = await zip.generateAsync({ type: 'blob' });
  saveBlob(blob, 'docscrub-batch.zip');
  flash(`Zipped ${done.length} files`);
}

// ---- Re-process files when toggles change ----------------------------------

const PERSIST = () => [...ALL_TOGGLES, ...Object.keys(PROTECT_MAP)];
function saveSettings() {
  try {
    const st = { dashReplacement: $('dashReplacement').value, keywordInput: $('keywordInput').value };
    for (const id of PERSIST()) st[id] = $(id).checked;
    localStorage.setItem('docscrub', JSON.stringify(st));
  } catch (e) { /* storage unavailable: settings just don't persist */ }
}
function loadSettings() {
  try {
    const st = JSON.parse(localStorage.getItem('docscrub') || 'null');
    if (!st) return;
    for (const id of PERSIST()) if (typeof st[id] === 'boolean') $(id).checked = st[id];
    if (typeof st.dashReplacement === 'string') $('dashReplacement').value = st.dashReplacement;
    if (typeof st.keywordInput === 'string') $('keywordInput').value = st.keywordInput;
  } catch (e) { /* ignore corrupt or blocked storage */ }
}

let reprocessTimer = null;
function scheduleReprocess() {
  saveSettings();
  processLive();
  clearTimeout(reprocessTimer);
  reprocessTimer = setTimeout(async () => {
    for (const e of fileEntries.slice()) await runEntry(e);
    renderFileList();
  }, 400);
}

// ---- Per-setting help: hover tooltip + interactive details modal -----------

const T_ALL = { zeroWidth: true, bidiControl: true, variationTags: true, homoglyph: true,
  whitespace: true, smartQuotes: true, dashes: true, ellipsisLigature: true };

const SETTINGS_INFO = {
  protectUrls: { title: 'Term Lock — URLs', kind: 'protect',
    short: 'Leaves URLs untouched by every other rule.',
    long: 'Web addresses often contain characters (dashes, case, tildes) that other rules would change and break the link. With this on, anything matching an http(s) URL is locked and passes through unchanged.',
    ex: 'Dash fix off inside links: visit https://my-site.com/a—b now.' },
  protectEmails: { title: 'Term Lock — Emails', kind: 'protect',
    short: 'Leaves email addresses unchanged.',
    long: 'Email addresses are locked so normalization never alters the local part or domain.',
    ex: 'Write to First.Last@Example.com — it stays exactly as typed.' },
  protectCode: { title: 'Term Lock — Inline code', kind: 'protect',
    short: 'Protects text inside `backticks`.',
    long: 'Anything between backticks is treated as code and locked, so quotes, dashes and spacing inside it are preserved verbatim.',
    ex: 'Run `git commit —amend` and keep “quotes” inside code intact.' },
  protectCitations: { title: 'Term Lock — Citations', kind: 'protect',
    short: 'Protects [1] and (Author, Year) citations.',
    long: 'Bracketed numeric refs like [1] and author-year refs like (Smith, 2025) are locked so punctuation and spacing inside them are never changed.',
    ex: 'As shown [1] and by others (Smith, 2025), results held.' },
  protectMath: { title: 'Term Lock — Math', kind: 'protect',
    short: 'Protects $…$ and \\(…\\) formulas.',
    long: 'LaTeX-style math spans are locked, so operators and symbols inside formulas are left alone.',
    ex: 'The identity $a—b = c$ must not have its dash changed.' },
  protectKeywords: { title: 'Term Lock — Custom keywords', kind: 'protect',
    short: 'Protects words you list in the box below.',
    long: 'Type comma-separated terms in the field under this toggle (names, product terms, acronyms). Each exact occurrence is locked from every change.',
    ex: 'Enter your own keywords in the box — those exact words stay untouched.' },

  zeroWidth: { title: 'Zero-width & invisible characters', kind: 'text',
    short: 'Deletes invisible zero-width marks and BOM.',
    long: 'Characters like the zero-width space (U+200B), joiners, word joiner, BOM and soft hyphen are invisible but can carry hidden watermarks or tracking. This removes them entirely.',
    ex: 'He​llo⁠ wor‌ld﻿' },
  bidiControl: { title: 'Bidirectional & control codes', kind: 'text',
    short: 'Strips text-direction overrides and control codes.',
    long: 'Right-to-left/left-to-right override characters can secretly reorder text, and stray ASCII control codes can hide data. These are removed (normal tabs and newlines are kept).',
    ex: 'price: 100‮ usd ‬end' },
  variationTags: { title: 'Variation selectors & tag payloads', kind: 'text',
    short: 'Removes emoji variation selectors and hidden tag glyphs.',
    long: 'Variation selectors (U+FE00–FE0F) and the Unicode Tag block (U+E0000–E007F) are invisible and can smuggle an entire hidden message after a normal character. All are stripped.',
    ex: 'secret󠁴󠁡󠁧 message' },
  homoglyph: { title: 'Homoglyph / confusable normalizer', kind: 'text',
    short: 'Turns look-alike letters into plain ASCII.',
    long: 'Cyrillic and Greek letters that look identical to Latin ones (е, а, о, р, с) can be used to disguise text. This maps them back to ASCII when the word is mixed with Latin letters or made only of lookalikes; genuine Cyrillic/Greek words are left alone.',
    ex: 'PаyPаl аnd Gооgle lоgin (привет stays)' },
  homoglyphNFKC: { title: 'NFKC compatibility pass', kind: 'nfkc',
    short: 'Extra Unicode normalization, ASCII-safe only.',
    long: 'Applies Unicode NFKC normalization per character but keeps the result only when it becomes plain ASCII — catching fullwidth and other decorative variants without touching real non-Latin text.',
    ex: 'Ｈｅｌｌｏ Ｗｏｒｌｄ' },
  whitespace: { title: 'Exotic whitespace normalizer', kind: 'text',
    short: 'Converts fancy spaces to a normal space.',
    long: 'Non-breaking, narrow, figure, thin, hair and ideographic spaces all look like a space but are different characters. They are converted to the plain ASCII space (0x20).',
    ex: 'a b c　d' },
  smartQuotes: { title: 'Smart quote & apostrophe flattener', kind: 'text',
    short: 'Curly quotes and primes become straight.',
    long: 'Typographic “curly” quotes, guillemets and prime marks are converted to straight ASCII " and \'.',
    ex: '“Hello,” it’s a 6′ «test»' },
  dashes: { title: 'Em / en-dash normalizer', kind: 'text',
    short: 'Replaces — and – with your chosen form.',
    long: 'Em-dashes (—) and en-dashes (–) are replaced with the alternative you pick in the dropdown (spaced hyphen, comma, or plain hyphen).',
    ex: 'This — right here — and pages 10–12.' },
  dashReplacement: { title: 'Dash replacement choice', kind: 'text', runAs: 'dashes',
    short: 'Pick what — and – turn into.',
    long: 'Controls the replacement used by the em/en-dash rule: a spaced hyphen " - ", a comma ", ", or a plain hyphen "-". Number ranges like 10–12 always become 10-12. Change the dropdown, then re-open this to see it applied.',
    ex: 'This — right here — stays readable.' },
  ellipsisLigature: { title: 'Ellipsis & ligature splitter', kind: 'text',
    short: 'Expands … to ... and splits ﬁ ﬂ ligatures.',
    long: 'The single ellipsis character … becomes three dots, and typographic ligatures (ﬁ, ﬂ, ﬀ, ﬃ, ﬄ) are decomposed into normal ASCII letters.',
    ex: 'Wait… the ﬁnal ﬂourish was ﬀine.' },
  stripFormatting: { title: 'Strip all formatting', kind: 'text',
    short: 'Flattens Markdown & HTML to plain text.',
    long: 'Removes headings, bold/italic, links (keeping the visible text), lists, blockquotes, code fences, tables, HTML tags with their colours and styles, LaTeX styling and terminal colour codes, leaving only the readable words. A dropped .docx or .pdf then downloads as a plain .txt.',
    ex: '# Title\n\n**Bold** and *italic* with a [link](https://x.com).\n\n- first item\n- second item\n\n<b>raw html</b> too.' },

  docxMetadata: { title: '.docx metadata wipe', kind: 'file',
    short: 'Clears author/company/title props in a Word file.',
    long: 'On a .docx upload, this blanks the identifying fields in docProps/core.xml and app.xml (author, last-modified-by, title, company, etc.) and deletes custom.xml. It removes data only — it never writes a fake author or history in its place.',
    static: 'Before:  dc:creator = "Jane Doe", Company = "Acme"\nAfter:   dc:creator = "",         Company = ""\n\n(Upload a .docx to apply — this rule does not act on typed text.)' },
  docxRsid: { title: '.docx revision (rsid) strip', kind: 'file',
    short: 'Removes Word’s per-edit revision fingerprints.',
    long: 'Word stamps runs of text with rsid identifiers that can reveal editing sessions. This strips w:rsidR / w:rsidRPr / w:rsidP / w:rsidRDefault attributes and the rsid table. Removal only; nothing fabricated.',
    static: 'Before:  <w:p w:rsidR="00A1B2C3" w:rsidP="00FF00FF">\nAfter:   <w:p>\n\n(Upload a .docx to apply.)' },
  pdfMeta: { title: '.pdf metadata & XMP scrub', kind: 'file',
    short: 'Clears PDF document info and XMP metadata.',
    long: 'On a PDF upload, clears the Info dictionary (Title, Author, Producer, Creator, Keywords) and removes the XMP metadata stream via pdf-lib. The page content is kept as-is.',
    static: 'Before:  Author = "Jane Doe", Producer = "Word 2021"\nAfter:   Author = "",         Producer = ""\n\n(Upload a .pdf to apply.)' },
  imageMeta: { title: 'Image EXIF / XMP / C2PA strip', kind: 'file',
    short: 'Removes camera, GPS and provenance data from images.',
    long: 'Re-encodes PNG/JPG/WebP through a canvas, which drops EXIF (including GPS), XMP, IPTC, C2PA and PNG text chunks. Pixels are preserved (JPEG re-encodes at 95% quality).',
    static: 'Before:  GPS = 40.69°,-74.04°, Camera = iPhone 15\nAfter:   (no metadata — pixels only)\n\n(Upload an image to apply.)' },
};

function runExample(id, text) {
  const info = SETTINGS_INFO[id];
  const base = { homoglyphNFKC: false, dashReplacement: $('dashReplacement').value,
    protect: { urls: false, emails: false, code: false, citations: false, math: false } };
  if (info.kind === 'nfkc') return SAN.sanitizeText(text, { ...base, transforms: { homoglyph: true }, homoglyphNFKC: true }).output;
  if (info.kind === 'protect') {
    const key = { protectUrls: 'urls', protectEmails: 'emails', protectCode: 'code',
      protectCitations: 'citations', protectMath: 'math', protectKeywords: 'keywords' }[id];
    const protect = { ...base.protect, [key]: true };
    if (key === 'keywords') { protect.keywords = true; protect.keywordList = $('keywordInput').value.split(',').map((s) => s.trim()).filter(Boolean); }
    return SAN.sanitizeText(text, { ...base, transforms: { ...T_ALL }, protect }).output;
  }
  const tid = info.runAs || id;
  return SAN.sanitizeText(text, { ...base, transforms: { [tid]: true } }).output;
}

let tipTimer = null;
function showTip(row) {
  const id = row.dataset.info;
  const info = SETTINGS_INFO[id];
  if (!info) return;
  const tip = $('tooltip');
  tip.innerHTML = '';
  const txt = document.createElement('div');
  txt.textContent = info.short;
  tip.appendChild(txt);
  const more = document.createElement('span');
  more.className = 'tip-more';
  more.innerHTML = '<span class="tip-i">i</span> what it does';
  more.onclick = () => openModal(id);
  tip.appendChild(more);
  tip.classList.remove('hidden');
  const r = row.getBoundingClientRect();
  let top = r.top - tip.offsetHeight - 10;
  if (top < 8) top = r.bottom + 10;
  tip.style.left = Math.max(8, Math.min(r.left, window.innerWidth - tip.offsetWidth - 8)) + 'px';
  tip.style.top = top + 'px';
  tip.style.setProperty('--arrow', Math.min(60, r.width / 2) + 'px');
}
function hideTipSoon() { tipTimer = setTimeout(() => $('tooltip').classList.add('hidden'), 180); }
function cancelHide() { clearTimeout(tipTimer); }

function openModal(id) {
  const info = SETTINGS_INFO[id];
  if (!info) return;
  $('modalTitle').textContent = info.title;
  $('modalDesc').textContent = info.long;
  if (info.kind === 'file') {
    $('modalDemo').classList.add('hidden');
    const s = $('modalStatic');
    s.classList.remove('hidden');
    s.textContent = info.static;
  } else {
    $('modalStatic').classList.add('hidden');
    $('modalDemo').classList.remove('hidden');
    $('demoInput').value = info.ex;
    const run = () => { $('demoOutput').value = runExample(id, $('demoInput').value); };
    $('demoInput').oninput = run;
    run();
  }
  $('modal').classList.remove('hidden');
  $('tooltip').classList.add('hidden');
}
function closeModal() { $('modal').classList.add('hidden'); }

function initHelp() {
  document.querySelectorAll('.tgl').forEach((row) => {
    let id = row.dataset.info;
    if (!id) {
      const cb = row.querySelector('input[type=checkbox]');
      if (cb && SETTINGS_INFO[cb.id]) { id = cb.id; row.dataset.info = id; }
    }
    if (!id || !SETTINGS_INFO[id]) return;
    const b = document.createElement('button');
    b.className = 'info'; b.type = 'button'; b.textContent = 'i';
    b.title = 'What does this do?';
    b.onclick = (e) => { e.preventDefault(); e.stopPropagation(); openModal(id); };
    row.appendChild(b);
    row.addEventListener('pointerenter', () => { cancelHide(); showTip(row); }); // pointer* covers mouse, pen, touch
    row.addEventListener('pointerleave', hideTipSoon);
  });
  const tip = $('tooltip');
  tip.addEventListener('pointerenter', cancelHide);
  tip.addEventListener('pointerleave', hideTipSoon);
  $('modalClose').addEventListener('click', closeModal);
  $('modal').addEventListener('click', (e) => { if (e.target === $('modal')) closeModal(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeModal(); });
}

// ---- Init ------------------------------------------------------------------

function init() {
  $('input').addEventListener('input', processLive);
  for (const id of [...ALL_TOGGLES, 'dashReplacement']) $(id).addEventListener('change', scheduleReprocess);
  for (const id of Object.keys(PROTECT_MAP)) $(id).addEventListener('change', scheduleReprocess);
  $('keywordInput').addEventListener('input', scheduleReprocess);
  $('xrayToggle').addEventListener('change', toggleXray);

  $('presetSelect').addEventListener('change', (e) => applyPreset(e.target.value));

  document.querySelectorAll('[data-master]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const on = btn.dataset.on === '1';
      for (const id of CATEGORY[btn.dataset.master]) $(id).checked = on;
      scheduleReprocess();
    });
  });

  // File input / dropzone
  $('browseBtn').addEventListener('click', () => $('fileInput').click());
  $('fileInput').addEventListener('change', (e) => { handleFiles(e.target.files); e.target.value = ''; });
  const dz = $('dropzone');
  ['dragover', 'dragenter'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', (e) => { if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); });

  $('copyBtn').addEventListener('click', copyOutput);
  $('downloadBtn').addEventListener('click', downloadOutput);
  $('zipBtn').addEventListener('click', batchZip);

  loadSettings();
  initHelp();
  processLive();
}

document.addEventListener('DOMContentLoaded', init);
