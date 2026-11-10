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
const ALL_TOGGLES = [...TRANSFORM_IDS, 'homoglyphNFKC', 'docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'];

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
  meta: ['docxMetadata', 'docxRsid', 'pdfMeta', 'imageMeta'],
};

// ---- File processing -------------------------------------------------------

async function processEntry(entry) {
  const buf = await FP.readAsArrayBuffer(entry.file);
  const opts = readOptions();
  const ext = FP.extOf(entry.name);
  if (entry.kind === 'docx') {
    const { blob } = await FP.cleanDocx(buf, metaOptions());
    const text = await FP.docxToText(buf);
    return { blob, outName: cleanName(entry.name), editorText: cleanTextOf(text, opts) };
  }
  if (entry.kind === 'pdf') {
    const { blob } = metaOptions().pdfMeta ? await FP.cleanPdf(buf) : { blob: new Blob([buf], { type: 'application/pdf' }) };
    const text = await FP.pdfToText(buf);
    return { blob, outName: cleanName(entry.name), editorText: cleanTextOf(text, opts) };
  }
  if (entry.kind === 'image') {
    if (!metaOptions().imageMeta) return { blob: new Blob([buf]), outName: entry.name };
    const { blob } = await FP.cleanImage(buf, entry.file.type, ext === 'jpeg' ? 'jpg' : ext);
    return { blob, outName: cleanName(entry.name) };
  }
  // text
  const raw = await FP.readAsText(entry.file);
  const cleaned = cleanTextOf(raw, opts);
  return { blob: new Blob([cleaned], { type: 'text/plain' }), outName: cleanName(entry.name), editorText: cleaned, editorRaw: raw };
}

function cleanTextOf(text, opts) {
  return SAN.sanitizeText(text, opts).output;
}
function cleanName(name) {
  return name.replace(/(\.[^.]+)$/, '.clean$1');
}

async function handleFiles(files) {
  const arr = Array.from(files);
  for (const file of arr) {
    const entry = { file, kind: FP.kindOf(file.name), name: file.name, status: 'processing', blob: null, outName: null };
    fileEntries.push(entry);
    renderFileList();
    try {
      const res = await processEntry(entry);
      entry.blob = res.blob; entry.outName = res.outName; entry.status = 'done';
      // If this is the only file and it carries text, load it into the editor.
      if (arr.length === 1 && res.editorRaw != null) $('input').value = res.editorRaw;
      else if (arr.length === 1 && res.editorText != null) $('input').value = sourceTextFor(entry);
      if (arr.length === 1 && (entry.kind === 'docx' || entry.kind === 'pdf')) {
        $('input').value = await rawTextFor(entry);
      }
      processLive();
    } catch (e) {
      entry.status = 'error'; entry.error = e.message || String(e);
    }
    renderFileList();
  }
}

// Re-extract raw (uncleaned) text so the editor shows the before-state.
async function rawTextFor(entry) {
  const buf = await FP.readAsArrayBuffer(entry.file);
  if (entry.kind === 'docx') return FP.docxToText(buf);
  if (entry.kind === 'pdf') return FP.pdfToText(buf);
  return '';
}
function sourceTextFor() { return $('input').value; }

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

let reprocessTimer = null;
function scheduleReprocess() {
  processLive();
  clearTimeout(reprocessTimer);
  reprocessTimer = setTimeout(async () => {
    for (const e of fileEntries) {
      if (e.kind === 'text') continue; // text handled live; binaries re-clean
      try { const res = await processEntry(e); e.blob = res.blob; e.outName = res.outName; e.status = 'done'; }
      catch (err) { e.status = 'error'; e.error = err.message; }
    }
    renderFileList();
  }, 400);
}

// ---- Init ------------------------------------------------------------------

function init() {
  $('input').addEventListener('input', processLive);
  for (const id of [...ALL_TOGGLES, 'dashReplacement']) $(id).addEventListener('change', scheduleReprocess);
  for (const id of Object.keys(PROTECT_MAP)) $(id).addEventListener('change', processLive);
  $('keywordInput').addEventListener('input', processLive);
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

  processLive();
}

document.addEventListener('DOMContentLoaded', init);
