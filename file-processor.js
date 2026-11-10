// file-processor.js — Read supported files, extract their text, and REMOVE
// forensic metadata from containers. Removal only: this file clears identifying
// data and leaves it blank/absent. It never fabricates replacement authoring
// history, revision counts, timing, or session IDs.

'use strict';

const TEXT_EXT = new Set(['txt', 'md', 'json', 'csv', 'html', 'htm', 'xml', 'tex']);
const IMAGE_EXT = new Set(['png', 'jpg', 'jpeg', 'webp']);

function extOf(name) {
  const m = /\.([A-Za-z0-9]+)$/.exec(name || '');
  return m ? m[1].toLowerCase() : '';
}
function kindOf(name) {
  const e = extOf(name);
  if (e === 'docx') return 'docx';
  if (e === 'pdf') return 'pdf';
  if (IMAGE_EXT.has(e)) return 'image';
  if (TEXT_EXT.has(e)) return 'text';
  return 'text'; // unknown -> treat as text
}

function readAsText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(r.error);
    r.readAsText(file);
  });
}
function readAsArrayBuffer(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = () => rej(r.error);
    r.readAsArrayBuffer(file);
  });
}

function xmlUnescape(s) {
  return s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

// ---- DOCX ------------------------------------------------------------------

async function docxToText(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docFile = zip.file('word/document.xml');
  if (!docFile) return '';
  const xml = await docFile.async('string');
  // Paragraphs -> newlines, tabs -> \t, strip remaining tags.
  const withBreaks = xml
    .replace(/<w:p[ >]/g, '\n<w:p ').replace(/<\/w:p>/g, '\n')
    .replace(/<w:tab\b[^>]*\/?>/g, '\t')
    .replace(/<w:br\b[^>]*\/?>/g, '\n');
  const text = withBreaks.replace(/<[^>]+>/g, '');
  return xmlUnescape(text).replace(/\n{3,}/g, '\n\n').trim();
}

// Strip rsid revision-tracking attributes from document/settings XML.
function stripRsidAttrs(xml) {
  return xml
    .replace(/\s+w:rsid(?:R|RPr|P|RDefault|Tr|Del|Sect)="[0-9A-Fa-f]*"/g, '')
    .replace(/<w:rsids>[\s\S]*?<\/w:rsids>/g, '')
    .replace(/<w:rsids\b[^>]*\/>/g, '');
}

function xmlEscape(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// Run textFn over the visible text of every <w:t> node (hidden chars live there).
function cleanDocXmlText(xml, textFn) {
  return xml.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g, (m, open, body, close) => {
    const out = textFn(xmlUnescape(body));
    if (/^\s|\s$/.test(out) && !open.includes('xml:space')) open = '<w:t xml:space="preserve">';
    return open + xmlEscape(out) + close;
  });
}

// Drop run styling (fonts, colour, bold, highlight) and paragraph styling
// (headings, alignment, list numbering). Section layout (w:sectPr) is kept.
// shortcut: regex, not an XML parser; tracked-change property blocks (w:pPrChange) are dropped with the rest.
function stripDocxFormatting(xml) {
  return xml
    .replace(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/g, '')
    .replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/g, (m, body) => {
      const sect = /<w:sectPr[\s\S]*<\/w:sectPr>|<w:sectPr[^>]*\/>/.exec(body);
      return sect ? `<w:pPr>${sect[0]}</w:pPr>` : '';
    });
}

// Clear identifying fields in docProps/core.xml (leave tags, blank the text).
function wipeCoreXml(xml) {
  const blank = [
    'dc:creator', 'cp:lastModifiedBy', 'dc:title', 'dc:subject', 'dc:description',
    'cp:keywords', 'cp:category', 'cp:contentStatus', 'cp:revision', 'cp:lastPrinted',
    'dcterms:created', 'dcterms:modified',
  ];
  let out = xml;
  for (const tag of blank) {
    const t = tag.replace(':', '\\:');
    out = out.replace(new RegExp(`(<${t}[^>]*>)[\\s\\S]*?(</${t}>)`, 'g'), '$1$2');
  }
  return out;
}

function wipeAppXml(xml) {
  let out = xml;
  for (const tag of ['Company', 'Manager', 'TotalTime', 'Template', 'HyperlinkBase']) {
    out = out.replace(new RegExp(`(<${tag}[^>]*>)[\\s\\S]*?(</${tag}>)`, 'g'), '$1$2');
  }
  return out;
}

// opts: { metadata, rsid, stripFormatting: bool, textFn?: (string) => string }
async function cleanDocx(arrayBuffer, opts) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const changes = [];

  if (opts.metadata) {
    if (zip.file('docProps/core.xml')) {
      const x = await zip.file('docProps/core.xml').async('string');
      zip.file('docProps/core.xml', wipeCoreXml(x));
      changes.push('core.xml');
    }
    if (zip.file('docProps/app.xml')) {
      const x = await zip.file('docProps/app.xml').async('string');
      zip.file('docProps/app.xml', wipeAppXml(x));
      changes.push('app.xml');
    }
    if (zip.file('docProps/custom.xml')) {
      zip.remove('docProps/custom.xml');
      // Remove dangling references so Word doesn't flag a repair.
      const ct = zip.file('[Content_Types].xml');
      if (ct) {
        const x = await ct.async('string');
        zip.file('[Content_Types].xml',
          x.replace(/<Override[^>]*custom\.xml[^>]*\/>/g, ''));
      }
      const rels = zip.file('_rels/.rels');
      if (rels) {
        const x = await rels.async('string');
        zip.file('_rels/.rels',
          x.replace(/<Relationship[^>]*custom\.xml[^>]*\/>/g, ''));
      }
      changes.push('custom.xml (removed)');
    }
  }

  if (opts.rsid) {
    for (const path of ['word/document.xml', 'word/settings.xml']) {
      const f = zip.file(path);
      if (f) {
        const x = await f.async('string');
        zip.file(path, stripRsidAttrs(x));
        changes.push(path + ' rsid');
      }
    }
  }

  if (opts.textFn) {
    for (const f of zip.file(/^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/)) {
      let x = cleanDocXmlText(await f.async('string'), opts.textFn);
      if (opts.stripFormatting) x = stripDocxFormatting(x);
      zip.file(f.name, x);
    }
    changes.push('text cleaned');
  }

  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  return { blob, changes };
}

// ---- PDF -------------------------------------------------------------------

async function pdfToText(arrayBuffer) {
  const pdfjs = window.pdfjsLib;
  pdfjs.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  const doc = await pdfjs.getDocument({ data: arrayBuffer.slice(0) }).promise;
  let text = '';
  for (let p = 1; p <= doc.numPages; p++) {
    const page = await doc.getPage(p);
    const content = await page.getTextContent();
    text += content.items.map((it) => it.str).join(' ') + '\n\n';
  }
  return text.replace(/\n{3,}/g, '\n\n').trim();
}

async function cleanPdf(arrayBuffer) {
  const { PDFDocument } = window.PDFLib;
  const pdf = await PDFDocument.load(arrayBuffer, { updateMetadata: false });
  const changes = [];
  try {
    pdf.setTitle(''); pdf.setAuthor(''); pdf.setSubject(''); pdf.setKeywords([]);
    pdf.setProducer(''); pdf.setCreator('');
    changes.push('Info dictionary');
  } catch (e) { /* some fields may be absent */ }
  // Remove the XMP metadata stream from the catalog if present.
  try {
    const catalog = pdf.catalog;
    for (const key of ['Metadata', 'PieceInfo']) {
      const name = window.PDFLib.PDFName.of(key);
      if (catalog.has(name)) { catalog.delete(name); changes.push(key); }
    }
  } catch (e) { /* ignore */ }
  const bytes = await pdf.save({ useObjectStreams: false });
  return { blob: new Blob([bytes], { type: 'application/pdf' }), changes };
}

// ---- Images (re-encode via canvas drops EXIF/XMP/IPTC/C2PA/text chunks) ----

function cleanImage(arrayBuffer, mime, ext) {
  return new Promise((res, rej) => {
    const blob = new Blob([arrayBuffer], { type: mime || `image/${ext}` });
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      canvas.getContext('2d').drawImage(img, 0, 0);
      URL.revokeObjectURL(url);
      const outType = ext === 'png' ? 'image/png'
        : ext === 'webp' ? 'image/webp' : 'image/jpeg';
      canvas.toBlob(
        (b) => b ? res({ blob: b, changes: ['EXIF/XMP/IPTC/C2PA stripped (re-encoded)'] })
          : rej(new Error('encode failed')),
        outType, outType === 'image/jpeg' ? 0.95 : undefined);
    };
    img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('image decode failed')); };
    img.src = url;
  });
}

window.FileProcessor = {
  extOf, kindOf, readAsText, readAsArrayBuffer,
  docxToText, pdfToText, cleanDocx, cleanPdf, cleanImage, cleanDocXmlText, stripDocxFormatting,
};
