// sanitizer.js — Text cleanup: hidden/invisible character stripping, watermark
// removal, and typography normalization. Each transform is individually
// toggleable. Also computes Term Lock protected spans and an X-ray diff.
//
// Scope note: this file intentionally contains ONLY privacy/cleanup transforms
// (removing hidden data and normalizing typography). It deliberately does not
// inject errors, fabricate "human" imperfection, or try to defeat detectors.

'use strict';

// ---- Character sets / maps -------------------------------------------------

const ZERO_WIDTH = new Set([
  0x200b, 0x200c, 0x200d, 0x2060, 0xfeff, 0x00ad, 0x180e,
]);

// Bidirectional controls + ASCII control remnants (tab/newline/CR kept).
function isBidiOrControl(cp) {
  if (cp === 0x200e || cp === 0x200f) return true;
  if (cp >= 0x202a && cp <= 0x202e) return true;
  if (cp >= 0x2066 && cp <= 0x2069) return true;
  if (cp === 0x09 || cp === 0x0a || cp === 0x0d) return false;
  if (cp <= 0x08 || cp === 0x0b || cp === 0x0c) return true;
  if (cp >= 0x0e && cp <= 0x1f) return true;
  if (cp === 0x7f) return true;
  return false;
}

// Variation selectors + Unicode tag block (steganographic payload carriers).
function isVariationOrTag(cp) {
  if (cp >= 0xfe00 && cp <= 0xfe0f) return true;
  if (cp >= 0xe0100 && cp <= 0xe01ef) return true;
  if (cp >= 0xe0000 && cp <= 0xe007f) return true;
  return false;
}

// Exotic whitespace -> ASCII space.
const EXOTIC_SPACE = new Set([0x00a0, 0x202f, 0x2007, 0x3000]);
function isExoticSpace(cp) {
  if (EXOTIC_SPACE.has(cp)) return true;
  if (cp >= 0x2000 && cp <= 0x200a) return true; // en/em/thin/hair spaces
  return false;
}

// Smart quotes / primes -> straight ASCII.
const SMART_QUOTES = {
  0x201c: '"', 0x201d: '"', 0x201e: '"', 0x00ab: '"', 0x00bb: '"', 0x2033: '"',
  0x2018: "'", 0x2019: "'", 0x201a: "'", 0x2039: "'", 0x203a: "'", 0x2032: "'",
};

// Typographic ligatures -> ASCII.
const LIGATURES = {
  0xfb00: 'ff', 0xfb01: 'fi', 0xfb02: 'fl', 0xfb03: 'ffi', 0xfb04: 'ffl', 0xfb05: 'st', 0xfb06: 'st',
};

// Common confusables (Cyrillic/Greek/math lookalikes) -> ASCII Latin.
const CONFUSABLES = {
  0x0430: 'a', 0x0435: 'e', 0x043e: 'o', 0x0440: 'p', 0x0441: 'c', 0x0443: 'y', 0x0445: 'x',
  0x0410: 'A', 0x0412: 'B', 0x0415: 'E', 0x041a: 'K', 0x041c: 'M', 0x041d: 'H', 0x041e: 'O',
  0x0420: 'P', 0x0421: 'C', 0x0422: 'T', 0x0425: 'X', 0x0456: 'i', 0x0458: 'j', 0x0405: 'S',
  0x0391: 'A', 0x0392: 'B', 0x0395: 'E', 0x0396: 'Z', 0x0397: 'H', 0x0399: 'I', 0x039a: 'K',
  0x039c: 'M', 0x039d: 'N', 0x039f: 'O', 0x03a1: 'P', 0x03a4: 'T', 0x03a5: 'Y', 0x03a7: 'X',
  0x03bf: 'o', 0x03b1: 'a', 0x0192: 'f', 0x2212: '-', 0x2215: '/', 0x2044: '/',
};

// ---- Term Lock (protected spans) ------------------------------------------

const PROTECT_PATTERNS = {
  urls: /\bhttps?:\/\/[^\s<>"')\]]+/gi,
  emails: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g,
  code: /`[^`]*`/g,
  citations: /\[\d+(?:\s*[,–-]\s*\d+)*\]|\([A-Z][A-Za-z.'-]+(?:\s+(?:et al\.?|and|&)\s+[A-Z][A-Za-z.'-]+)*,\s*\d{4}[a-z]?\)/g,
  math: /\$\$[\s\S]*?\$\$|\$[^$\n]+\$|\\\([\s\S]*?\\\)|\\\[[\s\S]*?\\\]/g,
};

// Returns a boolean array of length text.length; true = locked (don't touch).
function computeProtectedMask(text, protectOpts) {
  const mask = new Array(text.length).fill(false);
  const mark = (re) => {
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(text)) !== null) {
      for (let i = m.index; i < m.index + m[0].length; i++) mask[i] = true;
      if (m.index === re.lastIndex) re.lastIndex++; // guard zero-length
    }
  };
  for (const key of ['urls', 'emails', 'code', 'citations', 'math']) {
    if (protectOpts[key]) mark(PROTECT_PATTERNS[key]);
  }
  if (protectOpts.keywords && protectOpts.keywordList) {
    for (const raw of protectOpts.keywordList) {
      const kw = raw.trim();
      if (!kw) continue;
      const esc = kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      mark(new RegExp(esc, 'g'));
    }
  }
  return mask;
}

// ---- Formatting stripper (Markdown + HTML -> plain text) -------------------

function stripFormatting(text) {
  let s = text;
  s = s.replace(/```[^\n]*\n([\s\S]*?)```/g, '$1');      // fenced code -> inner
  s = s.replace(/~~~[^\n]*\n([\s\S]*?)~~~/g, '$1');
  s = s.replace(/!\[([^\]]*)\]\([^)]*\)/g, '$1');          // image -> alt
  s = s.replace(/\[([^\]]+)\]\([^)]*\)/g, '$1');           // link -> text
  s = s.replace(/^\s{0,3}#{1,6}\s+/gm, '');                // headings
  s = s.replace(/^\s{0,3}>\s?/gm, '');                     // blockquotes
  s = s.replace(/^\s*\|?[-:\s|]+\|[-:\s|]*$/gm, '');       // table separator rows
  s = s.replace(/^\s{0,3}([-*+]|\d+[.)])\s+/gm, '');       // list markers
  s = s.replace(/^\s{0,3}([-*_])(?:\s*\1){2,}\s*$/gm, ''); // horizontal rules
  s = s.replace(/(\*\*|__)(.*?)\1/g, '$2');                // bold
  s = s.replace(/(\*|_)(.*?)\1/g, '$2');                   // italic
  s = s.replace(/~~(.*?)~~/g, '$1');                       // strikethrough
  s = s.replace(/`([^`]+)`/g, '$1');                       // inline code
  s = s.replace(/\s*\|\s*/g, ' ');                         // table cell pipes
  s = s.replace(/<\/?[a-zA-Z][^>]*>/g, '');                // HTML tags
  s = s.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
  return s;
}

// ---- Main sanitize pass ----------------------------------------------------
// opts: { transforms: {id:bool}, dashReplacement, homoglyphNFKC, protect:{...} }
// Returns { output, segments, counts } where segments drive the X-ray diff.

function sanitizeText(text, opts) {
  const t = opts.transforms || {};
  if (t.stripFormatting) text = stripFormatting(text);
  const mask = computeProtectedMask(text, opts.protect || {});
  const dashRepl = opts.dashReplacement != null ? opts.dashReplacement : ' - ';

  const segments = []; // {text, type|null}
  const counts = { hidden: 0, typography: 0, confusable: 0 };
  let out = '';

  const pushUnchanged = (s) => {
    if (!s) return;
    out += s;
    const last = segments[segments.length - 1];
    if (last && last.type === null) last.text += s;
    else segments.push({ text: s, type: null });
  };
  const pushChanged = (orig, repl, type) => {
    out += repl;
    segments.push({ text: repl, orig, type });
  };

  let i = 0;
  while (i < text.length) {
    const cp = text.codePointAt(i);
    const charLen = cp > 0xffff ? 2 : 1;
    const ch = text.slice(i, i + charLen);

    if (mask[i]) { pushUnchanged(ch); i += charLen; continue; }

    let repl = null;   // replacement string
    let type = null;   // diff category

    if (t.zeroWidth && ZERO_WIDTH.has(cp)) { repl = ''; type = 'hidden'; }
    else if (t.bidiControl && isBidiOrControl(cp)) { repl = ''; type = 'hidden'; }
    else if (t.variationTags && isVariationOrTag(cp)) { repl = ''; type = 'hidden'; }
    else if (t.homoglyph && CONFUSABLES[cp] != null) { repl = CONFUSABLES[cp]; type = 'confusable'; }
    else if (t.homoglyph && opts.homoglyphNFKC) {
      const n = ch.normalize('NFKC');
      if (n !== ch && /^[\x20-\x7e]*$/.test(n)) { repl = n; type = 'confusable'; }
    }
    else if (t.whitespace && isExoticSpace(cp)) { repl = ' '; type = 'typography'; }
    else if (t.smartQuotes && SMART_QUOTES[cp] != null) { repl = SMART_QUOTES[cp]; type = 'typography'; }
    else if (t.dashes && (cp === 0x2014 || cp === 0x2013)) { repl = dashRepl; type = 'typography'; }
    else if (t.ellipsisLigature && cp === 0x2026) { repl = '...'; type = 'typography'; }
    else if (t.ellipsisLigature && LIGATURES[cp] != null) { repl = LIGATURES[cp]; type = 'typography'; }

    if (repl === null) { pushUnchanged(ch); }
    else {
      pushChanged(ch, repl, type);
      if (type === 'hidden') counts.hidden++;
      else if (type === 'confusable') counts.confusable++;
      else counts.typography++;
    }
    i += charLen;
  }

  // Collapse doubled spaces created when a space-flanked dash becomes " - ".
  // Only touches runs of 2+ spaces next to a hyphen, so "word-internal" stays.
  if (t.dashes && dashRepl === ' - ') out = out.replace(/ {2,}-/g, ' -').replace(/- {2,}/g, '- ');

  return { output: out, segments, counts };
}

// Count hidden/watermark characters present in text (for the risk scoreboard).
function countHidden(text) {
  let n = 0;
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i);
    const len = cp > 0xffff ? 2 : 1;
    if (ZERO_WIDTH.has(cp) || isBidiOrControl(cp) || isVariationOrTag(cp)) n++;
    i += len;
  }
  return n;
}

window.Sanitizer = { sanitizeText, countHidden, computeProtectedMask, stripFormatting };
