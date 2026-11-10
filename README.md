# [DocScrub](https://azp3001.github.io/Humanize/)

> Live site: **https://azp3001.github.io/Humanize/**

A 100% client-side, zero-build web app for cleaning and anonymizing documents.
Hosts directly on GitHub Pages (no npm, no bundler).

**What it does**
- Strips hidden/watermark characters (zero-width, bidi/control, variation
  selectors & Unicode tag payloads) and normalizes homoglyphs/confusables.
- Normalizes typography (exotic whitespace, smart quotes, em/en dashes,
  ellipsis, ligatures).
- **Removes** metadata from `.docx`, `.pdf`, and images (EXIF/XMP/C2PA, core/app
  props, rsid revision marks). Removal only — it does not fabricate replacement
  authoring history.
- Term Lock to protect URLs, emails, code, citations, math, and custom keywords.
- Live dual-pane editor, X-ray diff, scoreboard, and single/batch export.

Every feature has its own toggle, plus category master switches and presets.

**Stack:** static HTML/CSS/vanilla ES6. CDN libs: JSZip, pdf-lib, pdf.js.
Open `index.html` or serve the folder — no build step.

## Scope note
This is a privacy/cleanup tool. It intentionally omits any feature intended to
fabricate forensic history or evade AI-text detection. If you use AI as an
accessibility aid, the robust route is a disclosed disability accommodation —
cleaning hidden data from your own files is fine; forging provenance is not.
