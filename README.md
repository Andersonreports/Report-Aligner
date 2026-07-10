# Report Aligner

A small browser-based tool that fixes Word (`.docx`) report formatting to match a
reference template — fonts, paragraph alignment, spacing, margins, and heading
styles. Runs entirely client-side (via [JSZip](https://stuk.github.io/jszip/) and
direct OOXML manipulation); no report is ever uploaded to a server.

Hosted on GitHub Pages directly from the repo root (Settings → Pages → Source:
"Deploy from a branch" → `main` / `/ (root)`).

## How it works

A `.docx` file is a zip of XML parts. The tool:

1. Reads the reference template's `word/styles.xml` to learn each named style's
   (Normal, Heading 2, Title, etc.) effective font/alignment/spacing — resolving
   Word's `basedOn` inheritance chain, since many styles only override a few
   properties and inherit the rest.
2. Reads the template's page setup (margins, size, orientation) and header.
3. Applies all of that to the uploaded report: page setup on every section,
   paragraph/run formatting on every paragraph (body, tables, headers/footers),
   and — if the report has no header at all — copies the template's header in.
4. Because real-world drafts often tag the same heading with inconsistent (or
   no) Word style names, paragraphs are also matched by their **text** against
   a small dictionary of known recurring section headings (see `TITLE_TEXTS` /
   `HEADING2_TEXTS` in `aligner.js`) as a more reliable fallback than style
   names alone. The very first non-empty paragraph in the document is always
   treated as the title.

### What it deliberately does NOT touch

Table borders, shading, and cell colors are left exactly as the report author
set them. In these reports, table/callout-box color often carries clinical
meaning (e.g. a green vs. red result box) — normalizing all tables to one
"template table style" would risk changing what a report visually communicates.
Table *cell paragraph* formatting (font, alignment, spacing) is still fixed.

## Using it

1. Open the site. Pick a template from the dropdown — templates are managed
   in the repo (see below), not uploaded by end users.
2. Upload a report to fix — it previews immediately on the right.
3. Click **Fix Alignment**. The preview switches to the fixed version (toggle
   back to "Original" to compare) and a download button appears.

There's an optional "heuristic headings" checkbox that tries to promote
paragraphs that *look* like headings (short, bold, larger than body text) but
aren't tagged with any heading style. It's best-effort and off by default.

Document preview (both original and fixed) is rendered client-side via
[docx-preview](https://github.com/VolodymyrBaydalka/docxjs).

### Managing templates

There is deliberately no way to upload a template from the site itself —
templates only change when someone commits a change to `templates/`. See
`templates/README.md`: drop a `.docx` in `templates/` and add an entry to
`templates/index.json`. **Never commit a real patient report as a template**;
use a de-identified file. This repo is public.

## Local development

```
python -m http.server 8934   # or any static file server, from the repo root
```

Then open http://localhost:8934/.

### Regression testing the engine

`dev/` has a Node-based harness (polyfills `DOMParser`/`XMLSerializer` via
`@xmldom/xmldom` since the engine is written for the browser) that runs the
real `aligner.js` against generated fixture `.docx` files:

```
cd dev && npm install && python make_fixtures.py   # generates scratch fixtures via python-docx
node test_harness.js
```

`test_real.js` is set up to test against real report files placed at the repo
root — those files are gitignored (`*.docx`, `*.pdf`) and must never be
committed, since real reports contain patient PHI.
