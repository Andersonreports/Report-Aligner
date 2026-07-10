# Report Aligner

A small browser-based tool that fixes Word (`.docx`) report formatting to match a
reference template — fonts, paragraph alignment, spacing, margins, and heading
styles. Runs entirely client-side (via [JSZip](https://stuk.github.io/jszip/) and
direct OOXML manipulation); no report is ever uploaded to a server.

Hosted on GitHub Pages from the `web/` folder (see `.github/workflows/pages.yml`).

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
   `HEADING2_TEXTS` in `web/aligner.js`) as a more reliable fallback than style
   names alone. The very first non-empty paragraph in the document is always
   treated as the title.

### What it deliberately does NOT touch

Table borders, shading, and cell colors are left exactly as the report author
set them. In these reports, table/callout-box color often carries clinical
meaning (e.g. a green vs. red result box) — normalizing all tables to one
"template table style" would risk changing what a report visually communicates.
Table *cell paragraph* formatting (font, alignment, spacing) is still fixed.

## Using it

1. Open the site.
2. Upload your team's reference `.docx` template (once — it's remembered in
   that browser via IndexedDB for next time).
3. Upload a report to fix, click **Fix Alignment**, download the result.

There's an optional "heuristic headings" checkbox that tries to promote
paragraphs that *look* like headings (short, bold, larger than body text) but
aren't tagged with any heading style. It's best-effort and off by default.

### Shipping a default template for the whole team

By default each person uploads the template once in their own browser. If you
want everyone to get the correct template automatically with no manual step,
commit a **de-identified** template (placeholder patient info, not a real
report) to `web/team-template/template.docx` — the app fetches it on load if
present. Do not commit a real patient report here; this repo is public.

## Local development

```
cd web && python -m http.server 8934   # or any static file server
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
