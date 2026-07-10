# Report templates

Each entry here is a reference `.docx` whose formatting (fonts, alignment,
spacing, margins, headings) gets applied to uploaded reports.

To add a template:

1. Put the `.docx` file in this folder (e.g. `no_pathogenic.docx`).
2. Add an entry to `index.json`:

```json
{
  "templates": [
    { "name": "No pathogenic with additional", "file": "no_pathogenic.docx" }
  ]
}
```

3. Commit and push — it'll show up in the dropdown on the site.

**Do not commit a real patient report as a template.** Use a de-identified
file (placeholder patient name/PIN/etc.) — this repo is public.
