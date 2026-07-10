const fs = require('fs');
const path = require('path');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const JSZip = require('jszip');

global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;
global.JSZip = JSZip;
global.window = global;

const alignerSrc = fs.readFileSync(path.join(__dirname, '..', 'web', 'aligner.js'), 'utf8');
eval(alignerSrc);

async function main() {
  const templateBuf = fs.readFileSync(path.join(__dirname, 'template.docx'));
  const reportBuf = fs.readFileSync(path.join(__dirname, 'report.docx'));

  const profile = await window.WordAligner.extractTemplateProfile(templateBuf);
  console.log('--- profile (trimmed) ---');
  console.log(JSON.stringify({ ...profile, tableStyleXmlById: Object.keys(profile.tableStyleXmlById) }, null, 2));

  const outBlobLike = await window.WordAligner.fixReportDocx(reportBuf, profile, { heuristicHeadings: true });
  // node's jszip generateAsync 'blob' isn't valid outside browser; use nodebuffer instead for the test
  const zip = await JSZip.loadAsync(reportBuf);
  fs.writeFileSync(path.join(__dirname, 'fixed_report_raw.bin'), Buffer.from(await outBlobLike.arrayBuffer ? await outBlobLike.arrayBuffer() : outBlobLike));
  console.log('wrote fixed_report_raw.bin');
}

main().catch(e => { console.error(e); process.exit(1); });
