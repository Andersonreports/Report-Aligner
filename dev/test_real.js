const fs = require('fs');
const path = require('path');
const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');
const JSZip = require('jszip');

global.DOMParser = DOMParser;
global.XMLSerializer = XMLSerializer;
global.JSZip = JSZip;
global.window = global;

const alignerSrc = fs.readFileSync(path.join(__dirname, '..', 'aligner.js'), 'utf8');
eval(alignerSrc);

const ROOT = path.join(__dirname, '..');

async function run(draftName) {
  console.log('\n' + '='.repeat(80));
  console.log('Fixing:', draftName);
  const templateBuf = fs.readFileSync(path.join(ROOT, 'Baby_Helina_Princy_Whole_Exome_Whole_Mitochondrial_report_15_06.docx'));
  const draftBuf = fs.readFileSync(path.join(ROOT, draftName));

  const profile = await window.WordAligner.extractTemplateProfile(templateBuf);
  console.log('template has default header:', !!profile.defaultHeader, profile.defaultHeader ? `(${profile.defaultHeader.media.length} media)` : '');

  const blob = await window.WordAligner.fixReportDocx(draftBuf, profile, { heuristicHeadings: true });
  const outPath = path.join(__dirname, draftName.replace('.docx', '_FIXED.docx'));
  fs.writeFileSync(outPath, Buffer.from(await blob.arrayBuffer()));
  console.log('wrote', outPath);
}

async function main() {
  await run('GEN_T_2.0.spik_4099_outfastp_R.docx');
  await run('GEN_T_2.0.spik_4385_outfastp_R.docx');
}

main().catch(e => { console.error(e); process.exit(1); });
