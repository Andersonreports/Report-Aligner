const DB_NAME = 'report-aligner';
const STORE = 'templates';
const DEFAULT_TEMPLATE_URL = './team-template/template.docx';

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => req.result.createObjectStore(STORE);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveTemplateToDb(name, arrayBuffer) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put({ name, arrayBuffer }, 'current');
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function loadTemplateFromDb() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).get('current');
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

let currentProfile = null;
let currentTemplateName = null;

const templateStatus = document.getElementById('template-status');
const templateInput = document.getElementById('template-input');
const reportInput = document.getElementById('report-input');
const fixBtn = document.getElementById('fix-btn');
const heuristicCheckbox = document.getElementById('heuristic-checkbox');
const reportStatus = document.getElementById('report-status');
const downloadLink = document.getElementById('download-link');

function setTemplateStatus(text, isError) {
  templateStatus.textContent = text;
  templateStatus.className = isError ? 'status error' : 'status ok';
}

function setReportStatus(text, isError) {
  reportStatus.textContent = text;
  reportStatus.className = isError ? 'status error' : 'status ok';
}

async function useTemplateBuffer(name, arrayBuffer) {
  currentProfile = await window.WordAligner.extractTemplateProfile(arrayBuffer);
  currentTemplateName = name;
  setTemplateStatus(`Using template: ${name}`, false);
  fixBtn.disabled = !reportInput.files.length;
}

async function init() {
  try {
    const cached = await loadTemplateFromDb();
    if (cached) {
      await useTemplateBuffer(cached.name, cached.arrayBuffer);
      return;
    }
  } catch (e) {
    // IndexedDB unavailable or empty; fall through to fetching the default.
  }

  try {
    const resp = await fetch(DEFAULT_TEMPLATE_URL);
    if (resp.ok) {
      const buf = await resp.arrayBuffer();
      await useTemplateBuffer('team-template/template.docx (bundled)', buf);
      return;
    }
  } catch (e) {
    // No bundled template shipped with this site; wait for a manual upload.
  }

  setTemplateStatus('No template loaded yet — upload your team\'s template below.', true);
}

templateInput.addEventListener('change', async () => {
  const file = templateInput.files[0];
  if (!file) return;
  setTemplateStatus('Reading template...', false);
  try {
    const buf = await file.arrayBuffer();
    await useTemplateBuffer(file.name, buf);
    await saveTemplateToDb(file.name, buf);
  } catch (e) {
    console.error(e);
    setTemplateStatus('Could not read that template: ' + e.message, true);
  }
});

reportInput.addEventListener('change', () => {
  fixBtn.disabled = !(currentProfile && reportInput.files.length);
  downloadLink.classList.add('hidden');
  setReportStatus('', false);
});

fixBtn.addEventListener('click', async () => {
  const file = reportInput.files[0];
  if (!file || !currentProfile) return;

  fixBtn.disabled = true;
  downloadLink.classList.add('hidden');
  setReportStatus('Fixing alignment...', false);

  try {
    const buf = await file.arrayBuffer();
    const blob = await window.WordAligner.fixReportDocx(buf, currentProfile, {
      heuristicHeadings: heuristicCheckbox.checked,
    });
    const url = URL.createObjectURL(blob);
    const outName = file.name.replace(/\.docx$/i, '') + '_aligned.docx';
    downloadLink.href = url;
    downloadLink.download = outName;
    downloadLink.textContent = `Download ${outName}`;
    downloadLink.classList.remove('hidden');
    setReportStatus('Done. Nothing left this browser.', false);
  } catch (e) {
    console.error(e);
    setReportStatus('Could not fix that file: ' + e.message, true);
  } finally {
    fixBtn.disabled = !reportInput.files.length;
  }
});

init();
