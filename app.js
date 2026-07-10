const templateSelect = document.getElementById('template-select');
const templateStatus = document.getElementById('template-status');

const dropzone = document.getElementById('dropzone');
const dropzoneFilename = document.getElementById('dropzone-filename');
const reportInput = document.getElementById('report-input');
const fixBtn = document.getElementById('fix-btn');
const heuristicCheckbox = document.getElementById('heuristic-checkbox');
const reportStatus = document.getElementById('report-status');
const downloadLink = document.getElementById('download-link');

const previewContainer = document.getElementById('preview-container');
const previewStyles = document.getElementById('preview-styles');
const previewOriginalBtn = document.getElementById('preview-original-btn');
const previewFixedBtn = document.getElementById('preview-fixed-btn');

let currentProfile = null;
let originalArrayBuffer = null;
let fixedBlob = null;

function renderChip(el, text, isError) {
  if (!text) {
    el.innerHTML = '';
    return;
  }
  const span = document.createElement('span');
  span.className = 'chip' + (isError ? ' error' : ' ok');
  span.textContent = text;
  el.innerHTML = '';
  el.appendChild(span);
}

function setTemplateStatus(text, isError) {
  renderChip(templateStatus, text, isError);
}

function setReportStatus(text, isError) {
  renderChip(reportStatus, text, isError);
}

async function loadTemplateManifest() {
  try {
    const resp = await fetch('./templates/index.json');
    if (!resp.ok) return [];
    const data = await resp.json();
    return Array.isArray(data.templates) ? data.templates : [];
  } catch (e) {
    return [];
  }
}

function populateTemplateSelect(templates) {
  templateSelect.innerHTML = '';
  if (templates.length === 0) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = 'No templates available yet';
    opt.disabled = true;
    opt.selected = true;
    templateSelect.appendChild(opt);
    templateSelect.disabled = true;
    return;
  }
  templateSelect.disabled = false;
  for (const t of templates) {
    const opt = document.createElement('option');
    opt.value = t.file;
    opt.textContent = t.name;
    templateSelect.appendChild(opt);
  }
}

async function useTemplateBuffer(name, arrayBuffer) {
  currentProfile = await window.WordAligner.extractTemplateProfile(arrayBuffer);
  setTemplateStatus(`Using template: ${name}`, false);
  fixBtn.disabled = !reportInput.files.length;
}

async function loadSelectedTemplate() {
  const value = templateSelect.value;
  if (!value) {
    setTemplateStatus('No templates available yet. Add one to templates/index.json.', true);
    currentProfile = null;
    fixBtn.disabled = true;
    return;
  }

  setTemplateStatus('Loading template...', false);
  try {
    const resp = await fetch('./templates/' + value);
    if (!resp.ok) throw new Error('template file not found: ' + value);
    const buf = await resp.arrayBuffer();
    const name = templateSelect.options[templateSelect.selectedIndex].textContent;
    await useTemplateBuffer(name, buf);
  } catch (e) {
    console.error(e);
    setTemplateStatus('Could not load that template: ' + e.message, true);
    currentProfile = null;
    fixBtn.disabled = true;
  }
}

templateSelect.addEventListener('change', loadSelectedTemplate);

async function renderPreview(arrayBufferOrBlob) {
  previewContainer.innerHTML = '';
  try {
    await window.docx.renderAsync(arrayBufferOrBlob, previewContainer, previewStyles, {
      inWrapper: true,
      ignoreLastRenderedPageBreak: true,
    });
  } catch (e) {
    console.error(e);
    previewContainer.innerHTML = '<p class="preview-placeholder">Could not preview this document.</p>';
  }
}

function setActiveToggle(which) {
  previewOriginalBtn.classList.toggle('active', which === 'original');
  previewFixedBtn.classList.toggle('active', which === 'fixed');
}

previewOriginalBtn.addEventListener('click', () => {
  if (!originalArrayBuffer) return;
  setActiveToggle('original');
  renderPreview(originalArrayBuffer.slice(0));
});

previewFixedBtn.addEventListener('click', async () => {
  if (!fixedBlob) return;
  setActiveToggle('fixed');
  renderPreview(await fixedBlob.arrayBuffer());
});

reportInput.addEventListener('change', async () => {
  fixBtn.disabled = !(currentProfile && reportInput.files.length);
  downloadLink.classList.add('hidden');
  setReportStatus('', false);
  fixedBlob = null;
  previewFixedBtn.disabled = true;

  const file = reportInput.files[0];
  if (!file) {
    previewOriginalBtn.disabled = true;
    previewContainer.innerHTML = '<p class="preview-placeholder">Upload a report to preview it here.</p>';
    dropzoneFilename.classList.add('hidden');
    return;
  }

  dropzoneFilename.textContent = file.name;
  dropzoneFilename.classList.remove('hidden');

  originalArrayBuffer = await file.arrayBuffer();
  previewOriginalBtn.disabled = false;
  setActiveToggle('original');
  renderPreview(originalArrayBuffer.slice(0));
});

['dragenter', 'dragover'].forEach(evt => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });
});
['dragleave', 'dragend'].forEach(evt => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
  });
});
dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  const files = e.dataTransfer.files;
  if (!files || !files.length) return;
  reportInput.files = files;
  reportInput.dispatchEvent(new Event('change'));
});

fixBtn.addEventListener('click', async () => {
  const file = reportInput.files[0];
  if (!file || !currentProfile) return;

  fixBtn.disabled = true;
  downloadLink.classList.add('hidden');
  setReportStatus('Fixing alignment...', false);

  try {
    const buf = originalArrayBuffer || await file.arrayBuffer();
    fixedBlob = await window.WordAligner.fixReportDocx(buf, currentProfile, {
      heuristicHeadings: heuristicCheckbox.checked,
    });
    const url = URL.createObjectURL(fixedBlob);
    const outName = file.name.replace(/\.docx$/i, '') + '_aligned.docx';
    downloadLink.href = url;
    downloadLink.download = outName;
    downloadLink.textContent = `Download ${outName}`;
    downloadLink.classList.remove('hidden');
    setReportStatus('Done. Nothing left this browser.', false);

    previewFixedBtn.disabled = false;
    setActiveToggle('fixed');
    await renderPreview(await fixedBlob.arrayBuffer());
  } catch (e) {
    console.error(e);
    setReportStatus('Could not fix that file: ' + e.message, true);
  } finally {
    fixBtn.disabled = !reportInput.files.length;
  }
});

async function init() {
  const templates = await loadTemplateManifest();
  populateTemplateSelect(templates);
  await loadSelectedTemplate();
}

init();
