// Doc Hub — shared helpers + tab switching
'use strict';

function $(sel) { return document.querySelector(sel); }
function $all(sel) { return [...document.querySelectorAll(sel)]; }

function switchTab(name) {
  $all('.tab').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
  $all('.panel').forEach(p => p.classList.toggle('active', p.id === 'tab-' + name));
}

$all('.tab').forEach(btn => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));

function downloadBlob(blob, filename) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

function setMsg(el, text, cls) {
  el.textContent = text || '';
  el.className = 'msg' + (cls ? ' ' + cls : '');
}

// generic drag & drop wiring
function wireDropzone(zoneEl, inputEl, onFiles) {
  zoneEl.addEventListener('dragover', e => { e.preventDefault(); zoneEl.classList.add('over'); });
  zoneEl.addEventListener('dragleave', () => zoneEl.classList.remove('over'));
  zoneEl.addEventListener('drop', e => {
    e.preventDefault();
    zoneEl.classList.remove('over');
    if (e.dataTransfer.files.length) onFiles([...e.dataTransfer.files]);
  });
  inputEl.addEventListener('change', () => {
    if (inputEl.files.length) onFiles([...inputEl.files]);
    inputEl.value = '';
  });
}
