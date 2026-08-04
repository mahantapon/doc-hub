// Tab 3: Word format-safe editor
// หลักการ: แก้เฉพาะข้อความใน <w:t> ของไฟล์ .docx เดิม ไม่มีการแปลงไฟล์
// ส่วนที่ไม่ได้แก้จะถูกเก็บ byte เดิมไว้ทั้งหมด → format คงเดิม 100%
'use strict';
console.log('[Doc Hub] Word editor build v5 loaded — Enter=ขึ้นบรรทัด, Backspace=ลบบรรทัด');
(() => {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';

  let doc = null; // { zip, fileName, parts:[{path, xmlDecl, dom, origMap, dirty}] }
  let nodeSpan = new WeakMap(); // w:t node -> span (rebuilt ทุกครั้งที่ render)
  let pendingFocus = null;      // { node, offset } เอา caret ไปวางหลัง re-render

  // ---------- caret helpers ----------
  function caretOffset(span) {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return (span.textContent || '').length;
    const r = sel.getRangeAt(0);
    const pre = r.cloneRange();
    pre.selectNodeContents(span);
    try { pre.setEnd(r.endContainer, r.endOffset); } catch (_) { return (span.textContent || '').length; }
    return pre.toString().length;
  }
  function placeCaret(span, off) {
    span.focus();
    const sel = window.getSelection(), r = document.createRange();
    const tn = span.firstChild;
    if (tn && tn.nodeType === 3) r.setStart(tn, Math.min(off, tn.length));
    else r.setStart(span, 0);
    r.collapse(true); sel.removeAllRanges(); sel.addRange(r);
  }

  const partLabel = p =>
    p === 'word/document.xml' ? 'เนื้อหาเอกสาร' :
    /header/.test(p) ? 'หัวกระดาษ' :
    /footer/.test(p) ? 'ท้ายกระดาษ' :
    /footnotes/.test(p) ? 'เชิงอรรถ' :
    /endnotes/.test(p) ? 'อ้างอิงท้ายเรื่อง' : p;

  // ---------- run merge (rendering-neutral normalization) ----------
  const serializeRPr = r => {
    const rPr = [...r.children].find(c => c.namespaceURI === W_NS && c.localName === 'rPr');
    return rPr ? new XMLSerializer().serializeToString(rPr) : '';
  };

  // run ที่มีแค่ rPr + w:t เดียว (ไม่มี tab/br/รูป) ถึงจะ merge ได้
  const isSimpleTextRun = r => {
    let tCount = 0;
    for (const c of r.children) {
      if (c.namespaceURI !== W_NS) return false;
      if (c.localName === 'rPr') continue;
      if (c.localName === 't') { tCount++; continue; }
      return false;
    }
    return tCount === 1;
  };

  // Word ชอบสับ run ย่อย ๆ ทั้งที่ format เท่ากัน → รวมกลับเพื่อให้แก้ข้อความง่าย
  // (รวมเฉพาะ run ติดกันที่ rPr เหมือนกันเป๊ะ = การแสดงผลไม่เปลี่ยนแม้แต่จุดเดียว)
  function mergeAdjacentRuns(dom) {
    const parents = new Set([...dom.getElementsByTagNameNS(W_NS, 'r')].map(r => r.parentNode));
    for (const parent of parents) {
      let prev = null;
      for (const child of [...parent.children]) {
        if (child.namespaceURI === W_NS && child.localName === 'r' && isSimpleTextRun(child)) {
          if (prev && serializeRPr(prev) === serializeRPr(child)) {
            const pt = prev.getElementsByTagNameNS(W_NS, 't')[0];
            const ct = child.getElementsByTagNameNS(W_NS, 't')[0];
            pt.textContent += ct.textContent;
            child.remove();
            continue;
          }
          prev = child;
        } else {
          prev = null;
        }
      }
    }
  }

  // ---------- load ----------
  async function loadDocx(file) {
    try {
      const zip = await JSZip.loadAsync(file);
      if (!zip.file('word/document.xml'))
        throw new Error('ไฟล์นี้ไม่ใช่ .docx (ถ้าเป็น .doc รุ่นเก่า ให้เปิดใน Word แล้ว Save As เป็น .docx ก่อน)');
      const names = Object.keys(zip.files)
        .filter(p => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(p))
        .sort((a, b) => a === 'word/document.xml' ? -1 : b === 'word/document.xml' ? 1 : a.localeCompare(b));
      const parts = [];
      for (const path of names) {
        const text = await zip.file(path).async('string');
        const dom = new DOMParser().parseFromString(text, 'application/xml');
        if (dom.getElementsByTagName('parsererror').length)
          throw new Error('อ่านโครงสร้างไฟล์ไม่ได้ (' + path + ')');
        const xmlDecl = (text.match(/^<\?xml[^>]*\?>/) || [''])[0];
        mergeAdjacentRuns(dom);
        const origMap = new WeakMap();
        for (const t of dom.getElementsByTagNameNS(W_NS, 't')) origMap.set(t, t.textContent);
        parts.push({ path, xmlDecl, dom, origMap, dirty: false });
      }
      doc = { zip, fileName: file.name, parts };
      $('#wordFileName').textContent = file.name;
      $('#wordStart').classList.add('hidden');
      $('#wordWork').classList.remove('hidden');
      setMsg($('#wordMsg'), '');
      setMsg($('#wordOpenMsg'), '');
      renderEditor();
      updateEditCount();
      renderPreview(file);
    } catch (err) {
      setMsg($('#wordOpenMsg'), 'เปิดไฟล์ไม่สำเร็จ: ' + err.message, 'err');
    }
  }

  // ---------- editor ----------
  function paraAtoms(p) {
    const atoms = [];
    const walker = document.createTreeWalker(p, NodeFilter.SHOW_ELEMENT, {
      acceptNode(n) {
        // ข้ามย่อหน้าซ้อน (textbox) — จะถูกแสดงเป็นย่อหน้าของตัวเองอยู่แล้ว
        if (n !== p && n.namespaceURI === W_NS && n.localName === 'p') return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });
    let n;
    while ((n = walker.nextNode())) {
      if (n.namespaceURI !== W_NS) continue;
      if (n.localName === 't') atoms.push({ type: 't', node: n });
      else if (n.localName === 'tab') atoms.push({ type: 'tab' });
      else if (n.localName === 'br') atoms.push({ type: 'br' });
      else if (n.localName === 'drawing') atoms.push({ type: 'img' });
    }
    return atoms;
  }

  function runStyle(t) {
    const css = {};
    let r = t.parentNode;
    while (r && !(r.namespaceURI === W_NS && r.localName === 'r')) r = r.parentNode;
    if (!r) return css;
    const rPr = [...r.children].find(c => c.namespaceURI === W_NS && c.localName === 'rPr');
    if (!rPr) return css;
    const get = ln => [...rPr.children].find(c => c.namespaceURI === W_NS && c.localName === ln);
    const val = el => el ? (el.getAttributeNS(W_NS, 'val') || el.getAttribute('w:val')) : null;
    const b = get('b');
    if (b && val(b) !== '0' && val(b) !== 'false') css.fontWeight = '700';
    const i = get('i');
    if (i && val(i) !== '0' && val(i) !== 'false') css.fontStyle = 'italic';
    const u = get('u');
    if (u && val(u) && val(u) !== 'none') css.textDecoration = 'underline';
    const szv = val(get('sz'));
    if (szv) css.fontSize = Math.max(12, Math.min(30, (parseInt(szv, 10) / 2) * 1.05)) + 'px';
    const cv = val(get('color'));
    if (cv && cv !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(cv)) css.color = '#' + cv;
    return css;
  }

  const isEdited = (part, node) => {
    const o = part.origMap.has(node) ? part.origMap.get(node) : null;
    return o === null ? node.textContent !== '' : node.textContent !== o;
  };

  // Enter = แทรกขึ้นบรรทัดใหม่ (<w:br/>) ตรงตำแหน่ง cursor โดยผ่าครึ่ง w:t เดิม
  function insertBreak(part, node, span) {
    const off = caretOffset(span);
    const text = node.textContent;
    const run = node.parentNode;
    node.textContent = text.slice(0, off);
    const br = part.dom.createElementNS(W_NS, 'w:br');
    const newT = part.dom.createElementNS(W_NS, 'w:t');
    newT.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    newT.textContent = text.slice(off);
    const anchor = node.nextSibling;
    run.insertBefore(br, anchor);
    run.insertBefore(newT, anchor);
    part.dirty = true;
    pendingFocus = { node: newT, offset: 0 };
    renderEditor(); updateEditCount();
  }

  // Backspace ต้นบรรทัด = ลบ <w:br/> ที่อยู่ก่อนหน้า (เอาบรรทัดที่เพิ่งขึ้นออก)
  function removeBreakBefore(part, node) {
    const prev = node.previousElementSibling;
    if (!(prev && prev.namespaceURI === W_NS && prev.localName === 'br')) return false;
    prev.remove();
    part.dirty = true;
    pendingFocus = { node, offset: 0 };
    renderEditor(); updateEditCount();
    return true;
  }

  function makeSpan(part, node) {
    const span = document.createElement('span');
    span.className = 'wt';
    span.contentEditable = 'true';
    span.spellcheck = false;
    span.textContent = node.textContent;
    Object.assign(span.style, runStyle(node));
    span.classList.toggle('edited', isEdited(part, node));
    span.addEventListener('keydown', e => {
      if ((e.metaKey || e.ctrlKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) { e.preventDefault(); return; }
      if (e.key === 'Enter') { e.preventDefault(); insertBreak(part, node, span); return; }
      if (e.key === 'Backspace' && caretOffset(span) === 0) {
        if (removeBreakBefore(part, node)) e.preventDefault();
      }
    });
    span.addEventListener('input', () => {
      node.textContent = span.textContent.replace(/[\r\n]+/g, ' ');
      part.dirty = true;
      span.classList.toggle('edited', isEdited(part, node));
      updateEditCount();
    });
    span.addEventListener('blur', () => {
      if (span.children.length) span.textContent = node.textContent; // ล้าง markup ที่ติดมากับการ paste
    });
    nodeSpan.set(node, span);
    return span;
  }

  function renderEditor() {
    const ed = $('#wordEditor');
    nodeSpan = new WeakMap();
    ed.innerHTML = '';
    for (const part of doc.parts) {
      const rows = [];
      for (const p of [...part.dom.getElementsByTagNameNS(W_NS, 'p')]) {
        const atoms = paraAtoms(p);
        if (!atoms.some(a => a.type === 't' && a.node.textContent !== '')) continue;
        const div = document.createElement('div');
        div.className = 'para';
        for (const a of atoms) {
          if (a.type === 't') div.appendChild(makeSpan(part, a.node));
          else if (a.type === 'tab') div.insertAdjacentHTML('beforeend', '<span class="wsym">⇥</span>');
          else if (a.type === 'br') div.insertAdjacentHTML('beforeend', '<span class="wsym">↵</span><br>');
          else if (a.type === 'img') div.insertAdjacentHTML('beforeend', '<span class="wsym">🖼️</span>');
        }
        rows.push(div);
      }
      if (!rows.length) continue;
      const label = document.createElement('div');
      label.className = 'partlabel';
      label.textContent = partLabel(part.path);
      ed.appendChild(label);
      rows.forEach(r => ed.appendChild(r));
    }
    if (!ed.children.length) ed.innerHTML = '<div class="msg">ไม่พบข้อความที่แก้ได้ในไฟล์นี้</div>';
    if (pendingFocus) {
      const sp = nodeSpan.get(pendingFocus.node);
      if (sp) placeCaret(sp, pendingFocus.offset);
      pendingFocus = null;
    }
  }

  function editedCount() {
    let c = 0;
    for (const part of doc.parts)
      for (const t of part.dom.getElementsByTagNameNS(W_NS, 't')) if (isEdited(part, t)) c++;
    return c;
  }

  function updateEditCount() {
    const c = editedCount();
    const el = $('#editCount');
    el.textContent = 'แก้แล้ว ' + c + ' จุด';
    el.classList.toggle('on', c > 0);
  }

  // ---------- find & replace ----------
  function replaceAll() {
    const find = $('#findText').value;
    const repl = $('#replText').value;
    if (!doc) return;
    if (!find) { setMsg($('#wordMsg'), 'พิมพ์คำที่จะค้นหาก่อน', 'err'); return; }
    let count = 0, cross = 0;
    for (const part of doc.parts) {
      for (const t of [...part.dom.getElementsByTagNameNS(W_NS, 't')]) {
        if (t.textContent.includes(find)) {
          count += t.textContent.split(find).length - 1;
          t.textContent = t.textContent.split(find).join(repl);
          part.dirty = true;
        }
      }
      for (const p of [...part.dom.getElementsByTagNameNS(W_NS, 'p')]) {
        const txt = [...p.getElementsByTagNameNS(W_NS, 't')].map(t => t.textContent).join('');
        if (txt.includes(find)) cross++;
      }
    }
    renderEditor();
    updateEditCount();
    setMsg($('#wordMsg'),
      `แทนที่แล้ว ${count} จุด` +
      (cross ? ` · มีอีก ${cross} ย่อหน้าที่คำนี้ถูกแบ่งคร่อมรูปแบบตัวอักษร ต้องคลิกแก้เองในช่องซ้าย` : ''),
      cross ? 'err' : 'ok');
  }

  // ---------- build ----------
  async function buildDocx() {
    for (const part of doc.parts) {
      if (!part.dirty) continue; // ส่วนที่ไม่ได้แก้ = เก็บ byte เดิมไว้ทั้งไฟล์
      for (const t of [...part.dom.getElementsByTagNameNS(W_NS, 't')]) {
        if (/^\s|\s$/.test(t.textContent)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
      }
      let xml = new XMLSerializer().serializeToString(part.dom);
      if (!xml.startsWith('<?xml'))
        xml = (part.xmlDecl || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>') + '\r\n' + xml;
      doc.zip.file(part.path, xml);
    }
    return doc.zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    });
  }

  // ---------- preview ----------
  async function renderPreview(blob) {
    const cont = $('#wordPreview');
    cont.innerHTML = '<div class="msg">กำลังสร้างตัวอย่าง...</div>';
    try {
      await window.docx.renderAsync(blob, cont, null, { inWrapper: true });
      requestAnimationFrame(fitPreview);
    } catch (e) {
      cont.innerHTML = '<div class="msg err">แสดงตัวอย่างไม่ได้ (ไม่กระทบไฟล์จริง): ' + e.message + '</div>';
    }
  }

  // ย่อหน้ากระดาษพรีวิวให้พอดีความกว้างของกรอบ (ไม่ล้นออกด้านขวา)
  function fitPreview() {
    const cont = $('#wordPreview');
    if (!cont) return;
    const wrap = cont.querySelector('.docx-wrapper');
    if (!wrap) return;
    wrap.style.zoom = '1';
    const page = wrap.querySelector('section') || wrap.firstElementChild;
    if (!page) return;
    const avail = cont.clientWidth - 20;
    if (avail <= 0) return;
    const w = page.offsetWidth;
    if (w > avail) wrap.style.zoom = (avail / w).toFixed(4);
  }
  window.addEventListener('resize', () => { if (doc) fitPreview(); });

  // ---------- wiring ----------
  wireDropzone($('#dropDocx'), $('#fileDocx'), files => {
    const f = files.find(f => /\.docx$/i.test(f.name));
    if (f) loadDocx(f);
    else setMsg($('#wordOpenMsg'), 'รองรับเฉพาะไฟล์ .docx', 'err');
  });
  $('#pickDocx').addEventListener('click', () => $('#fileDocx').click());
  $('#btnReplace').addEventListener('click', replaceAll);
  $('#btnRefresh').addEventListener('click', async () => {
    if (!doc) return;
    renderPreview(await buildDocx());
  });
  $('#btnDownload').addEventListener('click', async () => {
    if (!doc) return;
    downloadBlob(await buildDocx(), doc.fileName);
    setMsg($('#wordMsg'), 'ดาวน์โหลดแล้ว — format เดิม 100% ส่งต่อได้เลย', 'ok');
  });
  $('#btnCloseDoc').addEventListener('click', () => {
    doc = null;
    $('#wordEditor').innerHTML = '';
    $('#wordPreview').innerHTML = '';
    $('#wordWork').classList.add('hidden');
    $('#wordStart').classList.remove('hidden');
  });

  // ---------- dev hooks (สำหรับทดสอบอัตโนมัติ ไม่กระทบการใช้งานจริง) ----------
  window.__doc = {
    get doc() { return doc; },
    buildDocx,
    loadDocx,
    async b64() {
      const bl = await buildDocx();
      const buf = new Uint8Array(await bl.arrayBuffer());
      let s = '';
      for (let i = 0; i < buf.length; i += 0x8000)
        s += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
      return btoa(s);
    }
  };
  const qp = new URLSearchParams(location.search);
  if (qp.has('dev')) {
    const f = qp.get('dev') && qp.get('dev') !== '1' ? qp.get('dev') : 'test.docx';
    fetch('test/' + f)
      .then(r => { if (!r.ok) throw new Error('no test file'); return r.blob(); })
      .then(b => { switchTab('word'); return loadDocx(new File([b], f)); })
      .catch(e => console.warn('dev load failed:', e.message));
  }
})();
