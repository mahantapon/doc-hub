// Word format-safe editor — แก้ข้อความ "ในหน้าเอกสารเป๊ะ" (docx-preview ที่ patch ให้ติดเลขอ้างอิง)
// หลักการเดิม: แก้เฉพาะข้อความใน <w:t> ของไฟล์เดิม ไม่มีการแปลงไฟล์ → format คงเดิม 100%
// การจับคู่: ตัววาดถูก patch ให้จดลำดับ w:t ตอนอ่าน (window.__dxTexts) และห่อทุกก้อนข้อความ
// ด้วย span.dx-t ที่มี __dxTid → จับคู่กลับไป w:t ในไฟล์ได้แบบเป๊ะ มีตรวจสอบก่อนเปิดให้แก้
// จุดไหนตรวจไม่ผ่าน = ล็อกอ่านอย่างเดียว (กันไฟล์พัง)
'use strict';
(() => {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

  // ตัววาด (docx-preview) อ่าน mc:AlternateContent เฉพาะลูกตัวแรกของฝั่ง Fallback เสมอ
  // → รายการ w:t ของเราต้องเลือกฝั่งเดียวกัน ไม่งั้นลำดับไม่ตรงและจับคู่ไม่ได้
  // และข้อความใน AlternateContent มี 2 สำเนา (Word แสดงฝั่ง Choice, ตัววาดแสดง Fallback)
  // ถ้าแก้สำเนาเดียว Word จะไม่เห็นการแก้ → ล็อกไว้อ่านอย่างเดียวเพื่อความถูกต้อง
  function collectTNodes(dom) {
    const excluded = [];
    const lockedRoots = [...dom.getElementsByTagNameNS(MC_NS, 'AlternateContent')];
    for (const ac of lockedRoots) {
      const fb = [...ac.children].find(c => c.namespaceURI === MC_NS && c.localName === 'Fallback');
      const chosen = fb ? fb.firstElementChild : null;
      for (const child of ac.children) {
        if (child === fb) {
          for (const sub of child.children) if (sub !== chosen) excluded.push(sub);
        } else excluded.push(child);
      }
    }
    const hasAncestorIn = (n, list) => {
      for (let a = n.parentNode; a; a = a.parentNode) if (list.includes(a)) return true;
      return false;
    };
    const tNodes = [...dom.getElementsByTagNameNS(W_NS, 't')].filter(t => !hasAncestorIn(t, excluded));
    const locked = new Set(tNodes.filter(t => hasAncestorIn(t, lockedRoots)));
    return { tNodes, locked };
  }

  let doc = null; // { zip, fileName, parts:[{path,xmlDecl,dom,tNodes,orig,dirty}] }
  let tidMap = [];      // tid -> { part, idx }
  let tidSpans = new Map(); // tid -> [span, ...] (header/footer ซ้ำได้หลายหน้า)

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
        // หมายเหตุ: "ไม่" รวม run ที่ติดกัน — ลำดับ w:t ต้องตรงกับตัววาดแบบตัวต่อตัว
        const { tNodes, locked } = collectTNodes(dom);
        parts.push({ path, xmlDecl, dom, tNodes, locked, orig: tNodes.map(t => t.textContent), dirty: false });
      }
      doc = { zip, fileName: file.name, parts };
      $('#wordFileName').textContent = file.name;
      $('#wordStart').classList.add('hidden');
      $('#wordWork').classList.remove('hidden');
      setMsg($('#wordMsg'), '');
      setMsg($('#wordOpenMsg'), '');
      await renderView();
      updateEditCount();
    } catch (err) {
      setMsg($('#wordOpenMsg'), 'เปิดไฟล์ไม่สำเร็จ: ' + err.message, 'err');
    }
  }

  // ---------- build ----------
  async function buildDocx() {
    for (const part of doc.parts) {
      if (!part.dirty) continue; // ส่วนที่ไม่ได้แก้ = เก็บ byte เดิมไว้ทั้งไฟล์
      for (const t of [...part.dom.getElementsByTagNameNS(W_NS, 't')])
        if (/^\s|\s$/.test(t.textContent)) t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
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

  // ---------- mapping: tid -> (part, idx) ----------
  // __dxTexts คือรายการข้อความตามลำดับที่ตัววาด "อ่านจากไฟล์จริง" — แต่ละ part ของเรา
  // ต้องปรากฏเป็นช่วงต่อเนื่องในนั้น (ข้อความเหมือนกันตัวต่อตัว) จับคู่แบบไม่ทับซ้อน
  function computeTidMap(dxTexts) {
    const map = new Array(dxTexts.length).fill(null);
    const claimed = new Array(dxTexts.length).fill(false);
    const partsSorted = [...doc.parts].sort((a, b) => b.tNodes.length - a.tNodes.length);
    for (const part of partsSorted) {
      const seq = part.tNodes.map(t => t.textContent);
      if (!seq.length) continue;
      let found = -1;
      for (let i = 0; i + seq.length <= dxTexts.length; i++) {
        let ok = true;
        for (let j = 0; j < seq.length; j++) {
          if (claimed[i + j] || dxTexts[i + j] !== seq[j]) { ok = false; break; }
        }
        if (ok) { found = i; break; }
      }
      if (found < 0) continue; // part นี้จับคู่ไม่ได้ → span ของมันจะถูกล็อก
      for (let j = 0; j < seq.length; j++) {
        claimed[found + j] = true;
        map[found + j] = { part, idx: j };
      }
    }
    return map;
  }

  // ---------- render (single exact view, editable) ----------
  async function renderView() {
    const cont = $('#wordCanvas');
    cont.innerHTML = '<div class="msg" style="padding:14px">กำลังวาดเอกสาร...</div>';
    window.__dxTexts = [];
    const blob = await buildDocx();
    cont.innerHTML = '';
    await window.docx.renderAsync(blob, cont, null, { inWrapper: true });
    tidMap = computeTidMap(window.__dxTexts);
    window.__dxTexts = null; // ปิด log ระหว่างที่ไม่ได้ render
    wireSpans(cont);
  }

  function wireSpans(cont) {
    tidSpans = new Map();
    let editable = 0, locked = 0;
    for (const span of [...cont.getElementsByClassName('dx-t')]) {
      const tid = span.__dxTid;
      const m = (tid !== undefined) ? tidMap[tid] : null;
      // ด่านตรวจ: ข้อความบนจอต้องตรงกับในไฟล์เป๊ะ + ไม่ใช่โซนล็อก ถึงจะเปิดให้แก้
      if (!m || span.textContent !== m.part.tNodes[m.idx].textContent ||
          m.part.locked.has(m.part.tNodes[m.idx])) {
        if (span.textContent.trim() !== '') locked++;
        continue;
      }
      editable++;
      if (!tidSpans.has(tid)) tidSpans.set(tid, []);
      tidSpans.get(tid).push(span);
      span.contentEditable = 'true';
      span.spellcheck = false;
      span.classList.toggle('edited', m.part.tNodes[m.idx].textContent !== m.part.orig[m.idx]);
      span.addEventListener('keydown', e => {
        if (e.key === 'Enter') e.preventDefault();
        if ((e.metaKey || e.ctrlKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) e.preventDefault();
      });
      span.addEventListener('input', () => {
        const txt = span.textContent.replace(/[\r\n]+/g, ' ');
        m.part.tNodes[m.idx].textContent = txt;
        m.part.dirty = true;
        const changed = txt !== m.part.orig[m.idx];
        for (const s of tidSpans.get(tid)) {         // header/footer โผล่หลายหน้า → sync กัน
          if (s !== span) s.textContent = txt;
          s.classList.toggle('edited', changed);
        }
        updateEditCount();
      });
      span.addEventListener('blur', () => {
        if (span.children.length) span.textContent = m.part.tNodes[m.idx].textContent;
      });
    }
    if (locked > 0)
      setMsg($('#wordMsg'), `พร้อมแก้ไข ${editable} จุด · มี ${locked} จุดที่ระบบล็อกไว้ (จับคู่กับไฟล์ไม่ได้ จึงไม่เปิดให้แก้เพื่อความปลอดภัยของไฟล์)`, '');
    window.__wireStats = { editable, locked };
  }

  function editedCount() {
    let c = 0;
    for (const part of doc.parts) part.tNodes.forEach((n, i) => { if (n.textContent !== part.orig[i]) c++; });
    return c;
  }
  function updateEditCount() {
    const c = editedCount(), el = $('#editCount');
    el.textContent = 'แก้แล้ว ' + c + ' จุด';
    el.classList.toggle('on', c > 0);
  }

  // ---------- find & replace ----------
  async function replaceAll() {
    if (!doc) return;
    const find = $('#findText').value, repl = $('#replText').value;
    if (!find) { setMsg($('#wordMsg'), 'พิมพ์คำที่จะค้นหาก่อน', 'err'); return; }
    let count = 0, cross = 0;
    for (const part of doc.parts) {
      for (const t of part.tNodes) {
        if (part.locked.has(t)) continue; // โซนกล่องลอยมี 2 สำเนา — ไม่แตะ
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
    await renderView();
    updateEditCount();
    setMsg($('#wordMsg'),
      `แทนที่แล้ว ${count} จุด` +
      (cross ? ` · มีอีก ${cross} ย่อหน้าที่คำนี้ถูกแบ่งคร่อมรูปแบบ ต้องคลิกแก้เองในหน้าเอกสาร` : ''),
      cross ? 'err' : 'ok');
  }

  // ---------- wiring ----------
  wireDropzone($('#dropDocx'), $('#fileDocx'), files => {
    const f = files.find(f => /\.docx$/i.test(f.name));
    if (f) loadDocx(f);
    else setMsg($('#wordOpenMsg'), 'รองรับเฉพาะไฟล์ .docx', 'err');
  });
  $('#pickDocx').addEventListener('click', () => $('#fileDocx').click());
  $('#btnReplace').addEventListener('click', replaceAll);
  $('#btnRepage').addEventListener('click', async () => {
    if (!doc) return;
    await renderView();
    setMsg($('#wordMsg'), 'จัดหน้าใหม่แล้ว', 'ok');
  });
  $('#btnDownload').addEventListener('click', async () => {
    if (!doc) return;
    downloadBlob(await buildDocx(), doc.fileName);
    setMsg($('#wordMsg'), 'ดาวน์โหลดแล้ว — format เดิม 100% ส่งต่อได้เลย', 'ok');
  });
  $('#btnCloseDoc').addEventListener('click', () => {
    doc = null;
    $('#wordCanvas').innerHTML = '';
    $('#wordWork').classList.add('hidden');
    $('#wordStart').classList.remove('hidden');
  });

  // ---------- dev hooks ----------
  window.__doc = {
    get doc() { return doc; },
    buildDocx, loadDocx,
    async b64() {
      const buf = new Uint8Array(await (await buildDocx()).arrayBuffer());
      let s = '';
      for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
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
