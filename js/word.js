// Word format-safe editor — WYSIWYG-ish: แก้ข้อความในหน้าเอกสารจริง (จัดหน้า/ตาราง/รูป/ฟอนต์)
// หลักการ: แก้เฉพาะข้อความใน <w:t> ของไฟล์เดิม ไม่มีการแปลงไฟล์ → format คงเดิม 100%
'use strict';
(() => {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const XML_NS = 'http://www.w3.org/XML/1998/namespace';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const IMG_OK = /^(png|jpg|jpeg|gif|bmp|webp|svg)$/;

  let doc = null; // { zip, fileName, parts:[{path,xmlDecl,dom,tNodes,orig,dirty}], images:{path:{rid:url}} }

  const partLabel = p =>
    p === 'word/document.xml' ? 'เนื้อหาเอกสาร' :
    /header/.test(p) ? 'หัวกระดาษ' :
    /footer/.test(p) ? 'ท้ายกระดาษ' :
    /footnotes/.test(p) ? 'เชิงอรรถ' :
    /endnotes/.test(p) ? 'อ้างอิงท้ายเรื่อง' : p;

  // ---------- small xml helpers ----------
  const kids = (el, ln) => [...el.children].filter(c => c.namespaceURI === W_NS && c.localName === ln);
  const kid = (el, ln) => kids(el, ln)[0] || null;
  const wval = el => el ? (el.getAttributeNS(W_NS, 'val') || el.getAttribute('w:val')) : null;
  const wattr = (el, n) => el ? (el.getAttributeNS(W_NS, n) || el.getAttribute('w:' + n)) : null;

  // ---------- run merge (rendering-neutral) ----------
  const serializeRPr = r => {
    const rPr = kid(r, 'rPr');
    return rPr ? new XMLSerializer().serializeToString(rPr) : '';
  };
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
  function mergeAdjacentRuns(dom) {
    const parents = new Set([...dom.getElementsByTagNameNS(W_NS, 'r')].map(r => r.parentNode));
    for (const parent of parents) {
      let prev = null;
      for (const child of [...parent.children]) {
        if (child.namespaceURI === W_NS && child.localName === 'r' && isSimpleTextRun(child)) {
          if (prev && serializeRPr(prev) === serializeRPr(child)) {
            kid(prev, 't').textContent += kid(child, 't').textContent;
            child.remove();
            continue;
          }
          prev = child;
        } else prev = null;
      }
    }
  }

  // ---------- images (resolve embedded pictures per part) ----------
  async function loadImages(zip, path, dom) {
    const map = {};
    const dir = path.replace(/[^/]+$/, '');
    const relsPath = dir + '_rels/' + path.slice(dir.length) + '.rels';
    const relsFile = zip.file(relsPath);
    if (!relsFile) return map;
    const rdom = new DOMParser().parseFromString(await relsFile.async('string'), 'application/xml');
    const rels = {};
    for (const rel of rdom.getElementsByTagName('Relationship'))
      rels[rel.getAttribute('Id')] = rel.getAttribute('Target');
    for (const blip of dom.getElementsByTagNameNS(A_NS, 'blip')) {
      const rid = blip.getAttributeNS(R_NS, 'embed') || blip.getAttribute('r:embed');
      if (!rid || rid in map) continue;
      let target = rels[rid];
      if (!target) continue;
      const full = (target.startsWith('/') ? target.slice(1) : dir + target).replace(/\/\.\//g, '/');
      const ext = (full.match(/\.(\w+)$/) || [, ''])[1].toLowerCase();
      const f = zip.file(full);
      if (!f || !IMG_OK.test(ext)) { map[rid] = null; continue; } // emf/wmf ฯลฯ เบราว์เซอร์วาดไม่ได้
      map[rid] = `data:image/${ext === 'jpg' ? 'jpeg' : ext};base64,` + await f.async('base64');
    }
    return map;
  }

  // ---------- load ----------
  async function loadDocx(file) {
    try {
      const zip = await JSZip.loadAsync(file);
      if (!zip.file('word/document.xml'))
        throw new Error('ไฟล์นี้ไม่ใช่ .docx (ถ้าเป็น .doc รุ่นเก่า ให้เปิดใน Word แล้ว Save As เป็น .docx ก่อน)');
      const names = Object.keys(zip.files)
        .filter(p => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(p));
      // ลำดับแสดงผลให้เหมือนเอกสาร: หัวกระดาษ → เนื้อหา → ท้ายกระดาษ → อื่น ๆ
      const rank = p => p.includes('header') ? 0 : p.includes('document') ? 1 : p.includes('footer') ? 2 : 3;
      names.sort((a, b) => rank(a) - rank(b) || a.localeCompare(b));
      const parts = [], images = {};
      for (const path of names) {
        const text = await zip.file(path).async('string');
        const dom = new DOMParser().parseFromString(text, 'application/xml');
        if (dom.getElementsByTagName('parsererror').length)
          throw new Error('อ่านโครงสร้างไฟล์ไม่ได้ (' + path + ')');
        const xmlDecl = (text.match(/^<\?xml[^>]*\?>/) || [''])[0];
        mergeAdjacentRuns(dom);
        const tNodes = [...dom.getElementsByTagNameNS(W_NS, 't')];
        parts.push({ path, xmlDecl, dom, tNodes, orig: tNodes.map(t => t.textContent), dirty: false });
        images[path] = await loadImages(zip, path, dom);
      }
      doc = { zip, fileName: file.name, parts, images };
      $('#wordFileName').textContent = file.name;
      $('#wordStart').classList.add('hidden');
      $('#wordWork').classList.remove('hidden');
      $('#exactPane').classList.add('hidden');
      $('#btnExact').textContent = '🔍 ตัวอย่างเป๊ะ';
      setMsg($('#wordMsg'), '');
      setMsg($('#wordOpenMsg'), '');
      renderDocView();
      updateEditCount();
    } catch (err) {
      setMsg($('#wordOpenMsg'), 'เปิดไฟล์ไม่สำเร็จ: ' + err.message, 'err');
    }
  }

  // ---------- run + paragraph styling ----------
  function runStyle(t) {
    const css = {};
    let r = t.parentNode;
    while (r && !(r.namespaceURI === W_NS && r.localName === 'r')) r = r.parentNode;
    const rPr = r && kid(r, 'rPr');
    if (!rPr) return css;
    const g = ln => kid(rPr, ln);
    if (g('b') && wval(g('b')) !== '0' && wval(g('b')) !== 'false') css.fontWeight = '700';
    if (g('i') && wval(g('i')) !== '0' && wval(g('i')) !== 'false') css.fontStyle = 'italic';
    if (g('u') && wval(g('u')) && wval(g('u')) !== 'none') css.textDecoration = 'underline';
    const sz = wval(g('sz'));
    if (sz) css.fontSize = Math.max(9, Math.min(56, Math.round(parseInt(sz, 10) / 2 * 1.34))) + 'px';
    const c = wval(g('color'));
    if (c && c !== 'auto' && /^[0-9A-Fa-f]{6}$/.test(c)) css.color = '#' + c;
    const hl = wval(g('highlight'));
    if (hl && hl !== 'none') css.background = hl;
    return css;
  }

  function paraStyle(p) {
    const s = {}, pPr = kid(p, 'pPr');
    if (!pPr) return s;
    const jc = wval(kid(pPr, 'jc'));
    if (jc) s.textAlign = ({ center: 'center', right: 'right', end: 'right', both: 'justify',
      distribute: 'justify', thaiDistribute: 'justify', left: 'left', start: 'left' })[jc] || 'left';
    const ind = kid(pPr, 'ind');
    if (ind) {
      const fl = wattr(ind, 'firstLine'), hg = wattr(ind, 'hanging'),
        left = wattr(ind, 'left') || wattr(ind, 'start');
      if (left) s.paddingLeft = (parseInt(left) / 15) + 'px';
      if (fl) s.textIndent = (parseInt(fl) / 15) + 'px';
      if (hg) s.textIndent = '-' + (parseInt(hg) / 15) + 'px';
    }
    const sp = kid(pPr, 'spacing');
    if (sp) {
      const bef = wattr(sp, 'before'), aft = wattr(sp, 'after'),
        line = wattr(sp, 'line'), rule = wattr(sp, 'lineRule');
      if (bef) s.marginTop = (parseInt(bef) / 15) + 'px';
      if (aft) s.marginBottom = (parseInt(aft) / 15) + 'px';
      if (line && (!rule || rule === 'auto')) s.lineHeight = (parseInt(line) / 240).toFixed(2);
    }
    return s;
  }

  // ---------- editable text span (maps to a w:t node) ----------
  function makeSpan(part, idx) {
    const node = part.tNodes[idx];
    const span = document.createElement('span');
    span.className = 'wt';
    span.contentEditable = 'true';
    span.spellcheck = false;
    span.textContent = node.textContent;
    Object.assign(span.style, runStyle(node));
    span.classList.toggle('edited', node.textContent !== part.orig[idx]);
    span.addEventListener('keydown', e => {
      if (e.key === 'Enter') e.preventDefault();
      if ((e.metaKey || e.ctrlKey) && ['b', 'i', 'u'].includes(e.key.toLowerCase())) e.preventDefault();
    });
    span.addEventListener('input', () => {
      const txt = span.textContent.replace(/[\r\n]+/g, ' ');
      node.textContent = txt;
      part.dirty = true;
      span.classList.toggle('edited', txt !== part.orig[idx]);
      updateEditCount();
    });
    span.addEventListener('blur', () => { if (span.children.length) span.textContent = node.textContent; });
    return span;
  }

  // ---------- render document tree into editable HTML ----------
  function appendImage(div, ctx, rid) {
    const src = rid ? ctx.images[rid] : null;
    if (src) { const img = document.createElement('img'); img.className = 'wimg'; img.src = src; div.appendChild(img); }
    else { const s = document.createElement('span'); s.className = 'wimgph'; s.textContent = '🖼'; div.appendChild(s); }
  }

  function walkInline(node, div, ctx) {
    for (const ch of node.childNodes) {
      if (ch.nodeType !== 1) continue;
      if (ch.namespaceURI === W_NS) {
        const ln = ch.localName;
        if (ln === 't') { const i = ctx.idxMap.get(ch); if (i !== undefined) div.appendChild(makeSpan(ctx.part, i)); }
        else if (ln === 'tab') { const s = document.createElement('span'); s.className = 'wtab'; div.appendChild(s); }
        else if (ln === 'br' || ln === 'cr') div.appendChild(document.createElement('br'));
        else if (ln === 'p') { /* nested para handled by caller */ }
        else if (ln === 'txbxContent') { const box = document.createElement('div'); box.className = 'wtxbox'; renderBlocks(box, ch, ctx); div.appendChild(box); }
        else walkInline(ch, div, ctx);
      } else {
        if (ch.localName === 'blip') appendImage(div, ctx, ch.getAttributeNS(R_NS, 'embed') || ch.getAttribute('r:embed'));
        else walkInline(ch, div, ctx);
      }
    }
  }

  function renderPara(p, ctx) {
    const div = document.createElement('div');
    div.className = 'para';
    Object.assign(div.style, paraStyle(p));
    walkInline(p, div, ctx);
    if (!div.childNodes.length) div.innerHTML = '<br>'; // ย่อหน้าว่าง = เว้นบรรทัด (คงจังหวะหน้า)
    return div;
  }

  function renderTable(tbl, ctx) {
    const table = document.createElement('table');
    table.className = 'doctable';
    for (const tr of kids(tbl, 'tr')) {
      const row = table.insertRow();
      for (const tc of kids(tr, 'tc')) {
        const td = row.insertCell();
        const span = kid(kid(tc, 'tcPr') || tc, 'gridSpan');
        if (span && wval(span)) td.colSpan = parseInt(wval(span));
        renderBlocks(td, tc, ctx);
      }
    }
    return table;
  }

  function renderBlocks(container, el, ctx) {
    for (const ch of el.children) {
      if (ch.namespaceURI !== W_NS) continue;
      if (ch.localName === 'p') container.appendChild(renderPara(ch, ctx));
      else if (ch.localName === 'tbl') container.appendChild(renderTable(ch, ctx));
      else if (ch.localName === 'sdt') { const c = kid(ch, 'sdtContent'); if (c) renderBlocks(container, c, ctx); }
    }
  }

  function renderDocView() {
    const ed = $('#wordEditor');
    ed.innerHTML = '';
    let any = false;
    for (const part of doc.parts) {
      const isMain = part.path === 'word/document.xml';
      const root = isMain ? part.dom.getElementsByTagNameNS(W_NS, 'body')[0] : part.dom.documentElement;
      if (!root) continue;
      const ctx = { part, idxMap: new Map(part.tNodes.map((n, i) => [n, i])), images: doc.images[part.path] || {} };
      if (!isMain) {
        const lbl = document.createElement('div');
        lbl.className = 'partlabel';
        lbl.textContent = partLabel(part.path);
        ed.appendChild(lbl);
      }
      const sec = document.createElement('div');
      sec.className = 'docsection ' + (isMain ? 'main' : 'aux');
      renderBlocks(sec, root, ctx);
      ed.appendChild(sec);
      any = any || sec.childNodes.length > 0;
    }
    if (!any) ed.innerHTML = '<div class="msg">ไม่พบข้อความที่แก้ได้ในไฟล์นี้</div>';
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
  function replaceAll() {
    if (!doc) return;
    const find = $('#findText').value, repl = $('#replText').value;
    if (!find) { setMsg($('#wordMsg'), 'พิมพ์คำที่จะค้นหาก่อน', 'err'); return; }
    let count = 0, cross = 0;
    for (const part of doc.parts) {
      for (const t of part.tNodes) {
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
    renderDocView();
    updateEditCount();
    setMsg($('#wordMsg'),
      `แทนที่แล้ว ${count} จุด` +
      (cross ? ` · มีอีก ${cross} ย่อหน้าที่คำนี้ถูกแบ่งคร่อมรูปแบบ ต้องคลิกแก้เองในหน้าเอกสาร` : ''),
      cross ? 'err' : 'ok');
  }

  // ---------- build ----------
  async function buildDocx() {
    for (const part of doc.parts) {
      if (!part.dirty) continue;
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

  async function renderPreview(blob) {
    const cont = $('#wordPreview');
    cont.innerHTML = '<div class="msg">กำลังสร้างตัวอย่าง...</div>';
    try { await window.docx.renderAsync(blob, cont, null, { inWrapper: true }); }
    catch (e) { cont.innerHTML = '<div class="msg err">แสดงตัวอย่างไม่ได้ (ไม่กระทบไฟล์จริง): ' + e.message + '</div>'; }
  }

  // ---------- wiring ----------
  wireDropzone($('#dropDocx'), $('#fileDocx'), files => {
    const f = files.find(f => /\.docx$/i.test(f.name));
    if (f) loadDocx(f);
    else setMsg($('#wordOpenMsg'), 'รองรับเฉพาะไฟล์ .docx', 'err');
  });
  $('#pickDocx').addEventListener('click', () => $('#fileDocx').click());
  $('#btnReplace').addEventListener('click', replaceAll);
  $('#btnExact').addEventListener('click', async () => {
    if (!doc) return;
    const pane = $('#exactPane');
    if (pane.classList.contains('hidden')) {
      pane.classList.remove('hidden');
      $('#btnExact').textContent = '🔍 ซ่อนตัวอย่างเป๊ะ';
      renderPreview(await buildDocx());
    } else {
      pane.classList.add('hidden');
      $('#btnExact').textContent = '🔍 ตัวอย่างเป๊ะ';
    }
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
    $('#exactPane').classList.add('hidden');
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
