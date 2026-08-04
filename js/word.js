// Tab: แก้ Word — WYSIWYG จริงด้วย SuperDoc (แก้ในหน้าเอกสารเหมือนเปิด Word)
// เปิด .docx → คลิกแก้ตรงไหนก็ได้ พิมพ์/ลบ/จัดหน้า → ดาวน์โหลดกลับเป็น .docx
// ทดสอบกับสัญญาไทยจริง: ฟอนต์ฝัง, จัดคำแบบไทย, กล่องลอย, หัว-ท้ายกระดาษ คงครบหลัง export
'use strict';
console.log('[Doc Hub] Word editor build v9 (SuperDoc WYSIWYG)');
(() => {
  let sd = null, fileName = null;

  const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

  async function loadDocx(file) {
    try {
      if (!window.SuperDoc) throw new Error('ตัว editor ยังโหลดไม่เสร็จ — รอสักครู่แล้วลองใหม่');
      // SuperDoc ใช้ MIME type ในการรู้จักชนิดไฟล์ — บางทางมา (ลากจากบางแอป) type ว่าง ต้องเติมให้
      if (file.type !== DOCX_MIME) file = new File([file], file.name, { type: DOCX_MIME });
      fileName = file.name;
      $('#wordStart').classList.add('hidden');
      $('#wordWork').classList.remove('hidden');
      $('#wordFileName').textContent = file.name;
      setMsg($('#wordMsg'), 'กำลังเปิดเอกสาร...', '');
      if (sd) { try { sd.destroy(); } catch (_) {} sd = null; }
      $('#sdToolbar').innerHTML = '';
      $('#sdEditor').innerHTML = '';
      sd = new window.SuperDoc({
        selector: '#sdEditor',
        toolbar: '#sdToolbar',
        document: file,
        documentMode: 'editing',
        pagination: true,
        onReady: () => setMsg($('#wordMsg'), 'พร้อมแก้ไข ✓ คลิกในเอกสารแล้วพิมพ์/ลบได้เหมือนใช้ Word', 'ok'),
        onException: e => setMsg($('#wordMsg'), 'มีปัญหาในการแสดงผล: ' + (e && (e.message || e)), 'err')
      });
      setMsg($('#wordOpenMsg'), '');
    } catch (e) {
      $('#wordWork').classList.add('hidden');
      $('#wordStart').classList.remove('hidden');
      setMsg($('#wordOpenMsg'), 'เปิดไฟล์ไม่สำเร็จ: ' + e.message, 'err');
    }
  }

  async function download() {
    if (!sd) return;
    try {
      setMsg($('#wordMsg'), 'กำลังบันทึกไฟล์...', '');
      const blobs = await sd.exportEditorsToDOCX();
      if (!blobs || !blobs[0]) throw new Error('สร้างไฟล์ไม่สำเร็จ');
      downloadBlob(blobs[0], fileName || 'document.docx');
      setMsg($('#wordMsg'), 'ดาวน์โหลดแล้ว ✓ เปิดใน Word / ส่งต่อได้เลย', 'ok');
    } catch (e) {
      setMsg($('#wordMsg'), 'บันทึกไม่สำเร็จ: ' + e.message, 'err');
    }
  }

  function closeDoc() {
    if (sd) { try { sd.destroy(); } catch (_) {} sd = null; }
    $('#sdEditor').innerHTML = '';
    $('#sdToolbar').innerHTML = '';
    $('#wordWork').classList.add('hidden');
    $('#wordStart').classList.remove('hidden');
  }

  wireDropzone($('#dropDocx'), $('#fileDocx'), files => {
    const f = files.find(f => /\.docx$/i.test(f.name));
    if (f) loadDocx(f);
    else setMsg($('#wordOpenMsg'), 'รองรับเฉพาะไฟล์ .docx', 'err');
  });
  $('#pickDocx').addEventListener('click', () => $('#fileDocx').click());
  $('#btnDownload').addEventListener('click', download);
  $('#btnCloseDoc').addEventListener('click', closeDoc);

  // ---------- dev hooks (ทดสอบอัตโนมัติ) ----------
  window.__doc = {
    get sd() { return sd; },
    loadDocx,
    async b64() {
      const blobs = await sd.exportEditorsToDOCX();
      const buf = new Uint8Array(await blobs[0].arrayBuffer());
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
