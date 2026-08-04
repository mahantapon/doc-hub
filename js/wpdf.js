// Tab: Word → PDF (ฟอนต์ไม่เพี้ยน)
// เรนเดอร์ .docx ด้วย docx-preview (ซึ่งโหลด "ฟอนต์ที่ฝังในไฟล์" มาใช้จริง)
// แล้วสร้าง PDF ด้วยเครื่องพิมพ์ของเบราว์เซอร์ (เวกเตอร์ ฟอนต์คมชัด เลือกข้อความได้)
// มีตัวเลือกบังคับแปลงฟอนต์เป็น Sarabun ก่อน = กันเพี้ยน 100% แม้ไฟล์ไม่ได้ฝังฟอนต์
'use strict';
(() => {
  let lastFile = null;

  async function prepare(file) {
    lastFile = file;
    const cont = $('#wpdfRender');
    setMsg($('#wpdfMsg'), 'กำลังเตรียมเอกสาร...', '');
    $('#wpdfPrint').disabled = true;
    cont.innerHTML = '';
    try {
      let blob = file;
      if ($('#wpdfSarabun').checked && window.DocHubFont) {
        const r = await window.DocHubFont.convertOne(file, 'Sarabun', false);
        blob = r.blob;
      }
      await window.docx.renderAsync(blob, cont, null, { inWrapper: true });
      await (document.fonts ? document.fonts.ready.catch(() => {}) : null);
      $('#wpdfName').textContent = file.name;
      $('#wpdfPrint').disabled = false;
      setMsg($('#wpdfMsg'), 'พร้อมแล้ว ✓ ตรวจหน้าตาด้านล่าง แล้วกด "สร้าง PDF" → ในหน้าต่างพิมพ์เลือกปลายทางเป็น "บันทึกเป็น PDF"', 'ok');
    } catch (e) {
      setMsg($('#wpdfMsg'), 'อ่านไฟล์ไม่สำเร็จ: ' + e.message, 'err');
    }
  }

  async function toPdf() {
    if (!$('#wpdfRender').querySelector('.docx-wrapper')) return;
    document.documentElement.classList.add('printing-wpdf');
    try { await document.fonts.ready; } catch (_) {}
    setTimeout(() => window.print(), 80);
  }
  window.addEventListener('afterprint', () => document.documentElement.classList.remove('printing-wpdf'));

  wireDropzone($('#dropWpdf'), $('#fileWpdf'), files => {
    const f = files.find(f => /\.docx$/i.test(f.name));
    if (f) prepare(f);
    else setMsg($('#wpdfMsg'), 'รองรับเฉพาะไฟล์ .docx', 'err');
  });
  $('#pickWpdf').addEventListener('click', () => $('#fileWpdf').click());
  $('#wpdfSarabun').addEventListener('change', () => { if (lastFile) prepare(lastFile); });
  $('#wpdfPrint').addEventListener('click', toPdf);
})();
