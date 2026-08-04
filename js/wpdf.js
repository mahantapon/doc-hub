// Tab: Word → PDF (ฟอนต์ไม่เพี้ยน)
// เรนเดอร์ .docx ด้วย docx-preview แล้ว:
//  - ปุ่มหลัก "ดาวน์โหลด PDF" = html2canvas + jsPDF (ได้ไฟล์ลงเครื่องตรง ๆ ทีละหน้า)
//  - ปุ่มรอง "พิมพ์" = window.print() (เวกเตอร์ เลือกข้อความได้ แต่ต้องกด Save as PDF เอง)
// ค่าเริ่มต้น "บังคับ Sarabun" = เปิด → ไม่เพี้ยนแม้ไฟล์ไม่ได้ฝังฟอนต์
'use strict';
(() => {
  let lastFile = null, baseName = 'document';

  async function prepare(file) {
    lastFile = file;
    baseName = file.name.replace(/\.docx$/i, '');
    const cont = $('#wpdfRender');
    setMsg($('#wpdfMsg'), 'กำลังเตรียมเอกสาร...', '');
    $('#wpdfDownload').disabled = true;
    $('#wpdfPrint').disabled = true;
    cont.innerHTML = '';
    try {
      let blob = file;
      if ($('#wpdfSarabun').checked && window.DocHubFont) {
        const r = await window.DocHubFont.convertOne(file, 'Sarabun', false);
        blob = r.blob;
      }
      // ignoreLastRenderedPageBreak:false = แบ่งหน้าตามจุดแบ่งหน้าจริงของ Word (ได้หลายหน้า ไม่ใช่หน้ายาวหน้าเดียว)
      await window.docx.renderAsync(blob, cont, null, { inWrapper: true, ignoreLastRenderedPageBreak: false });
      try { await document.fonts.ready; } catch (_) {}
      $('#wpdfName').textContent = file.name;
      $('#wpdfDownload').disabled = false;
      $('#wpdfPrint').disabled = false;
      setMsg($('#wpdfMsg'), 'พร้อมแล้ว ✓ ตรวจหน้าตาด้านล่าง แล้วกด "ดาวน์โหลด PDF"', 'ok');
    } catch (e) {
      setMsg($('#wpdfMsg'), 'อ่านไฟล์ไม่สำเร็จ: ' + e.message, 'err');
    }
  }

  // ปุ่มหลัก: สร้างไฟล์ PDF แล้วดาวน์โหลดลงเครื่องเลย (แปลงหน้าเอกสารเป็นภาพทีละหน้า)
  async function downloadPdf() {
    const pages = [...$('#wpdfRender').querySelectorAll('.docx-wrapper > section')];
    if (!pages.length) return;
    if (!window.jspdf || !window.html2canvas) {
      setMsg($('#wpdfMsg'), 'ไลบรารีสร้าง PDF โหลดไม่ครบ — ลองรีเฟรชหน้าใหม่', 'err'); return;
    }
    $('#wpdfDownload').disabled = true;
    try { await document.fonts.ready; } catch (_) {}
    const { jsPDF } = window.jspdf;
    let pdf = null;
    const PX2PT = 0.75; // 96px = 72pt
    for (let i = 0; i < pages.length; i++) {
      setMsg($('#wpdfMsg'), `กำลังสร้าง PDF... หน้า ${i + 1}/${pages.length}`, '');
      const sec = pages[i];
      const canvas = await window.html2canvas(sec, {
        scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false,
        windowWidth: sec.scrollWidth, width: sec.scrollWidth, height: sec.scrollHeight
      });
      const wpt = sec.offsetWidth * PX2PT, hpt = sec.offsetHeight * PX2PT;
      const orient = wpt > hpt ? 'landscape' : 'portrait';
      if (!pdf) pdf = new jsPDF({ unit: 'pt', format: [wpt, hpt], orientation: orient });
      else pdf.addPage([wpt, hpt], orient);
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, wpt, hpt);
    }
    pdf.save(baseName + '.pdf');
    $('#wpdfDownload').disabled = false;
    setMsg($('#wpdfMsg'), 'ดาวน์โหลด PDF แล้ว ✓ (' + pages.length + ' หน้า)', 'ok');
  }

  // ปุ่มรอง: พิมพ์ผ่านเบราว์เซอร์ (เวกเตอร์ เลือกข้อความได้)
  async function toPrint() {
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
  $('#wpdfDownload').addEventListener('click', downloadPdf);
  $('#wpdfPrint').addEventListener('click', toPrint);
})();
