// Tab 2: Google Sheet / Google Docs link -> real file download
// ใช้ session Google ในเบราว์เซอร์ของผู้ใช้เอง จึงโหลดได้แม้ไฟล์ไม่ได้แชร์สาธารณะ
'use strict';

// รับได้ทั้งลิงก์ Sheet และ Docs → คืนชนิด + id
function extractGoogleFile(link) {
  const s = (link || '').trim();
  let m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return { kind: 'spreadsheets', id: m[1] };
  m = s.match(/\/document\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return { kind: 'document', id: m[1] };
  m = s.match(/\/presentation\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return { kind: 'presentation', id: m[1] };
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return { kind: 'spreadsheets', id: s }; // id ดิบ = เดาว่า Sheet
  return null;
}

// ปุ่ม "ไฟล์ต้นฉบับ" = .xlsx สำหรับ Sheet, .docx สำหรับ Docs, .pptx สำหรับสไลด์
const NATIVE = { spreadsheets: { fmt: 'xlsx', label: 'Excel (.xlsx)' },
                 document: { fmt: 'docx', label: 'Word (.docx)' },
                 presentation: { fmt: 'pptx', label: 'PowerPoint (.pptx)' } };
const KIND_TH = { spreadsheets: 'Google Sheet', document: 'Google Docs', presentation: 'Google Slides' };

function updateBoqUI() {
  const f = extractGoogleFile($('#boqLink').value);
  const nativeBtn = $('#boqNative');
  if (f && NATIVE[f.kind]) {
    nativeBtn.textContent = '⬇️ โหลดเป็น ' + NATIVE[f.kind].label;
    nativeBtn.dataset.kind = f.kind;
    setMsg($('#boqMsg'), 'พบ ' + KIND_TH[f.kind] + ' — เลือกฟอร์แมตที่จะโหลด', '');
  } else {
    nativeBtn.textContent = '⬇️ โหลดไฟล์ต้นฉบับ';
  }
}

function boqDownload(format) {
  const msg = $('#boqMsg');
  const f = extractGoogleFile($('#boqLink').value);
  if (!f) {
    setMsg(msg, 'ลิงก์ไม่ถูกต้อง — วางลิงก์ Google Sheet, Google Docs หรือ Google Slides', 'err');
    return;
  }
  const fmt = format === 'native' ? NATIVE[f.kind].fmt : format;
  window.open(`https://docs.google.com/${f.kind}/d/${f.id}/export?format=${fmt}`, '_blank');
  setMsg(msg, fmt === 'pdf'
    ? 'กำลังดาวน์โหลด PDF — หน้าตาจะเหมือนที่เห็นใน ' + KIND_TH[f.kind] + ' เป๊ะ'
    : 'กำลังดาวน์โหลดไฟล์ต้นฉบับ (.' + fmt + ')', 'ok');
}

$('#boqNative').addEventListener('click', () => boqDownload('native'));
$('#boqPdf').addEventListener('click', () => boqDownload('pdf'));
$('#boqLink').addEventListener('input', updateBoqUI);
$('#boqLink').addEventListener('keydown', e => { if (e.key === 'Enter') boqDownload('pdf'); });
