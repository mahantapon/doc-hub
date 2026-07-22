// Tab 2: BOQ — Google Sheet link -> real file download (uses the browser's own Google session)
'use strict';

function extractSheetId(link) {
  const s = (link || '').trim();
  const m = s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s; // raw file id pasted directly
  return null;
}

function boqDownload(format) {
  const msg = $('#boqMsg');
  const id = extractSheetId($('#boqLink').value);
  if (!id) {
    setMsg(msg, 'ลิงก์ไม่ถูกต้อง — ต้องเป็นลิงก์ Google Sheet (มี /spreadsheets/d/... อยู่ในลิงก์)', 'err');
    return;
  }
  const url = `https://docs.google.com/spreadsheets/d/${id}/export?format=${format}`;
  window.open(url, '_blank');
  setMsg(msg, format === 'xlsx'
    ? 'กำลังดาวน์โหลดเป็น .xlsx — ได้ไฟล์แล้วส่งไลน์ให้ช่างได้เลย'
    : 'กำลังดาวน์โหลดเป็น PDF', 'ok');
}

$('#boqXlsx').addEventListener('click', () => boqDownload('xlsx'));
$('#boqPdf').addEventListener('click', () => boqDownload('pdf'));
$('#boqLink').addEventListener('keydown', e => { if (e.key === 'Enter') boqDownload('xlsx'); });
