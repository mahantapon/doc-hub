// Tab: Template Generator — fill a form, get a perfect .docx from a token template.
// The template .docx already contains {{TOKENS}} as single contiguous runs, so a
// plain string-replace in document.xml is 100% safe. Fonts/boxes/layout untouched.
'use strict';
(() => {
  const TEMPLATES = [{
    id: 'contract-renovate',
    name: 'สัญญาว่าจ้างผู้รับเหมา (ก่อสร้าง / ต่อเติม / รีโนเวท)',
    file: 'templates/contract-renovate.docx',
    outName: 'สัญญาว่าจ้าง',
    fields: [
      { token: 'CONTRACT_NO', label: 'เลขที่สัญญา', type: 'text', def: 'P2026-06-01' },
      { token: 'CONTRACT_DATE', label: 'วันที่ทำสัญญา', type: 'date' },
      { token: 'CONTRACTOR_NAME', label: 'ชื่อผู้รับจ้าง (บริษัท / บุคคล)', type: 'text', wide: true },
      { token: 'CONTRACTOR_ID', label: 'เลขนิติบุคคล / บัตรประชาชน ผู้รับจ้าง', type: 'text' },
      { token: 'ACCOUNT_NO', label: 'เลขที่บัญชีผู้รับจ้าง', type: 'text' },
      { token: 'CONTRACTOR_ADDR', label: 'ที่อยู่ผู้รับจ้าง', type: 'text', wide: true },
      { token: 'SITE_ADDR', label: 'สถานที่ก่อสร้าง', type: 'text', wide: true },
      { token: 'CONTRACT_VALUE', label: 'มูลค่างาน (บาท)', type: 'number', money: true },
    ],
    // derived tokens filled automatically
    derived: { CONTRACT_VALUE_TEXT: v => bahtText(v.CONTRACT_VALUE_RAW) }
  }];

  const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
    'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];

  function thaiDate(iso) {
    if (!iso) return '';
    const [y, m, d] = iso.split('-').map(Number);
    return `${d} เดือน ${TH_MONTHS[m - 1]} ${y + 543}`;
  }

  function formatMoney(raw) {
    const n = parseFloat(String(raw).replace(/,/g, ''));
    if (isNaN(n)) return '';
    return n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function bahtText(raw) {
    let number = parseFloat(String(raw).replace(/,/g, ''));
    if (isNaN(number)) return '';
    const t = ['ศูนย์', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
    const u = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
    function say(n) {
      n = String(n).replace(/^0+/, '') || '0';
      if (n === '0') return '';
      if (n.length > 6) return say(n.slice(0, n.length - 6)) + 'ล้าน' + say(n.slice(n.length - 6));
      let s = '';
      const len = n.length;
      for (let i = 0; i < len; i++) {
        const digit = +n[i], pos = len - i - 1;
        if (digit === 0) continue;
        if (pos === 0 && digit === 1 && len > 1) s += 'เอ็ด';
        else if (pos === 1 && digit === 1) s += 'สิบ';
        else if (pos === 1 && digit === 2) s += 'ยี่สิบ';
        else s += t[digit] + u[pos];
      }
      return s;
    }
    const neg = number < 0; number = Math.abs(number);
    const [ip, dp] = number.toFixed(2).split('.');
    let res = say(ip) ? say(ip) + 'บาท' : 'ศูนย์บาท';
    res += dp === '00' ? 'ถ้วน' : say(dp) + 'สตางค์';
    return (neg ? 'ลบ' : '') + res;
  }

  const xmlEsc = s => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/[\r\n]+/g, ' ');

  let cur = TEMPLATES[0];

  function renderForm() {
    const sel = $('#tplSelect');
    sel.innerHTML = TEMPLATES.map(t => `<option value="${t.id}">${t.name}</option>`).join('');
    const form = $('#tplForm');
    form.innerHTML = cur.fields.map(f =>
      `<label class="fld ${f.wide ? 'wide' : ''}">${f.label}
        <input data-token="${f.token}" type="${f.type === 'number' ? 'text' : f.type}"
               inputmode="${f.type === 'number' ? 'decimal' : 'text'}"
               value="${f.def || ''}" autocomplete="off">
      </label>`).join('');
  }

  function collectValues() {
    const v = {};
    $all('#tplForm input').forEach(inp => { v[inp.dataset.token] = inp.value.trim(); });
    // transforms
    if ('CONTRACT_DATE' in v) v.CONTRACT_DATE = thaiDate(v.CONTRACT_DATE);
    if ('CONTRACT_VALUE' in v) {
      v.CONTRACT_VALUE_RAW = v.CONTRACT_VALUE;
      v.CONTRACT_VALUE = formatMoney(v.CONTRACT_VALUE);
    }
    if (cur.derived) for (const [k, fn] of Object.entries(cur.derived)) v[k] = fn(v);
    return v;
  }

  async function buildBlob() {
    const values = collectValues();
    const buf = await (await fetch(cur.file)).arrayBuffer();
    const zip = await JSZip.loadAsync(buf);
    let xml = await zip.file('word/document.xml').async('string');
    // replace every known token (blank if field left empty)
    const tokens = new Set([...cur.fields.map(f => f.token), ...Object.keys(cur.derived || {}), 'CONTRACT_VALUE_TEXT']);
    for (const tok of tokens) xml = xml.split(`{{${tok}}}`).join(xmlEsc(values[tok]));
    // clean any leftover unknown tokens so none leak into the document
    xml = xml.replace(/\{\{[A-Z_]+\}\}/g, '');
    zip.file('word/document.xml', xml);
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    });
  }

  async function preview() {
    const msg = $('#tplMsg');
    setMsg(msg, 'กำลังสร้างตัวอย่าง...', '');
    try {
      const blob = await buildBlob();
      window.__tplBlob = blob;
      const cont = $('#tplPreview');
      cont.innerHTML = '';
      await window.docx.renderAsync(blob, cont, null, { inWrapper: true });
      setMsg(msg, 'พร้อมแล้ว — ตรวจดูตัวอย่างด้านล่าง แล้วกดดาวน์โหลด', 'ok');
    } catch (e) {
      setMsg(msg, 'สร้างไม่สำเร็จ: ' + e.message, 'err');
    }
  }

  async function download() {
    try {
      const blob = window.__tplBlob || await buildBlob();
      const no = ($('#tplForm input[data-token="CONTRACT_NO"]') || {}).value || '';
      downloadBlob(blob, `${cur.outName}${no ? ' ' + no : ''}.docx`);
      setMsg($('#tplMsg'), 'ดาวน์โหลดแล้ว — ฟอนต์/เลย์เอาต์ครบ 100% เปิดใน Word ส่งต่อได้เลย', 'ok');
    } catch (e) { setMsg($('#tplMsg'), 'ดาวน์โหลดไม่สำเร็จ: ' + e.message, 'err'); }
  }

  function printPdf() {
    const cont = $('#tplPreview');
    if (!cont.innerHTML.trim()) { setMsg($('#tplMsg'), 'กด"สร้าง/พรีวิว" ก่อน', 'err'); return; }
    const w = window.open('', '_blank');
    w.document.write('<!DOCTYPE html><html><head><meta charset="utf-8"><title>' +
      cur.outName + '</title></head><body>' + cont.innerHTML +
      '<script>onload=()=>{print()}<\/script></body></html>');
    w.document.close();
  }

  // wiring
  renderForm();
  $('#tplSelect').addEventListener('change', e => {
    cur = TEMPLATES.find(t => t.id === e.target.value) || TEMPLATES[0];
    renderForm();
    $('#tplPreview').innerHTML = '';
  });
  $('#tplPreviewBtn').addEventListener('click', preview);
  $('#tplDownload').addEventListener('click', download);
  $('#tplPdf').addEventListener('click', printPdf);
})();
