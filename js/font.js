// Tab: เปลี่ยนฟอนต์ให้เข้ากับ Google Docs
// เปลี่ยนเฉพาะ "ชื่อฟอนต์" ในไฟล์ .docx เดิม (w:rFonts) → ฟอนต์ที่ Google Docs มีจริง
// ส่วนอื่นของไฟล์ (ข้อความ ตาราง รูป กล่อง หัว-ท้ายกระดาษ) ไม่ถูกแตะเลย
'use strict';
(() => {
  // ฟอนต์ไทยที่ Google Docs "ไม่มี" → ต้องเปลี่ยน
  const THAI_FONTS = [
    'TH SarabunPSK', 'TH Sarabun New', 'TH SarabunIT๙', 'TH Sarabun IT๙',
    'Cordia New', 'CordiaUPC', 'Cordia UPC',
    'Angsana New', 'AngsanaUPC', 'Angsana UPC',
    'Browallia New', 'BrowalliaUPC', 'DilleniaUPC', 'EucrosiaUPC', 'FreesiaUPC',
    'IrisUPC', 'JasmineUPC', 'KodchiangUPC', 'LilyUPC',
    'Leelawadee', 'Leelawadee UI', 'TH Niramit AS', 'TH Krub', 'TH Charmonman',
    'Tahoma', 'Microsoft Sans Serif'
  ];
  // ฟอนต์ละติน (เปลี่ยนก็ต่อเมื่อผู้ใช้ติ๊กเลือก)
  const LATIN_FONTS = [
    'Calibri', 'Calibri Light', 'Times New Roman', 'Arial', 'Segoe UI',
    'Courier New', 'Cambria', 'Verdana', 'Helvetica'
  ];
  // ฟอนต์ปลายทาง = ต้องมีใน Google Docs จริง
  const TARGETS = ['Sarabun', 'Noto Sans Thai', 'Niramit', 'Kanit', 'Prompt', 'Mitr', 'Maitree'];

  const FONT_PARTS = /^word\/(document|styles|header\d*|footer\d*|footnotes|endnotes|numbering)\.xml$/;
  const FONT_ATTR = /(w:(?:ascii|hAnsi|cs|eastAsia)=")([^"]*)(")/g;

  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  function fillTargets() {
    $('#fontTarget').innerHTML = TARGETS.map(t =>
      `<option value="${t}">${t}${t === 'Sarabun' ? ' (แนะนำ — ใกล้เคียง TH SarabunPSK)' : ''}</option>`).join('');
  }

  // อ่านรายชื่อฟอนต์ที่ไฟล์ใช้อยู่
  async function scanFonts(zip) {
    const found = {};
    for (const path of Object.keys(zip.files)) {
      if (!FONT_PARTS.test(path)) continue;
      const xml = await zip.file(path).async('string');
      let m;
      const re = new RegExp(FONT_ATTR.source, 'g');
      while ((m = re.exec(xml))) { if (m[2]) found[m[2]] = (found[m[2]] || 0) + 1; }
    }
    return found;
  }

  async function convertOne(file, target, alsoLatin) {
    const zip = await JSZip.loadAsync(file);
    if (!zip.file('word/document.xml')) throw new Error('ไม่ใช่ไฟล์ .docx');
    const before = await scanFonts(zip);
    const mapSet = new Set([...THAI_FONTS, ...(alsoLatin ? LATIN_FONTS : [])]);
    mapSet.delete(target); // ถ้าใช้ฟอนต์ปลายทางอยู่แล้ว ไม่ต้องแตะ

    let changed = 0;
    const changedFonts = new Set();
    for (const path of Object.keys(zip.files)) {
      if (!FONT_PARTS.test(path)) continue;
      const xml = await zip.file(path).async('string');
      const out = xml.replace(new RegExp(FONT_ATTR.source, 'g'), (m, p1, name, p3) => {
        if (!mapSet.has(name)) return m;
        changed++; changedFonts.add(name);
        return p1 + target + p3;
      });
      if (out !== xml) zip.file(path, out);
    }

    // ประกาศฟอนต์ปลายทางใน fontTable ให้ Word รู้จัก
    const ftPath = 'word/fontTable.xml';
    if (zip.file(ftPath)) {
      let ft = await zip.file(ftPath).async('string');
      if (!new RegExp(`<w:font w:name="${target}"`).test(ft)) {
        ft = ft.replace(/<\/w:fonts>\s*$/,
          `<w:font w:name="${target}"><w:charset w:val="00"/><w:family w:val="swiss"/>` +
          `<w:pitch w:val="variable"/></w:font></w:fonts>`);
        zip.file(ftPath, ft);
      }
    }

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    });
    return { blob, changed, changedFonts: [...changedFonts], before };
  }

  async function handleFiles(files) {
    files = files.filter(f => /\.docx$/i.test(f.name));
    const cont = $('#fontResults');
    if (!files.length) {
      cont.insertAdjacentHTML('afterbegin',
        '<div class="row"><span class="st">❌</span><span class="nm">รองรับเฉพาะไฟล์ .docx</span></div>');
      return;
    }
    const target = $('#fontTarget').value;
    const alsoLatin = $('#alsoLatin').checked;
    const out = [];
    for (const f of files) {
      const row = document.createElement('div');
      row.className = 'row';
      row.innerHTML = `<span class="st">⏳</span><span class="nm">${esc(f.name)}</span><span class="ex"></span>`;
      cont.prepend(row);
      try {
        const r = await convertOne(f, target, alsoLatin);
        out.push({ name: f.name.replace(/\.docx$/i, ` (${target}).docx`), blob: r.blob });
        row.querySelector('.st').textContent = r.changed ? '✅' : 'ℹ️';
        row.querySelector('.ex').innerHTML = r.changed
          ? `<span style="color:var(--green)">เปลี่ยน ${r.changed} จุด: ${esc(r.changedFonts.join(', '))} → ${esc(target)}</span>`
          : `<span class="err">ไม่พบฟอนต์ที่ต้องเปลี่ยน (ไฟล์ใช้: ${esc(Object.keys(r.before).join(', ') || '-')})</span>`;
      } catch (e) {
        row.querySelector('.st').textContent = '❌';
        row.querySelector('.ex').innerHTML = `<span class="err">${esc(e.message)}</span>`;
      }
    }
    if (!out.length) return;
    if (out.length === 1) downloadBlob(out[0].blob, out[0].name);
    else {
      const zip = new JSZip();
      out.forEach(o => zip.file(o.name, o.blob));
      downloadBlob(await zip.generateAsync({ type: 'blob' }), `docx-${target}.zip`);
    }
  }

  fillTargets();
  wireDropzone($('#dropFont'), $('#fileFont'), handleFiles);
  $('#pickFont').addEventListener('click', () => $('#fileFont').click());
})();
