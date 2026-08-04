// ล้างไฟล์ .docx ที่ได้จาก editor ให้ถูกมาตรฐาน OOXML ก่อนส่งให้ผู้ใช้
// แก้ 2 อย่างที่ทำให้ Microsoft Word ปฏิเสธไฟล์ (Word ตรวจเข้มกว่าโปรแกรมอื่น):
//  1) zip ต้องไม่มี entry ของโฟลเดอร์ (OPC ห้าม)
//  2) ลูกของ <w:pPr>/<w:rPr> ต้องเรียงตามลำดับ schema เป๊ะ
//     (เคสจริง: SuperDoc เขียน szCs ไว้ท้ายสุดหลัง lang → Word เปิดไม่ได้)
'use strict';
(() => {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  const PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl',
    'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
    'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
    'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection',
    'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];

  const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike',
    'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color',
    'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
    'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];

  // เรียงลูกให้ตรง schema — ตัวที่ไม่รู้จักคงตำแหน่งเดิมโดยผูกกับตัวก่อนหน้า
  function reorder(el, order) {
    const kids = [...el.children];
    if (kids.length < 2) return false;
    let lastKnown = -1;
    const keyed = kids.map((c, i) => {
      const name = c.namespaceURI === W_NS ? c.localName : null;
      const idx = name ? order.indexOf(name) : -1;
      if (idx >= 0) { lastKnown = idx; return { c, key: idx, i }; }
      return { c, key: lastKnown + 0.5, i };
    });
    const sorted = [...keyed].sort((a, b) => (a.key - b.key) || (a.i - b.i));
    if (sorted.every((s, i) => s.c === kids[i])) return false;
    for (const s of sorted) el.appendChild(s.c);
    return true;
  }

  function fixXml(xmlText) {
    const dom = new DOMParser().parseFromString(xmlText, 'application/xml');
    if (dom.getElementsByTagName('parsererror').length) return { text: xmlText, fixed: 0 };
    let fixed = 0;
    for (const [tag, order] of [['pPr', PPR_ORDER], ['rPr', RPR_ORDER]])
      for (const el of [...dom.getElementsByTagNameNS(W_NS, tag)])
        if (reorder(el, order)) fixed++;
    if (!fixed) return { text: xmlText, fixed: 0 };
    let out = new XMLSerializer().serializeToString(dom);
    const decl = (xmlText.match(/^<\?xml[^>]*\?>/) || [''])[0];
    if (decl && !out.startsWith('<?xml')) out = decl + '\r\n' + out;
    return { text: out, fixed };
  }

  // คืน Blob ใหม่ที่ Word เปิดได้ (ถ้าไฟล์ดีอยู่แล้วก็แค่สร้าง zip ใหม่ที่สะอาด)
  async function sanitizeDocx(blob) {
    const src = await JSZip.loadAsync(blob);
    const out = new JSZip();
    let fixedParts = 0, dropped = 0;
    const FIXABLE = /^word\/(document|header\d*|footer\d*|footnotes|endnotes|styles|numbering)\.xml$/;
    for (const path of Object.keys(src.files)) {
      const f = src.files[path];
      if (f.dir || path.endsWith('/')) { dropped++; continue; } // ตัด entry โฟลเดอร์ทิ้ง
      // createFolders:false = ห้าม JSZip สร้าง entry โฟลเดอร์ให้อัตโนมัติ (Word ไม่ยอมรับ)
      if (FIXABLE.test(path)) {
        const r = fixXml(await f.async('string'));
        if (r.fixed) fixedParts++;
        out.file(path, r.text, { createFolders: false });
      } else {
        out.file(path, await f.async('uint8array'), { createFolders: false });
      }
    }
    const fixedBlob = await out.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    });
    console.log(`[Doc Hub] docx sanitize: ตัด dir entries ${dropped} · เรียง schema ใหม่ ${fixedParts} ส่วน`);
    return fixedBlob;
  }

  // ---------- เตรียมไฟล์ก่อนส่งเข้า editor ----------
  // เอกสารไทยเก็บขนาดฟอนต์ 2 ค่า: w:sz (ละติน) กับ w:szCs (ไทย/complex script)
  // ตัว editor อ่านแค่ w:sz → ข้อความไทยเลยแสดงเล็ก และตอนบันทึกยังเขียนทับ szCs = sz
  // (ของจริงเจอ: ไทย 16pt ถูกหดเหลือ 12pt ในไฟล์)
  // แก้โดย: เฉพาะ run ที่เป็น "ไทยล้วน" ให้ตั้ง sz = ขนาดไทยที่ใช้จริง
  // ปลอดภัยเพราะ run ไทยล้วนไม่มีตัวอักษรละติน → w:sz ไม่มีผลต่อการแสดงผลใน Word
  const HAS_LATIN = /[A-Za-z0-9]/;
  const PARTS_TEXT = /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/;

  const kidVal = (el, ln) => {
    if (!el) return null;
    const k = [...el.children].find(c => c.namespaceURI === W_NS && c.localName === ln);
    return k ? (k.getAttributeNS(W_NS, 'val') || k.getAttribute('w:val')) : null;
  };

  // อ่านขนาดจาก styles.xml (ไล่ตาม basedOn) + ค่าเริ่มต้นเอกสาร
  function readStyles(stylesXml) {
    const map = {}; let defSzCs = null;
    if (!stylesXml) return { map, defSzCs };
    const dom = new DOMParser().parseFromString(stylesXml, 'application/xml');
    const dd = dom.getElementsByTagNameNS(W_NS, 'docDefaults')[0];
    if (dd) {
      const rd = dd.getElementsByTagNameNS(W_NS, 'rPrDefault')[0];
      const rpr = rd && rd.getElementsByTagNameNS(W_NS, 'rPr')[0];
      defSzCs = kidVal(rpr, 'szCs') || kidVal(rpr, 'sz');
    }
    for (const st of dom.getElementsByTagNameNS(W_NS, 'style')) {
      const id = st.getAttributeNS(W_NS, 'styleId') || st.getAttribute('w:styleId');
      if (!id) continue;
      const rpr = [...st.children].find(c => c.namespaceURI === W_NS && c.localName === 'rPr');
      map[id] = { szCs: kidVal(rpr, 'szCs'), sz: kidVal(rpr, 'sz'), basedOn: kidVal(st, 'basedOn') };
    }
    return { map, defSzCs };
  }

  function styleSzCs(map, id, depth = 0) {
    if (!id || !map[id] || depth > 8) return null;
    return map[id].szCs || map[id].sz || styleSzCs(map, map[id].basedOn, depth + 1);
  }

  function setSize(dom, rPr, val) {
    for (const ln of ['sz', 'szCs']) {
      let el = [...rPr.children].find(c => c.namespaceURI === W_NS && c.localName === ln);
      if (!el) { el = dom.createElementNS(W_NS, 'w:' + ln); rPr.appendChild(el); }
      el.setAttributeNS(W_NS, 'w:val', val);
    }
    reorder(rPr, RPR_ORDER);
  }

  async function prepareForEditor(blob) {
    const src = await JSZip.loadAsync(blob);
    const stylesXml = src.file('word/styles.xml') ? await src.file('word/styles.xml').async('string') : null;
    const { map, defSzCs } = readStyles(stylesXml);
    let touched = 0;
    for (const path of Object.keys(src.files)) {
      if (!PARTS_TEXT.test(path)) continue;
      const text = await src.file(path).async('string');
      const dom = new DOMParser().parseFromString(text, 'application/xml');
      if (dom.getElementsByTagName('parsererror').length) continue;
      let changed = false;
      for (const p of dom.getElementsByTagNameNS(W_NS, 'p')) {
        const pPr = [...p.children].find(c => c.namespaceURI === W_NS && c.localName === 'pPr');
        const pStyle = kidVal(pPr, 'pStyle');
        for (const r of p.getElementsByTagNameNS(W_NS, 'r')) {
          const txt = [...r.getElementsByTagNameNS(W_NS, 't')].map(t => t.textContent).join('');
          if (!txt.trim() || HAS_LATIN.test(txt)) continue;           // ข้ามถ้าไม่ใช่ไทยล้วน
          const rPr = [...r.children].find(c => c.namespaceURI === W_NS && c.localName === 'rPr');
          if (!rPr) continue;
          const sz = kidVal(rPr, 'sz'), szCs = kidVal(rPr, 'szCs');
          const thai = szCs || styleSzCs(map, kidVal(rPr, 'rStyle')) || styleSzCs(map, pStyle) || defSzCs;
          if (!thai || thai === sz) continue;                          // ตรงกันอยู่แล้ว ไม่ต้องแตะ
          setSize(dom, rPr, thai);
          changed = true; touched++;
        }
      }
      if (changed) {
        let out = new XMLSerializer().serializeToString(dom);
        const decl = (text.match(/^<\?xml[^>]*\?>/) || [''])[0];
        if (decl && !out.startsWith('<?xml')) out = decl + '\r\n' + out;
        src.file(path, out);
      }
    }
    console.log(`[Doc Hub] เตรียมไฟล์: ปรับขนาดฟอนต์ไทยให้ editor ${touched} จุด`);
    if (!touched) return blob;
    return src.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE'
    });
  }

  window.DocHubDocx = { sanitizeDocx, fixXml, prepareForEditor };
})();
