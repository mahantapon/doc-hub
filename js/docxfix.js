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

  window.DocHubDocx = { sanitizeDocx, fixXml };
})();
