// Tab 1: Excel <-> Google Sheet batch converter (Drive API, client-side only)
'use strict';
(() => {
  const LS_CID = 'dochub_client_id';
  const LS_KEY = 'dochub_api_key';
  // ค่ากลางของบริษัท (โปรเจ็กต์ doc-hub-503212, consent แบบ Internal @ruanganan.com)
  // เป็นค่าเปิดเผยได้ตามธรรมชาติของเว็บแอป — ถูกจำกัดโดเมน + จำกัดองค์กรแล้ว
  const DEFAULT_CID = '859658791030-7fmo965td9vngfei37di9u8gjpitqbp7.apps.googleusercontent.com';
  const DEFAULT_KEY = 'AIzaSyAy0wndU_wr5RZf-TpW5NB0p510nlEnNVE';
  const SCOPE = 'https://www.googleapis.com/auth/drive.file';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const GSHEET_MIME = 'application/vnd.google-apps.spreadsheet';

  let accessToken = null, tokenExp = 0, tokenClient = null, pickerReady = false;

  const cfg = () => ({
    cid: localStorage.getItem(LS_CID) || DEFAULT_CID,
    key: localStorage.getItem(LS_KEY) || DEFAULT_KEY
  });
  const esc = s => String(s).replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

  // ---------- config ----------
  function renderCfgBar() {
    const bar = $('#excelCfgBar');
    const c = cfg();
    bar.className = c.cid && c.key ? 'cfgbar ok' : 'cfgbar';
    bar.innerHTML = c.cid && c.key
      ? '✅ พร้อมใช้งาน — ครั้งแรกจะมีหน้าต่างล็อกอิน Google เด้งขึ้นมา (ใช้บัญชี @ruanganan.com) '
      : '⚠️ แท็บนี้ต้องตั้งค่าเชื่อม Google Drive ก่อน (ทำครั้งเดียว ~5 นาที ดูขั้นตอนใน SETUP.md) ';
    const b = document.createElement('button');
    b.className = 'btn';
    b.textContent = '⚙️ ตั้งค่า';
    b.addEventListener('click', openCfg);
    bar.appendChild(b);
  }
  function openCfg() {
    $('#cfgClientId').value = cfg().cid;
    $('#cfgApiKey').value = cfg().key;
    $('#cfgDialog').showModal();
  }
  $('#cfgSave').addEventListener('click', () => {
    localStorage.setItem(LS_CID, $('#cfgClientId').value.trim());
    localStorage.setItem(LS_KEY, $('#cfgApiKey').value.trim());
    $('#cfgDialog').close();
    renderCfgBar();
  });
  $('#cfgCancel').addEventListener('click', () => $('#cfgDialog').close());

  // ---------- auth ----------
  let tokenReject = null;
  function ensureToken() {
    return new Promise((resolve, reject) => {
      const c = cfg();
      if (!c.cid) return reject(new Error('ยังไม่ได้ตั้งค่า — กดปุ่ม ⚙️ ตั้งค่า มุมขวาบนของแถบเหลือง'));
      if (accessToken && Date.now() < tokenExp - 60000) return resolve(accessToken);
      if (!window.google || !google.accounts || !google.accounts.oauth2)
        return reject(new Error('สคริปต์ Google ยังโหลดไม่เสร็จ — รอ 2-3 วินาทีแล้วลองใหม่'));
      if (!tokenClient)
        tokenClient = google.accounts.oauth2.initTokenClient({
          client_id: c.cid, scope: SCOPE, callback: () => {},
          error_callback: e => {
            const fn = tokenReject; tokenReject = null;
            if (fn) fn(new Error(e && e.type === 'popup_failed_to_open'
              ? 'เบราว์เซอร์บล็อกหน้าต่างล็อกอิน — กดไอคอน pop-up ขวาบนของช่องที่อยู่ เลือก "อนุญาต" แล้วลองใหม่'
              : 'ล็อกอินไม่สำเร็จ: ' + (e && (e.message || e.type) || 'unknown')));
          }
        });
      tokenReject = reject;
      tokenClient.callback = resp => {
        tokenReject = null;
        if (resp.error) return reject(new Error('ล็อกอินไม่สำเร็จ: ' + resp.error));
        accessToken = resp.access_token;
        tokenExp = Date.now() + (resp.expires_in || 3600) * 1000;
        resolve(accessToken);
      };
      tokenClient.requestAccessToken({ prompt: '' });
    });
  }

  async function api(url, opts = {}) {
    const token = await ensureToken();
    const res = await fetch(url, { ...opts, headers: { Authorization: 'Bearer ' + token, ...(opts.headers || {}) } });
    if (!res.ok) {
      let detail = 'HTTP ' + res.status;
      try { detail = (await res.json()).error.message; } catch (_) {}
      throw new Error(detail);
    }
    return res;
  }

  // ---------- shared result rows ----------
  function resultRow(container, nameHtml) {
    const row = document.createElement('div');
    row.className = 'row';
    row.innerHTML = `<span class="st">⏳</span><span class="nm">${nameHtml}</span><span class="ex"></span>`;
    container.prepend(row);
    return {
      ok(html) { row.querySelector('.st').textContent = '✅'; if (html) row.querySelector('.ex').innerHTML = html; },
      fail(err) { row.querySelector('.st').textContent = '❌'; row.querySelector('.ex').innerHTML = `<span class="err">${err}</span>`; }
    };
  }

  // ---------- import: xlsx -> Google Sheet ----------
  async function ensureFolder(name) {
    const q = encodeURIComponent(
      `name='${name.replace(/'/g, "\\'")}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
    const found = await (await api(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`)).json();
    if (found.files && found.files.length) return found.files[0].id;
    const created = await (await api('https://www.googleapis.com/drive/v3/files?fields=id', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder' })
    })).json();
    return created.id;
  }

  async function importFiles(files) {
    files = files.filter(f => /\.(xlsx|xlsm|xls)$/i.test(f.name));
    const cont = $('#importResults');
    if (!files.length) { resultRow(cont, 'ไม่มีไฟล์ Excel ในรายการที่เลือก').fail('รองรับ .xlsx .xlsm .xls'); return; }
    let folderId;
    try {
      await ensureToken();
      // ถ้าเลือกโฟลเดอร์จาก Drive/Shared Drive แล้ว ใช้อันนั้น — ไม่งั้นสร้างตามชื่อใน My Drive
      folderId = destFolder ? destFolder.id
        : await ensureFolder($('#folderName').value.trim() || 'DocHub Import');
    } catch (e) { resultRow(cont, 'เชื่อม Drive').fail(esc(e.message)); return; }
    for (const f of files) {
      const row = resultRow(cont, esc(f.name));
      try {
        const meta = {
          name: f.name.replace(/\.(xlsx|xlsm|xls)$/i, ''),
          mimeType: GSHEET_MIME,
          parents: [folderId]
        };
        const boundary = 'dochub' + Math.random().toString(36).slice(2);
        const body = new Blob([
          `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(meta)}\r\n`,
          `--${boundary}\r\nContent-Type: ${XLSX_MIME}\r\n\r\n`,
          f,
          `\r\n--${boundary}--`
        ]);
        const res = await api('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,webViewLink', {
          method: 'POST',
          headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
          body
        });
        const j = await res.json();
        row.ok(`<a href="${j.webViewLink}" target="_blank" rel="noopener">เปิด Sheet</a>`);
      } catch (e) { row.fail(esc(e.message)); }
    }
  }

  // ---------- export: Google Sheet -> xlsx ----------
  function loadPicker() {
    return new Promise((resolve, reject) => {
      if (pickerReady) return resolve();
      if (!window.gapi) return reject(new Error('สคริปต์ Google ยังโหลดไม่เสร็จ — รอ 2-3 วินาทีแล้วลองใหม่'));
      gapi.load('picker', {
        callback: () => { pickerReady = true; resolve(); },
        onerror: () => reject(new Error('โหลดหน้าต่างเลือกไฟล์ไม่สำเร็จ'))
      });
    });
  }

  async function pickSheets() {
    const cont = $('#exportResults');
    try {
      const token = await ensureToken();
      await loadPicker();
      const c = cfg();
      if (!c.key) throw new Error('ยังไม่ได้ใส่ API Key — กดปุ่ม ⚙️ ตั้งค่า');
      // แท็บแรก: My Drive ส่วนตัว + ไฟล์ที่แชร์มาให้ (ห้ามใส่ setEnableDrives ไม่งั้นกลายเป็น Shared Drive)
      const view = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS)
        .setIncludeFolders(true)
        .setMode(google.picker.DocsViewMode.LIST);
      // แท็บสอง: Shared Drives ของบริษัท
      const driveView = new google.picker.DocsView(google.picker.ViewId.SPREADSHEETS)
        .setEnableDrives(true)
        .setIncludeFolders(true)
        .setMode(google.picker.DocsViewMode.LIST);
      new google.picker.PickerBuilder()
        .addView(view)
        .addView(driveView)
        .setOAuthToken(token)
        .setDeveloperKey(c.key)
        .setAppId(c.cid.split('-')[0])
        .enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
        .enableFeature(google.picker.Feature.SUPPORT_DRIVES)
        .setCallback(data => {
          if (data[google.picker.Response.ACTION] === google.picker.Action.PICKED)
            exportSheets(data[google.picker.Response.DOCUMENTS]);
        })
        .build()
        .setVisible(true);
    } catch (e) { resultRow(cont, 'เลือกไฟล์จาก Drive').fail(esc(e.message)); }
  }

  // ---------- destination folder picker (รองรับ Shared Drives) ----------
  let destFolder = null; // { id, name }
  function renderDest() {
    const el = $('#destLabel');
    if (destFolder) {
      el.innerHTML = `📁 <b>${esc(destFolder.name)}</b> <button id="clearDest" class="btn ghost" style="padding:2px 8px">✖</button>`;
      $('#clearDest').addEventListener('click', () => { destFolder = null; renderDest(); });
      $('#folderName').disabled = true;
    } else {
      el.textContent = '';
      $('#folderName').disabled = false;
    }
  }

  async function pickFolder() {
    const cont = $('#importResults');
    try {
      const token = await ensureToken();
      await loadPicker();
      const c = cfg();
      // แท็บแรก: โฟลเดอร์ใน My Drive · แท็บสอง: โฟลเดอร์ใน Shared Drives
      const view = new google.picker.DocsView(google.picker.ViewId.FOLDERS)
        .setSelectFolderEnabled(true)
        .setIncludeFolders(true)
        .setMode(google.picker.DocsViewMode.LIST);
      const driveFolderView = new google.picker.DocsView(google.picker.ViewId.FOLDERS)
        .setSelectFolderEnabled(true)
        .setIncludeFolders(true)
        .setEnableDrives(true)
        .setMode(google.picker.DocsViewMode.LIST);
      new google.picker.PickerBuilder()
        .addView(view)
        .addView(driveFolderView)
        .setOAuthToken(token)
        .setDeveloperKey(c.key)
        .setAppId(c.cid.split('-')[0])
        .enableFeature(google.picker.Feature.SUPPORT_DRIVES)
        .setTitle('เลือกโฟลเดอร์ปลายทาง (My Drive หรือ Shared Drive)')
        .setCallback(data => {
          if (data[google.picker.Response.ACTION] === google.picker.Action.PICKED) {
            const d = data[google.picker.Response.DOCUMENTS][0];
            destFolder = { id: d.id, name: d.name };
            renderDest();
          }
        })
        .build()
        .setVisible(true);
    } catch (e) { resultRow(cont, 'เลือกโฟลเดอร์').fail(esc(e.message)); }
  }

  async function exportSheets(docs) {
    const cont = $('#exportResults');
    const out = [];
    for (const d of docs) {
      const row = resultRow(cont, esc(d.name));
      try {
        const res = await api(
          `https://www.googleapis.com/drive/v3/files/${d.id}/export?mimeType=${encodeURIComponent(XLSX_MIME)}`);
        out.push({ name: d.name + '.xlsx', blob: await res.blob() });
        row.ok();
      } catch (e) { row.fail(esc(e.message)); }
    }
    if (!out.length) return;
    if (out.length === 1) downloadBlob(out[0].blob, out[0].name);
    else {
      const zip = new JSZip();
      out.forEach(o => zip.file(o.name, o.blob));
      downloadBlob(await zip.generateAsync({ type: 'blob' }), 'sheets-export.zip');
    }
  }

  // ---------- wiring ----------
  renderCfgBar();
  wireDropzone($('#dropXlsx'), $('#fileXlsx'), importFiles);
  $('#pickXlsx').addEventListener('click', () => $('#fileXlsx').click());
  $('#btnPickSheets').addEventListener('click', pickSheets);
  $('#btnPickFolder').addEventListener('click', pickFolder);
})();
