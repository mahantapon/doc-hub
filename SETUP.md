# การตั้งค่าเชื่อม Google Drive

> ✅ **ตั้งค่าเสร็จแล้ว (22 ก.ค. 2026)** — Client ID + API Key ของบริษัทถูกฝังในเว็บแล้ว (โปรเจ็กต์ `doc-hub-503212`, consent แบบ Internal จำกัดเฉพาะบัญชี @ruanganan.com)
> **ทีมงานไม่ต้องทำอะไรตามเอกสารนี้** — เปิดเว็บ กดใช้ ล็อกอิน Google ครั้งแรกครั้งเดียวจบ
> เอกสารด้านล่างเก็บไว้เป็นคู่มืออ้างอิง เผื่อต้องสร้างใหม่ในอนาคต

ใช้กับแท็บ **"แปลง Excel ⇄ Sheet"** เท่านั้น — แท็บอื่นใช้ได้เลยไม่ต้องตั้งค่าอะไร

## ขั้นตอน

1. เข้า [Google Cloud Console](https://console.cloud.google.com) ด้วยบัญชี Google ของบริษัท
2. สร้างโปรเจ็กต์ใหม่ ตั้งชื่อ `doc-hub`
3. เมนู **APIs & Services → Library** → เปิดใช้งาน (Enable) 2 ตัว:
   - **Google Drive API**
   - **Google Picker API**
4. เมนู **APIs & Services → OAuth consent screen**:
   - ถ้าใช้บัญชี Google Workspace ขององค์กร → เลือก **Internal** (ง่ายสุด ไม่ต้องผ่านการตรวจสอบ และจำกัดเฉพาะคนในองค์กรอัตโนมัติ)
   - ถ้าเลือก Internal ไม่ได้ → เลือก **External** แล้วไปที่หน้า Test users เพิ่มอีเมลทีมงานทุกคนที่จะใช้
   - App name: `Doc Hub` → กด Save
5. เมนู **APIs & Services → Credentials → Create Credentials → OAuth client ID**:
   - Application type: **Web application**
   - Authorized JavaScript origins เพิ่ม 2 อัน:
     - `https://mahantapon.github.io`
     - `http://localhost:8090` (สำหรับทดสอบในเครื่อง)
   - กด Create แล้วคัดลอก **Client ID** (ลงท้าย `.apps.googleusercontent.com`)
6. **Create Credentials → API key** → คัดลอก **API key** (ขึ้นต้น `AIza`)
   - (แนะนำ) กด Edit ตัว key → Application restrictions: Websites → ใส่ `https://mahantapon.github.io/*` — กันคนอื่นแอบเอา key ไปใช้
7. เปิดเว็บ Doc Hub → แท็บ "แปลง Excel ⇄ Sheet" → กดปุ่ม **⚙️ ตั้งค่า** → วาง Client ID กับ API key → บันทึก

เสร็จแล้ว — กดใช้งานครั้งแรกจะมีหน้าต่างให้ล็อกอิน Google หนึ่งครั้ง

## หมายเหตุ

- ค่าที่ตั้งจะถูกจำไว้ในเบราว์เซอร์ของแต่ละคน (localStorage) — เพื่อนร่วมทีมแต่ละคนวาง Client ID / API key ชุดเดียวกันนี้ได้เลย
- สิทธิ์ที่ขอคือ `drive.file` = เว็บเห็นเฉพาะไฟล์ที่เว็บสร้างเอง + ไฟล์ที่เราเลือกผ่านหน้าต่างเลือกไฟล์เท่านั้น ไม่เห็นไฟล์อื่นใน Drive
- ไม่มีค่าใช้จ่าย — Drive API ฟรีในปริมาณการใช้งานระดับนี้
