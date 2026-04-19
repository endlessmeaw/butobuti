# 📋 คู่มือติดตั้ง LOTTOAI PRO Backend

## สิ่งที่สร้างให้

สร้างไฟล์ `Code.gs` — Google Apps Script Backend ที่รองรับการทำงานทั้งหมดของแอป LOTTOAI PRO

### ระบบที่รองรับ

| ระบบ | Action | Method | รายละเอียด |
|------|--------|--------|------------|
| 🔐 สมัครสมาชิก | `register` | POST | เช็คซ้ำ + Hash password ด้วย SHA-256 |
| 🔑 เข้าสู่ระบบ | `login` | POST | ตรวจสอบ username/password |
| 🎯 วิเคราะห์เลข | `analyze` | POST | 75 สูตร × 3 โหมด Engine + บังคับ/แบนเลข |
| 📝 ดึงประวัติ | `getHistory` | GET | ล่าสุด 20 รายการ |

### 75 สูตรวิเคราะห์ แบ่ง 4 กลุ่ม

| กลุ่ม | จำนวนสูตร | วิธีคิด |
|-------|-----------|---------|
| บวกคงที่ | 20 สูตร | บวกตัวเลขคู่ต่างๆ mod 10 |
| คูณทะลวง | 15 สูตร | คูณตัวเลขข้ามตำแหน่ง |
| บวกไขว้ | 25 สูตร | ลบ, ไขว้, รวมกลุ่ม |
| เลขเงา | 15 สูตร | เลขกลับ, +5, คู่เงา |

### 3 โหมด Engine

| โหมด | ชื่อ | วิธีให้คะแนน |
|------|------|-------------|
| A | ความถี่สูงสุด | นับจำนวนครั้งที่เลขออกจาก 75 สูตร |
| B | ลำดับคะแนน | สูตรท้ายมีน้ำหนักมากกว่า (1.0→2.0) |
| C | คะแนนดิบ | น้ำหนักต่างกันตามกลุ่มสูตร |

### Google Sheets ที่ใช้ (2 ชีท)

| Sheet | คอลัมน์ |
|-------|---------|
| **Users** | username, password (SHA-256), created_at |
| **History** | username, type, mode, input, result, time, combosTwo, combosThree |

---

## 🚀 วิธีติดตั้ง (ทำครั้งเดียว)

### ขั้นตอนที่ 1: สร้าง Google Sheets

1. ไปที่ [Google Sheets](https://sheets.google.com)
2. สร้าง Spreadsheet ใหม่ → ตั้งชื่อ **"LOTTOAI Database"**

### ขั้นตอนที่ 2: สร้าง Google Apps Script

1. ใน Google Sheets → คลิก **ส่วนขยาย (Extensions)** → **Apps Script**
2. จะเปิดหน้า Apps Script Editor ขึ้นมา
3. **ลบโค้ดเดิม** ทั้งหมดในไฟล์ `Code.gs`
4. **คัดลอกโค้ดทั้งหมด** จากไฟล์ `Code.gs` ที่อยู่ใน folder นี้ วางลงไป
5. กด **💾 Save** (Ctrl+S)

### ขั้นตอนที่ 3: สร้าง Sheet อัตโนมัติ

1. ในเมนูด้านบน → เลือกฟังก์ชัน **`setupSheets`** จาก dropdown
2. กด **▶ Run**
3. ครั้งแรกจะขอ **Authorization** → กด **Review Permissions** → เลือกบัญชี Google → **Allow**
4. ดู Google Sheets จะมี Sheet "Users" และ "History" ถูกสร้างขึ้น

### ขั้นตอนที่ 4: Deploy เป็น Web App

1. กด **Deploy** (ปุ่มมุมขวาบน) → **New deployment**
2. คลิก ⚙️ → เลือก **Web app**
3. ตั้งค่า:
   - **Description**: `LOTTOAI Backend v1`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. กด **Deploy**
5. **คัดลอก Web App URL** ที่ได้ (จะเป็น `https://script.google.com/macros/s/xxxxx/exec`)

### ขั้นตอนที่ 5: เชื่อมต่อ Frontend

1. เปิดไฟล์ `index.html`
2. ค้นหา `API_URL` → แก้ไขเป็น URL ที่ได้จากการ Deploy:

```javascript
const API_URL = "https://script.google.com/macros/s/xxxxxx_URL_ที่ได้จาก_Deploy/exec";
```

3. บันทึกไฟล์ → เปิดใช้งานได้เลย!

---

## 🧪 ทดสอบ (ไม่จำเป็นต้องทำ)

ใน Apps Script Editor สามารถรันฟังก์ชันทดสอบได้:

| ฟังก์ชัน | ทำอะไร |
|---------|--------|
| `testAnalyze` | ทดสอบวิเคราะห์เลข 835127 |
| `testRegister` | ทดสอบสมัครสมาชิก demo/1234 |
| `testLogin` | ทดสอบเข้าสู่ระบบ demo/1234 |

> เลือกฟังก์ชันจาก dropdown แล้วกด ▶ Run → ดูผลลัพธ์ใน **Execution log** (View → Logs)

---

## ⚠️ หมายเหตุสำคัญ

- ทุกครั้งที่แก้ไขโค้ดใน Apps Script ต้อง **Deploy ใหม่** (New deployment) แล้วเอา URL ใหม่ไปใส่ หรือใช้ **Manage deployments** แก้ version ที่มีอยู่
- Google Apps Script มี Quota จำกัดที่ **90 นาที/วัน** สำหรับ Execution time (ใช้งานปกติไม่เกินแน่นอน)
- ข้อมูลทั้งหมดเก็บใน Google Sheets → สามารถดู/แก้ไข/Export ได้ตลอดเวลา

---

## 📁 รายการไฟล์ทั้งหมดในโปรเจค

| ไฟล์ | ประเภท | รายละเอียด |
|------|--------|------------|
| `index.html` | Frontend | หน้าเว็บหลัก (PWA) |
| `icon.png` | Asset | ไอคอนแอป |
| `manifest.json` | Config | ตั้งค่า PWA |
| `Code.gs` | Backend | Google Apps Script (คัดลอกไปวางใน GAS) |
| `คู่มือติดตั้ง-Backend.md` | เอกสาร | ไฟล์นี้ |
