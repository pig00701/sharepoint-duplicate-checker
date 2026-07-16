# Power Query — เช็คไฟล์ซ้ำ/ใหม่/หาย บน SharePoint (Read-Only)

โปรเจกต์ Power Query สำหรับเปรียบเทียบไฟล์ Master กับไฟล์ Current บน SharePoint แล้วสรุปรายการ
ซ้ำ/ใหม่/หายไปลงไฟล์ Check.xlsx — **เป็นการอ่านข้อมูลเท่านั้น ไม่มีการเขียนกลับไปที่ไฟล์ต้นทาง**

---

## สรุปพฤติกรรม: read-only เสมอ

- `SharePoint.Files()` และ `Excel.Workbook()` แค่ **เปิดอ่านเนื้อหาไฟล์มาประมวลผลในหน่วยความจำ**
  ของ Power Query เท่านั้น ไม่มีฟังก์ชันเขียนกลับไปที่ไฟล์ต้นทาง
- ผลลัพธ์ (รายการซ้ำ/ใหม่/หายไป/รายงาน) โหลดเข้ามาอยู่ใน **ไฟล์ Check.xlsx เท่านั้น**
- ไฟล์ `Master.xlsx` และ `Current.xlsx` บน SharePoint **ไม่ถูกแตะต้องเลย** ไม่ว่าจะกด Refresh กี่ครั้ง
  ก็ตาม — แค่ไป "อ่านใหม่" แล้วคำนวณใหม่ทุกครั้ง ไม่มีการเขียนอะไรกลับ

## สิทธิ์ที่ต้องใช้

แค่สิทธิ์ **Read** บน SharePoint site/library ก็เพียงพอ ไม่จำเป็นต้องมีสิทธิ์ Edit บนไฟล์
Master/Current เลย — เหมาะกับกรณีที่ไฟล์เหล่านั้นมีคนอื่นดูแลอยู่ และไม่อยากให้กระบวนการเช็ค
ไปกระทบไฟล์เขา

---

## โครงสร้างโปรเจกต์

```
sharepoint-duplicate-checker/
├── m-code/
│   ├── fnConfigValue.pq     # ฟังก์ชันเปล่า: lookup ค่าจาก tblConfig
│   ├── fnExtractSheet.pq    # ฟังก์ชันเปล่า: แกะ binary ไฟล์ → ตาราง
│   └── CompareFiles.pq      # คิวรีหลัก: เทียบ Master vs Current → New/Missing/Duplicate
└── README.md
```

## วิธีติดตั้งใน Check.xlsx

### Step 1: สร้าง `tblConfig` (sheet "Config")

Insert > Table > 2 คอลัมน์ (Key, Value) > ตั้งชื่อ Table เป็น **`tblConfig`**

| Key | Value | หมายเหตุ |
|---|---|---|
| SiteUrl | `https://yourtenant.sharepoint.com/sites/YourSite` | URL ของ SharePoint site ที่เก็บ Master/Current |
| MasterFileName | `Master.xlsx` | ชื่อไฟล์ (หรือส่วนท้ายของ path) ที่จะ match ด้วย `Text.EndsWith` |
| CurrentFileName | `Current.xlsx` | เหมือนกัน สำหรับไฟล์ current |
| SheetName | `Sheet1` | ชื่อ sheet ที่มีข้อมูลจริงในทั้งสองไฟล์ (ต้องชื่อตรงกัน) |
| KeyColumn | `ID` | ชื่อคอลัมน์ที่ใช้เทียบว่าเป็นแถวเดียวกัน |

### Step 2: โหลดฟังก์ชัน 2 ตัว (Connection Only)

1. Blank Query → ตั้งชื่อ `fnConfigValue` → Advanced Editor → วางโค้ดจาก `m-code/fnConfigValue.pq`
2. Blank Query → ตั้งชื่อ `fnExtractSheet` → Advanced Editor → วางโค้ดจาก `m-code/fnExtractSheet.pq`

### Step 3: โหลด `CompareFiles`

Blank Query → ตั้งชื่อ `CompareFiles` → วางโค้ดจาก `m-code/CompareFiles.pq` → **Close & Load**

ผลลัพธ์เป็นตารางคอลัมน์ `KeyColumn` + `Status` (`New` / `Missing` / `Duplicate (Nx)`) — แก้ค่าใน
`tblConfig` แล้ว Refresh ได้เลย ไม่ต้องเปิด M code อีก

## ทำไมไม่โดน Formula.Firewall

`CompareFiles` อ่านทั้ง `Excel.CurrentWorkbook()` (data source #1: workbook ปัจจุบัน) และ
`SharePoint.Files()` (data source #2: SharePoint) **จากภายในคิวรีเดียวกัน** ส่วน
`fnConfigValue` / `fnExtractSheet` เป็นฟังก์ชันเปล่า (pure function) ที่ไม่แตะ data source เอง
— รับค่า/binary ที่ส่งเข้ามาแล้วประมวลผลอย่างเดียว จึงไม่ถูกนับเป็นการข้าม data source ระหว่าง
คิวรี (รูปแบบเดียวกับที่ใช้แก้ปัญหานี้ใน `excel_proj`)
