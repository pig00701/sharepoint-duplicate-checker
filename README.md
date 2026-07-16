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

## วิธีติดตั้ง

### Step 1: สร้าง Parameter `SiteUrl`

Power Query Editor → **Home > Manage Parameters > New Parameter**
- Name: `SiteUrl`
- Type: `Text`
- Current Value: `https://yourtenant.sharepoint.com/sites/YourSite`

> ต้องเป็น **Parameter** ไม่ใช่ค่าจาก `tblConfig` — ดูเหตุผลในหัวข้อ
> "ปัญหา Formula.Firewall ที่ step AllFiles" ด้านล่าง

### Step 2: สร้าง `tblConfig` (sheet "Config")

Insert > Table > 2 คอลัมน์ (Key, Value) > ตั้งชื่อ Table เป็น **`tblConfig`**

| Key | Value | หมายเหตุ |
|---|---|---|
| MasterFileName | `Master.xlsx` | ชื่อไฟล์ (หรือส่วนท้ายของ path) ที่จะ match ด้วย `Text.EndsWith` |
| CurrentFileName | `Current.xlsx` | เหมือนกัน สำหรับไฟล์ current |
| SheetName | `Sheet1` | ชื่อ sheet ที่มีข้อมูลจริงในทั้งสองไฟล์ (ต้องชื่อตรงกัน) |
| MasterKeyColumn | `ID` | ชื่อคอลัมน์ใน **Master** ที่ใช้เป็น key เทียบว่าเป็นแถวเดียวกัน |
| CurrentKeyColumn | `ID` | ชื่อคอลัมน์ใน **Current** ที่ใช้เป็น key — ตั้งแยกจาก Master ได้ ไม่ต้องสะกดตรงกัน (เช่น Master ใช้ `ID`, Current ใช้ `รหัส`) |

> **ตำแหน่งคอลัมน์ (column position/letter) ไม่มีผลเลย** — query เลือกด้วย**ชื่อ header**
> เสมอ (ผ่าน `Table.SelectColumns`) ไม่ว่า key จะอยู่คอลัมน์ไหนของแต่ละไฟล์ก็เทียบกันได้ปกติ
> ข้อกำหนดเดียวคือต้องรู้ **ชื่อ header จริง** ของแต่ละไฟล์มาใส่ใน `MasterKeyColumn` /
> `CurrentKeyColumn` ให้ตรง (ชื่อจะเหมือนกันหรือต่างกันระหว่างสองไฟล์ก็ได้)

### Step 3: โหลดฟังก์ชัน 2 ตัว (Connection Only)

1. Blank Query → ตั้งชื่อ `fnConfigValue` → Advanced Editor → วางโค้ดจาก `m-code/fnConfigValue.pq`
2. Blank Query → ตั้งชื่อ `fnExtractSheet` → Advanced Editor → วางโค้ดจาก `m-code/fnExtractSheet.pq`

### Step 4: โหลด `CompareFiles`

Blank Query → ตั้งชื่อ `CompareFiles` → วางโค้ดจาก `m-code/CompareFiles.pq` → **Close & Load**

ผลลัพธ์เป็นตารางคอลัมน์ `Key` + `Status` (`New` / `Missing` / `Duplicate (Nx)`) — แก้ค่าใน
`tblConfig` แล้ว Refresh ได้เลย ไม่ต้องเปิด M code อีก (ยกเว้น `SiteUrl` ที่แก้ผ่าน Manage Parameters)

---

## ปัญหา Formula.Firewall ที่ step `AllFiles`

### อาการ

Refresh `CompareFiles` แล้วเจอ error ที่ step `AllFiles`:

```
Formula.Firewall: Query 'CompareFiles' (step 'AllFiles') references
other queries or steps, so it may not directly access a data source.
```

### สาเหตุ

`AllFiles = SharePoint.Files(_CheckSiteUrl, ...)` รับค่า `_CheckSiteUrl` ที่ไหลมาจาก
`ConfigTable = Excel.CurrentWorkbook(){...}` — คือค่าที่มาจาก **data source หนึ่ง (workbook
ปัจจุบัน)** ถูกส่งต่อเข้า **data source อีกตัว (SharePoint)** โดยตรง

ต่างจากเคส `Folder.Files()` ใน `excel_proj` ที่รวมกับ `Excel.CurrentWorkbook()` ในคิวรีเดียวกัน
แล้วผ่านได้ปกติ (เพราะทั้งคู่เป็น local source, privacy level "None" เหมือนกัน) — แต่
`SharePoint.Files()` เป็น **network source ที่บังคับมี Privacy Level** (Organizational /
Public / Private) ทำให้ Formula Firewall เข้มงวดกว่า แม้จะรวมทุกอย่างไว้ในคิวรีเดียวกันแล้วก็ตาม
เทคนิค "รวมคิวรีเดียว + ฟังก์ชันเปล่า" ที่ใช้ได้กับ `Folder.Files()` จึงใช้ไม่ได้กับ
`SharePoint.Files()`

### วิธีแก้ (ใช้ในโค้ดปัจจุบัน): ทำ `SiteUrl` เป็น Power Query Parameter

Parameter ไม่ถูกนับเป็น data source เลย — เป็นค่าคงที่ที่ engine รู้อยู่แล้วก่อนรัน query
จึงไม่มี "การไหลข้าม data source" ให้ Formula Firewall ต้องเช็ค (`_CheckSiteUrl` ใน
`CompareFiles.pq` อ้าง Parameter `SiteUrl` ตรงๆ ไม่ผ่าน `tblConfig`/`fnConfigValue` อีกต่อไป)
วิธีเดียวกับที่โปรเจกต์ `power-query-excel-consolidator` ใช้กับ `FolderPath`

### ทางเลือก: ปิด privacy check แทน (ไม่ต้องแก้โค้ด)

**File → Options and Settings → Query Options → Privacy → "Always ignore Privacy Level
settings"** — เป็นการตั้งค่าระดับเครื่อง ไม่ติดไปกับไฟล์ Excel คนอื่นที่เปิดไฟล์เดียวกันต้อง
ตั้งเองด้วย จึงเสถียรน้อยกว่าวิธี Parameter ถ้าต้องแชร์ไฟล์ให้หลายคน

## ทำไมส่วนที่เหลือไม่โดน Formula.Firewall

`MasterFileName` / `CurrentFileName` / `SheetName` / `MasterKeyColumn` / `CurrentKeyColumn`
(มาจาก `tblConfig` ผ่าน `fnConfigValue`) ไม่ได้ถูกส่งเข้าฟังก์ชัน data source โดยตรง — ใช้แค่
เทียบ string ใน `Table.SelectRows` (กรองตารางที่ `SharePoint.Files()` คืนมาแล้ว) และเลือก
คอลัมน์ในตารางที่แกะออกมาแล้วเท่านั้น ส่วน `fnExtractSheet` รับ **binary content** ที่ดึงมาแล้ว
(ไม่ใช่ path/URL) เข้า `Excel.Workbook()` — ไม่นับเป็นการเปิด data source ใหม่ จึงไม่ติด
Firewall เหมือน `AllFiles`
