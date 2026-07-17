# SharePoint Duplicate Checker — ฉบับ VBA

Port ของ `CompareFiles.pq` (Power Query) มาเป็น VBA — เปิด Master/Current **read-only**
เปรียบเทียบด้วย key column ที่ตั้งค่าได้แยกกันสองไฟล์ แล้วสรุปผลลง sheet `Report`

## ความต่างจากฉบับ Power Query

- **ไม่ต่อ SharePoint URL ตรงๆ** — `Workbooks.Open` กับ path แบบ `https://` ไม่เสถียร
  (ต่างจาก `SharePoint.Files()` ที่ authenticate เองได้) ต้องชี้ `MasterFilePath` /
  `CurrentFilePath` ไปที่ **โฟลเดอร์ SharePoint/OneDrive ที่ sync ไว้ในเครื่องแล้ว** แทน
  เช่น `C:\Users\ชื่อคุณ\OneDrive - บริษัท\ชื่อ Library\Master.xlsx`
- ไม่มีปัญหา Formula.Firewall เลย (เป็น concept เฉพาะ Power Query) — VBA เปิดไฟล์ตามลำดับ
  คำสั่งตรงๆ
- Key column ของ Master กับ Current **ตั้งชื่อคนละอย่างกันได้** เหมือนฉบับ Power Query
  (ดูคอมเมนต์ใน `modConfig.bas`)

## โมดูล

| ไฟล์ | หน้าที่ |
|---|---|
| `modUtils.bas` | `FindListObject`, `IsBlankValue`, `ReadKeyColumnValues` (หา column ด้วยชื่อ header แถว 1 แล้วอ่านค่าไม่ว่างทั้งหมด) |
| `modConfig.bas` | อ่าน `tblConfig` (sheet "Config") พร้อม default ทุกค่ายกเว้น `MasterFilePath`/`CurrentFilePath` ที่บังคับ + เช็คกัน path แบบ https |
| `modCompare.bas` | **`RunCompareFiles`** — entrypoint หลัก: เปิดสองไฟล์ (read-only) → อ่าน key column ของแต่ละไฟล์ → ปิดไฟล์ทันที (ไม่ save) → เทียบหา New/Missing/Duplicate → เขียนผลลง sheet `Report` |

## ติดตั้ง

1. เปิดไฟล์รายงาน (บันทึกเป็น **.xlsm**) → `Alt+F11` เปิด VBA Editor
2. สร้าง module ใหม่ 3 อัน (Insert → Module) ตั้งชื่อตามตาราง แล้ว **copy-paste เนื้อหาแต่ละไฟล์ลงไป**
   (ข้ามบรรทัด `Attribute VB_Name = ...` บนสุด — VBE ใส่ให้เองตอนตั้งชื่อ module)
   > ใช้วิธี copy-paste แทน File → Import เพราะไฟล์ .bas ใน repo เก็บเป็น UTF-8 แต่ VBE
   > import แบบ ANSI — ข้อความภาษาไทยในโค้ดจะเพี้ยนถ้า import ตรงๆ
3. สร้าง sheet **Config** → ทำตารางคอลัมน์ `Key` / `Value` (Ctrl+T, ตั้งชื่อตาราง `tblConfig`):

   | Key | Value |
   |---|---|
   | MasterFilePath | `C:\Users\...\OneDrive - บริษัท\Library\Master.xlsx` |
   | CurrentFilePath | `C:\Users\...\OneDrive - บริษัท\Library\Current.xlsx` |
   | SheetName | `Sheet1` |
   | MasterKeyColumn | `ID` |
   | CurrentKeyColumn | `ID` |
   | MasterKeyColumnStartCol | `1` |
   | CurrentKeyColumnStartCol | `1` |

4. รัน `RunCompareFiles` → ดูผลใน sheet `Report`
5. (แนะนำ) วางปุ่มบน sheet Config ผูกกับ `RunCompareFiles`

## การรัน Test

End-to-end test ผ่าน Excel COM automation (`tests/vba_e2e_test.py`) — สร้าง Master.xlsx /
Current.xlsx ตัวอย่างจริง, import โมดูลทั้ง 3, รัน `RunCompareFiles`, แล้วเช็คว่า sheet
`Report` มีผลลัพธ์ตรงตามที่คำนวณไว้ล่วงหน้า (รวมถึงเคส key column ชื่อไม่ตรงกันระหว่างสองไฟล์):

```powershell
python tests/vba_e2e_test.py
```

ต้องมี `pywin32` และ `openpyxl` (`pip install pywin32 openpyxl`) และเครื่องต้องมี Excel ติดตั้งจริง
