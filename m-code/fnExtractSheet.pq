// fnExtractSheet — ฟังก์ชันเปล่า รับ "เนื้อไฟล์" (binary) ที่ดึงมาแล้วจาก SharePoint.Files()
// มาแกะเป็นตาราง ไม่ได้เปิด data source ของตัวเอง (Excel.Workbook บน binary ที่ส่งเข้ามา
// ไม่นับเป็น data source ใหม่ ต่างจากการส่ง path/URL ตรงๆ) จึงเรียกจากคิวรีหลักได้โดยไม่โดน
// Formula.Firewall

(Content as binary, SheetName as text) as table =>
let
    Workbook = Excel.Workbook(Content, null, true),
    SheetRow = Table.SelectRows(Workbook, each [Item] = SheetName and [Kind] = "Sheet"),
    Sheet = if Table.RowCount(SheetRow) = 0
        then error "ไม่พบ sheet '" & SheetName & "' ในไฟล์นี้"
        else SheetRow{0}[Data],
    Promoted = Table.PromoteHeaders(Sheet, [PromoteAllScalars = true])
in
    Promoted
