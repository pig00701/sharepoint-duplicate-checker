// CompareFiles — คิวรีหลัก: อ่าน Master.xlsx และ Current.xlsx จาก SharePoint (read-only)
// แล้วสรุปรายการ New / Missing / Duplicate เทียบตาม key ของแต่ละไฟล์ที่ตั้งค่าไว้ใน tblConfig
// (MasterKeyColumn / CurrentKeyColumn แยกกันได้ — ชื่อคอลัมน์ไม่จำเป็นต้องตรงกันทั้งสองไฟล์
// แค่ต้องเป็น "ค่าเดียวกัน" ที่ใช้เทียบว่าเป็นแถวเดียวกัน เช่น ID เดียวกันแต่หัวคอลัมน์คนละชื่อ)
// ผลลัพธ์โหลดเข้ามาเฉพาะใน workbook นี้เท่านั้น — ไม่เขียนอะไรกลับไปที่ Master.xlsx / Current.xlsx
// บน SharePoint เลย ต้องใช้แค่สิทธิ์ Read บน site/library
//
// SiteUrl ต้องเป็น Power Query Parameter (ไม่ใช่ค่าจาก tblConfig) — เพราะ SharePoint.Files()
// เป็น data source ที่ต้องมี Privacy Level ชัดเจน ต่างจาก Excel.CurrentWorkbook() (local)
// ถ้าส่งค่าที่มาจาก Excel.CurrentWorkbook() เข้า SharePoint.Files() โดยตรง (แม้อยู่คิวรีเดียวกัน)
// จะโดน Formula.Firewall ที่ step นี้ — ดูหัวข้อ "ปัญหา Formula.Firewall ที่ AllFiles" ใน README

let
    ConfigTable = Excel.CurrentWorkbook(){[Name = "tblConfig"]}[Content],

    MasterFileName   = fnConfigValue(ConfigTable, "MasterFileName", "Master.xlsx"),
    CurrentFileName  = fnConfigValue(ConfigTable, "CurrentFileName", "Current.xlsx"),
    SheetName        = fnConfigValue(ConfigTable, "SheetName", "Sheet1"),
    MasterKeyColumn  = fnConfigValue(ConfigTable, "MasterKeyColumn", "ID"),
    CurrentKeyColumn = fnConfigValue(ConfigTable, "CurrentKeyColumn", "ID"),

    _CheckSiteUrl = if SiteUrl = null or SiteUrl = "" then error "ต้องตั้งค่า Parameter 'SiteUrl' ก่อน (Home > Manage Parameters)" else SiteUrl,

    // data source #1: SharePoint library — อ่านอย่างเดียว (SiteUrl มาจาก Parameter ไม่ใช่ tblConfig)
    AllFiles = SharePoint.Files(_CheckSiteUrl, [ApiVersion = 15]),

    MasterMatch = Table.SelectRows(AllFiles, each Text.EndsWith([Name], MasterFileName)),
    CurrentMatch = Table.SelectRows(AllFiles, each Text.EndsWith([Name], CurrentFileName)),

    _CheckMaster = if Table.RowCount(MasterMatch) = 0
        then error "ไม่พบไฟล์ '" & MasterFileName & "' ใน " & _CheckSiteUrl
        else MasterMatch{0},
    _CheckCurrent = if Table.RowCount(CurrentMatch) = 0
        then error "ไม่พบไฟล์ '" & CurrentFileName & "' ใน " & _CheckSiteUrl
        else CurrentMatch{0},

    MasterTable = fnExtractSheet(_CheckMaster[Content], SheetName),
    CurrentTable = fnExtractSheet(_CheckCurrent[Content], SheetName),

    // เลือกด้วยชื่อคอลัมน์ของแต่ละไฟล์เอง แล้วเปลี่ยนชื่อเป็น "Key" กลาง ๆ ให้ตรงกันก่อนเทียบ
    // ไม่ว่าตำแหน่งคอลัมน์หรือชื่อ header ในสองไฟล์จะต่างกันแค่ไหนก็เทียบกันได้ปกติ
    MasterKeys = Table.RenameColumns(
        Table.Distinct(Table.SelectColumns(MasterTable, {MasterKeyColumn})),
        {{MasterKeyColumn, "Key"}}
    ),
    CurrentKeysAll = Table.RenameColumns(
        Table.SelectColumns(CurrentTable, {CurrentKeyColumn}),
        {{CurrentKeyColumn, "Key"}}
    ),
    CurrentKeysDistinct = Table.Distinct(CurrentKeysAll),

    MasterKeyList = Table.Column(MasterKeys, "Key"),
    CurrentKeyList = Table.Column(CurrentKeysDistinct, "Key"),

    // ซ้ำ: key เดิมโผล่มากกว่า 1 แถวใน Current
    DupCounts = Table.Group(CurrentKeysAll, {"Key"}, {{"Count", Table.RowCount, Int64.Type}}),
    Duplicates = Table.SelectRows(DupCounts, each [Count] > 1),
    TagDuplicates = Table.RemoveColumns(
        Table.AddColumn(Duplicates, "Status", each "Duplicate (" & Text.From([Count]) & "x)"),
        {"Count"}
    ),

    // ใหม่: อยู่ใน Current แต่ไม่อยู่ใน Master
    NewRows = Table.SelectRows(CurrentKeysDistinct, each not List.Contains(MasterKeyList, [Key])),
    TagNew = Table.AddColumn(NewRows, "Status", each "New"),

    // หายไป: อยู่ใน Master แต่ไม่อยู่ใน Current
    MissingRows = Table.SelectRows(MasterKeys, each not List.Contains(CurrentKeyList, [Key])),
    TagMissing = Table.AddColumn(MissingRows, "Status", each "Missing"),

    Report = Table.Combine({TagDuplicates, TagNew, TagMissing})
in
    Report
