// CompareFiles — คิวรีหลัก: อ่าน Master.xlsx และ Current.xlsx จาก SharePoint (read-only)
// แล้วสรุปรายการ New / Missing / Duplicate เทียบตาม KeyColumn ที่ตั้งค่าไว้ใน tblConfig
// ผลลัพธ์โหลดเข้ามาเฉพาะใน workbook นี้ (Check.xlsx) เท่านั้น — ไม่เขียนอะไรกลับไปที่
// Master.xlsx / Current.xlsx บน SharePoint เลย ต้องใช้แค่สิทธิ์ Read บน site/library

let
    ConfigTable = Excel.CurrentWorkbook(){[Name = "tblConfig"]}[Content],

    SiteUrl        = fnConfigValue(ConfigTable, "SiteUrl", null),
    MasterFileName = fnConfigValue(ConfigTable, "MasterFileName", "Master.xlsx"),
    CurrentFileName = fnConfigValue(ConfigTable, "CurrentFileName", "Current.xlsx"),
    SheetName      = fnConfigValue(ConfigTable, "SheetName", "Sheet1"),
    KeyColumn      = fnConfigValue(ConfigTable, "KeyColumn", "ID"),

    _CheckSiteUrl = if SiteUrl = null then error "ต้องตั้งค่า SiteUrl ใน tblConfig ก่อน" else SiteUrl,

    // data source #1: SharePoint library — อ่านอย่างเดียว
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

    MasterKeys = Table.Distinct(Table.SelectColumns(MasterTable, {KeyColumn})),
    CurrentKeysAll = Table.SelectColumns(CurrentTable, {KeyColumn}),
    CurrentKeysDistinct = Table.Distinct(CurrentKeysAll),

    MasterKeyList = Table.Column(MasterKeys, KeyColumn),
    CurrentKeyList = Table.Column(CurrentKeysDistinct, KeyColumn),

    // ซ้ำ: key เดิมโผล่มากกว่า 1 แถวใน Current
    DupCounts = Table.Group(CurrentKeysAll, {KeyColumn}, {{"Count", Table.RowCount, Int64.Type}}),
    Duplicates = Table.SelectRows(DupCounts, each [Count] > 1),
    TagDuplicates = Table.RemoveColumns(
        Table.AddColumn(Duplicates, "Status", each "Duplicate (" & Text.From([Count]) & "x)"),
        {"Count"}
    ),

    // ใหม่: อยู่ใน Current แต่ไม่อยู่ใน Master
    NewRows = Table.SelectRows(CurrentKeysDistinct, each not List.Contains(MasterKeyList, Record.Field(_, KeyColumn))),
    TagNew = Table.AddColumn(NewRows, "Status", each "New"),

    // หายไป: อยู่ใน Master แต่ไม่อยู่ใน Current
    MissingRows = Table.SelectRows(MasterKeys, each not List.Contains(CurrentKeyList, Record.Field(_, KeyColumn))),
    TagMissing = Table.AddColumn(MissingRows, "Status", each "Missing"),

    Report = Table.Combine({TagDuplicates, TagNew, TagMissing})
in
    Report
