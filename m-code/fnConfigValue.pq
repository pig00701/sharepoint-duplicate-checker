// fnConfigValue — ฟังก์ชันเปล่า (pure function) ไม่แตะ data source เอง
// lookup ค่าจาก ConfigTable (คอลัมน์ Key/Value) แบบ case/space-insensitive พร้อมค่า default
// ต้องเป็นฟังก์ชันแยกแบบนี้ (ไม่ใช่คิวรีที่อ่าน Excel.CurrentWorkbook เอง) เพื่อเลี่ยง
// Formula.Firewall ตอนถูกเรียกจากคิวรีที่แตะ data source อื่น (SharePoint.Files) — ดู README

(ConfigTable as table, SettingName as text, DefaultValue as any) as any =>
let
    Normalize = (t as text) as text => Text.Trim(Text.Lower(t)),
    Match = Table.SelectRows(
        ConfigTable,
        each Normalize(Text.From([Key])) = Normalize(SettingName)
    ),
    Result = if Table.RowCount(Match) = 0 then DefaultValue else Match{0}[Value]
in
    Result
