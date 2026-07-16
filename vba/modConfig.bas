Attribute VB_Name = "modConfig"
' ============================================================================
' modConfig.bas
' ============================================================================
' Parameter-sheet pattern: all settings live in a worksheet table so users
' change behavior from the sheet and re-run — never by editing VBA.
'
' tblConfig (sheet "Config", Ctrl+T, table name "tblConfig"):
'   Key               | Value
'   MasterFilePath     | C:\...\OneDrive - Org\Library\Master.xlsx
'   CurrentFilePath    | C:\...\OneDrive - Org\Library\Current.xlsx
'   SheetName          | Sheet1
'   MasterKeyColumn    | ID
'   CurrentKeyColumn   | ID
'
' Rules:
'   - Key names match case-insensitively, ignoring surrounding spaces.
'   - MasterFilePath / CurrentFilePath are required; the rest have defaults.
'   - MasterKeyColumn / CurrentKeyColumn do NOT need to match each other or
'     be in the same column position — each is looked up by header name in
'     its own file (see modUtils.ReadKeyColumnValues).
'
' Why local paths, not a SharePoint URL directly:
'   Workbooks.Open on an https:// SharePoint URL is unreliable from VBA
'   (silent auth prompts, automation errors) — unlike Power Query's
'   SharePoint.Files(), which authenticates properly. Point MasterFilePath /
'   CurrentFilePath at the local OneDrive/SharePoint-synced folder instead
'   (e.g. C:\Users\you\OneDrive - Company\Library\Master.xlsx).
' ============================================================================
Option Explicit

Public Type CompareConfig
    MasterFilePath As String
    CurrentFilePath As String
    SheetName As String
    MasterKeyColumn As String
    CurrentKeyColumn As String
End Type

Public Function ReadCompareConfig() As CompareConfig
    Dim cfg As CompareConfig
    Dim lo As ListObject
    Dim keyCol As Long
    Dim valCol As Long
    Dim data As Variant
    Dim r As Long
    Dim settingName As String
    Dim settingValue As Variant

    ' Defaults — everything except MasterFilePath/CurrentFilePath can be omitted.
    cfg.SheetName = "Sheet1"
    cfg.MasterKeyColumn = "ID"
    cfg.CurrentKeyColumn = "ID"

    Set lo = FindListObject("tblConfig")
    If lo Is Nothing Then
        Err.Raise vbObjectError + 601, "ReadCompareConfig", _
            "ไม่พบตาราง tblConfig ใน workbook นี้ — สร้าง sheet ""Config"" " & _
            "แล้วทำตารางคอลัมน์ Key/Value (Ctrl+T ตั้งชื่อ tblConfig)"
    End If

    keyCol = lo.ListColumns("Key").Index
    valCol = lo.ListColumns("Value").Index

    If Not lo.DataBodyRange Is Nothing Then
        data = lo.DataBodyRange.Value
        For r = 1 To UBound(data, 1)
            settingName = LCase$(Trim$(CStr(data(r, keyCol) & vbNullString)))
            settingValue = data(r, valCol)
            If Not IsBlankValue(settingValue) Then
                Select Case settingName
                    Case "masterfilepath":   cfg.MasterFilePath = Trim$(CStr(settingValue))
                    Case "currentfilepath":  cfg.CurrentFilePath = Trim$(CStr(settingValue))
                    Case "sheetname":        cfg.SheetName = Trim$(CStr(settingValue))
                    Case "masterkeycolumn":  cfg.MasterKeyColumn = Trim$(CStr(settingValue))
                    Case "currentkeycolumn": cfg.CurrentKeyColumn = Trim$(CStr(settingValue))
                End Select
            End If
        Next r
    End If

    If cfg.MasterFilePath = vbNullString Then
        Err.Raise vbObjectError + 602, "ReadCompareConfig", _
            "tblConfig ยังไม่ได้ตั้งค่า MasterFilePath — ใส่ path ไฟล์ Master ก่อน"
    End If
    If cfg.CurrentFilePath = vbNullString Then
        Err.Raise vbObjectError + 603, "ReadCompareConfig", _
            "tblConfig ยังไม่ได้ตั้งค่า CurrentFilePath — ใส่ path ไฟล์ Current ก่อน"
    End If

    cfg.MasterFilePath = StripQuotes(cfg.MasterFilePath)
    cfg.CurrentFilePath = StripQuotes(cfg.CurrentFilePath)
    ValidateLocalPath cfg.MasterFilePath, "MasterFilePath"
    ValidateLocalPath cfg.CurrentFilePath, "CurrentFilePath"

    ReadCompareConfig = cfg
End Function

Private Function StripQuotes(ByVal p As String) As String
    p = Trim$(p)
    Do While Len(p) >= 2 And Left$(p, 1) = """" And Right$(p, 1) = """"
        p = Trim$(Mid$(p, 2, Len(p) - 2))
    Loop
    StripQuotes = p
End Function

' Fail with a readable message BEFORE Workbooks.Open turns an https path
' into a cryptic automation error or a silent credential prompt.
Private Sub ValidateLocalPath(ByVal p As String, ByVal settingName As String)
    If LCase$(Left$(p, 4)) = "http" Then
        Err.Raise vbObjectError + 604, "ValidateLocalPath", _
            settingName & " เป็นลิงก์เว็บ (" & p & ") — VBA เปิดผ่าน https ไม่เสถียร " & _
            "ต้องใช้ path ในเครื่อง เช่นโฟลเดอร์ SharePoint/OneDrive ที่ sync ไว้ " & _
            "(C:\Users\...\OneDrive - บริษัท\...)"
    End If
End Sub
