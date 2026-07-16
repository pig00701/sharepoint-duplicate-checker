Attribute VB_Name = "modCompare"
' ============================================================================
' modCompare.bas — entrypoint: RunCompareFiles
' ============================================================================
' Port of the Power Query CompareFiles.pq logic to VBA. Opens Master and
' Current **read-only**, never saves them, and closes both before writing
' anything — the source files are never modified, same guarantee as the
' Power Query version (see README "สรุปพฤติกรรม: read-only เสมอ").
'
' Output: sheet "Report" in THIS workbook, table "ReportTable" with columns
' Key / Status ("New" / "Missing" / "Duplicate (Nx)"). Rows where the key
' matches in both files are not listed (nothing to report).
' ============================================================================
Option Explicit

Public Sub RunCompareFiles()
    Const PROC_NAME As String = "RunCompareFiles"
    Dim cfg As CompareConfig
    Dim masterWb As Workbook
    Dim currentWb As Workbook
    Dim masterKeys As Collection
    Dim currentKeys As Collection
    Dim outKey As New Collection
    Dim outStatus As New Collection
    Dim dupCount As Long, newCount As Long, missingCount As Long

    On Error GoTo ErrHandler

    cfg = ReadCompareConfig()

    Application.ScreenUpdating = False

    Set masterWb = Workbooks.Open(cfg.MasterFilePath, UpdateLinks:=0, ReadOnly:=True)
    Set currentWb = Workbooks.Open(cfg.CurrentFilePath, UpdateLinks:=0, ReadOnly:=True)

    Set masterKeys = ReadKeyColumnValues(masterWb, cfg.SheetName, cfg.MasterKeyColumn)
    Set currentKeys = ReadKeyColumnValues(currentWb, cfg.SheetName, cfg.CurrentKeyColumn)

    masterWb.Close SaveChanges:=False
    Set masterWb = Nothing
    currentWb.Close SaveChanges:=False
    Set currentWb = Nothing

    BuildReport masterKeys, currentKeys, outKey, outStatus, dupCount, newCount, missingCount
    WriteReport outKey, outStatus

    Application.ScreenUpdating = True

    MsgBox "เปรียบเทียบเสร็จแล้ว — New " & newCount & ", Missing " & missingCount & _
           ", Duplicate " & dupCount & " (ดูรายละเอียดที่ sheet ""Report"")", _
           vbInformation, "SharePoint Duplicate Checker"
    Exit Sub

ErrHandler:
    Dim errNumber As Long, errDescription As String
    errNumber = Err.Number: errDescription = Err.Description
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not masterWb Is Nothing Then masterWb.Close SaveChanges:=False
    If Not currentWb Is Nothing Then currentWb.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "เกิดข้อผิดพลาดใน " & PROC_NAME & " (" & errNumber & "): " & errDescription, _
           vbCritical, "SharePoint Duplicate Checker"
End Sub

' Duplicate: key appears more than once in Current.
' New: key in Current, not in Master.
' Missing: key in Master, not in Current.
Private Sub BuildReport(ByVal masterKeys As Collection, ByVal currentKeys As Collection, _
                         ByRef outKey As Collection, ByRef outStatus As Collection, _
                         ByRef dupCount As Long, ByRef newCount As Long, ByRef missingCount As Long)
    Dim currentCounts As Object
    Dim masterSet As Object
    Dim currentSet As Object
    Dim v As Variant
    Dim k As Variant

    Set currentCounts = CreateObject("Scripting.Dictionary")
    Set masterSet = CreateObject("Scripting.Dictionary")
    Set currentSet = CreateObject("Scripting.Dictionary")

    For Each v In currentKeys
        If currentCounts.Exists(v) Then
            currentCounts(v) = currentCounts(v) + 1
        Else
            currentCounts.Add v, 1
        End If
        If Not currentSet.Exists(v) Then currentSet.Add v, True
    Next v

    For Each v In masterKeys
        If Not masterSet.Exists(v) Then masterSet.Add v, True
    Next v

    dupCount = 0: newCount = 0: missingCount = 0

    For Each k In currentCounts.Keys
        If currentCounts(k) > 1 Then
            outKey.Add k
            outStatus.Add "Duplicate (" & currentCounts(k) & "x)"
            dupCount = dupCount + 1
        End If
    Next k

    For Each k In currentSet.Keys
        If Not masterSet.Exists(k) Then
            outKey.Add k
            outStatus.Add "New"
            newCount = newCount + 1
        End If
    Next k

    For Each k In masterSet.Keys
        If Not currentSet.Exists(k) Then
            outKey.Add k
            outStatus.Add "Missing"
            missingCount = missingCount + 1
        End If
    Next k
End Sub

Private Sub WriteReport(ByVal outKey As Collection, ByVal outStatus As Collection)
    Dim ws As Worksheet
    Dim i As Long
    Dim lo As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Report")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Report"
    Else
        On Error Resume Next
        ws.ListObjects("ReportTable").Delete
        On Error GoTo 0
        ws.Cells.Clear
    End If

    ws.Range("A1").Value = "Key"
    ws.Range("B1").Value = "Status"

    For i = 1 To outKey.Count
        ws.Cells(i + 1, 1).Value = outKey(i)
        ws.Cells(i + 1, 2).Value = outStatus(i)
    Next i

    If outKey.Count > 0 Then
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:B" & outKey.Count + 1), , xlYes)
        lo.Name = "ReportTable"
    End If
End Sub
