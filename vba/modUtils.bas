Attribute VB_Name = "modUtils"
' ============================================================================
' modUtils.bas — helpers shared by modConfig / modCompare
' ============================================================================
Option Explicit

Public Function FindListObject(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set FindListObject = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Trim$(v) = vbNullString)
    Else
        IsBlankValue = False
    End If
End Function

' Reads every non-blank value under the header cell that matches
' columnHeader (row 1, case/space-insensitive) on the given sheet of wb.
' Values are normalized with CStr/Trim so numeric and text IDs compare
' consistently — note this means 1 and "1" match, but "01" and 1 do not.
Public Function ReadKeyColumnValues(ByVal wb As Workbook, ByVal sheetName As String, _
                                     ByVal columnHeader As String) As Collection
    Dim ws As Worksheet
    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long
    Dim colIdx As Long
    Dim r As Long
    Dim v As Variant
    Dim result As New Collection

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 611, "ReadKeyColumnValues", _
            "ไม่พบ sheet '" & sheetName & "' ในไฟล์ " & wb.Name
    End If

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    colIdx = 0
    For c = 1 To lastCol
        If Trim$(LCase$(CStr(ws.Cells(1, c).Value & vbNullString))) = _
           Trim$(LCase$(columnHeader)) Then
            colIdx = c
            Exit For
        End If
    Next c
    If colIdx = 0 Then
        Err.Raise vbObjectError + 612, "ReadKeyColumnValues", _
            "ไม่พบคอลัมน์ '" & columnHeader & "' ใน sheet '" & sheetName & "' ของไฟล์ " & wb.Name
    End If

    lastRow = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).Row
    For r = 2 To lastRow
        v = ws.Cells(r, colIdx).Value
        If Not IsBlankValue(v) Then result.Add Trim$(CStr(v))
    Next r

    Set ReadKeyColumnValues = result
End Function
