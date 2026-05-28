Attribute VB_Name = "M0_Helpers"
Option Explicit

' ============================================================
' M0_Helpers - Utility functions dung chung
' ============================================================

Public Const SHEET_SKU_MASTER As String = "SKU_Master"
Public Const SHEET_PARAMETERS As String = "Parameters"
Public Const SHEET_ROI_OUTPUT As String = "ROI_Output"
Public Const SHEET_CANDIDATES As String = "Top_Candidates"
Public Const SHEET_DASHBOARD As String = "Dashboard"
Public Const SHEET_LOG As String = "Log"

Public Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function

Public Function LastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    LastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Sub WriteHeader(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(1, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 121)
            .HorizontalAlignment = xlLeft
        End With
    Next i
    ws.Rows(1).RowHeight = 22
    ws.Range(ws.Cells(1, 1), ws.Cells(1, UBound(headers) + 1)).AutoFilter
End Sub

Public Sub LogInfo(ByVal level As String, ByVal source As String, ByVal msg As String)
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_LOG)
    Dim nextRow As Long
    nextRow = LastRow(ws, 1) + 1
    If nextRow = 2 And ws.Cells(1, 1).Value = "" Then
        WriteHeader ws, Array("Timestamp", "Level", "Source", "Message")
        nextRow = 2
    End If
    ws.Cells(nextRow, 1).Value = Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = level
    ws.Cells(nextRow, 3).Value = source
    ws.Cells(nextRow, 4).Value = msg
End Sub

Public Sub AutoFitAllColumns(ByVal ws As Worksheet)
    ws.Cells.EntireColumn.AutoFit
End Sub

Public Sub FormatColumnAsCurrency(ByVal ws As Worksheet, ByVal col As Long)
    Dim lastR As Long
    lastR = LastRow(ws, col)
    If lastR < 2 Then Exit Sub
    ws.Range(ws.Cells(2, col), ws.Cells(lastR, col)).NumberFormat = "$#,##0"
End Sub

' --- Read parameter value from Parameters sheet by key name ---
Public Function GetParameter(ByVal key As String, ByVal defaultVal As Double) As Double
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_PARAMETERS)
    On Error GoTo 0
    If ws Is Nothing Then
        GetParameter = defaultVal
        Exit Function
    End If
    Dim lastR As Long: lastR = LastRow(ws, 1)
    Dim r As Long
    For r = 2 To lastR
        If CStr(ws.Cells(r, 1).Value) = key Then
            GetParameter = CDbl(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    GetParameter = defaultVal
End Function
