Attribute VB_Name = "M5_Dashboard"
Option Explicit

' ============================================================
' M5_Dashboard - Action buttons
' ============================================================

Public Sub Action_RunFullAnalysis()
    Dim wsMaster As Worksheet
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets(SHEET_SKU_MASTER)
    On Error GoTo 0
    If wsMaster Is Nothing Or LastRow(wsMaster, 1) < 2 Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Chua co SKU_Master data. Generate 50 SKU mau?", vbYesNo + vbQuestion, "Setup")
        If ans = vbYes Then
            Call GenerateSampleSKUMaster
        Else
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Call ComputeROIForAllSKUs    ' M2
    Call BuildTopCandidatesList  ' M3
    Application.ScreenUpdating = True

    MsgBox "Analysis xong! Xem ROI_Output + Top_Candidates.", vbInformation
End Sub

Public Sub Action_GenerateSample()
    Call GenerateSampleSKUMaster
End Sub

Public Sub Action_ComputeROI()
    Call ComputeROIForAllSKUs
End Sub

Public Sub Action_TopCandidates()
    Call BuildTopCandidatesList
End Sub

Public Sub Action_ResetOutputs()
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Xoa ROI_Output + Top_Candidates?", vbYesNo + vbQuestion, "Reset")
    If ans <> vbYes Then Exit Sub
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(SHEET_ROI_OUTPUT).Cells.Clear
    ThisWorkbook.Sheets(SHEET_CANDIDATES).Cells.Clear
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "Reset xong.", vbInformation
End Sub
