Attribute VB_Name = "M2_ROIEngine"
Option Explicit

' ============================================================
' M2_ROIEngine - Loop SKU master, compute ROI per SKU, write ROI_Output
' ============================================================

Public Sub ComputeROIForAllSKUs()
    Dim wsMaster As Worksheet, wsOut As Worksheet
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Sheets(SHEET_SKU_MASTER)
    On Error GoTo 0
    If wsMaster Is Nothing Then
        MsgBox "Khong tim thay SKU_Master. Generate sample truoc.", vbExclamation
        Exit Sub
    End If
    Dim lastR As Long: lastR = LastRow(wsMaster, 1)
    If lastR < 2 Then
        MsgBox "SKU_Master trong. Generate sample truoc.", vbExclamation
        Exit Sub
    End If

    Set wsOut = GetOrCreateSheet(SHEET_ROI_OUTPUT)
    wsOut.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "sku_id", "category", "abc_xyz", "annual_volume", "annual_revenue", _
        "lead_time_days", "status_quo_cost", "vmi_cost", _
        "annual_savings", "payback_months", "recommendation" _
    )
    WriteHeader wsOut, headers

    ' Read parameters
    Dim setupCost As Double: setupCost = GetParameter("vmi_setup_cost", 5000)

    Dim r As Long
    Dim outRow As Long: outRow = 2

    For r = 2 To lastR
        Dim sku As String: sku = CStr(wsMaster.Cells(r, 1).Value)
        Dim cat As String: cat = CStr(wsMaster.Cells(r, 2).Value)
        Dim abcXyz As String: abcXyz = CStr(wsMaster.Cells(r, 5).Value)
        Dim avgDemand As Double: avgDemand = CDbl(wsMaster.Cells(r, 6).Value)
        Dim sigmaDemand As Double: sigmaDemand = CDbl(wsMaster.Cells(r, 7).Value)
        Dim unitPrice As Double: unitPrice = CDbl(wsMaster.Cells(r, 8).Value)
        Dim ltDays As Double: ltDays = CDbl(wsMaster.Cells(r, 9).Value)
        Dim sigmaLTDays As Double: sigmaLTDays = CDbl(wsMaster.Cells(r, 10).Value)
        Dim annualVol As Double: annualVol = avgDemand * 52
        Dim annualRev As Double: annualRev = annualVol * unitPrice

        Dim sqCost As Double, vmiC As Double, savings As Double
        sqCost = StatusQuoCost(avgDemand, sigmaDemand, unitPrice, ltDays, sigmaLTDays)
        vmiC = VMICost(avgDemand, sigmaDemand, unitPrice, sigmaLTDays)
        savings = sqCost - vmiC

        Dim payback As Double
        payback = VMIPaybackMonths(setupCost, savings)

        Dim recommendation As String
        recommendation = VMIRecommendation(payback, abcXyz)

        With wsOut
            .Cells(outRow, 1).Value = sku
            .Cells(outRow, 2).Value = cat
            .Cells(outRow, 3).Value = abcXyz
            .Cells(outRow, 4).Value = Round(annualVol, 0)
            .Cells(outRow, 5).Value = Round(annualRev, 0)
            .Cells(outRow, 6).Value = ltDays
            .Cells(outRow, 7).Value = Round(sqCost, 0)
            .Cells(outRow, 8).Value = Round(vmiC, 0)
            .Cells(outRow, 9).Value = Round(savings, 0)
            .Cells(outRow, 10).Value = IIf(payback >= 9999, "n/a", Round(payback, 1))
            .Cells(outRow, 11).Value = recommendation
        End With

        outRow = outRow + 1
    Next r

    ' Format
    FormatColumnAsCurrency wsOut, 5
    FormatColumnAsCurrency wsOut, 7
    FormatColumnAsCurrency wsOut, 8
    FormatColumnAsCurrency wsOut, 9

    ' Sort by savings desc
    wsOut.Range("A2:K" & (outRow - 1)).Sort _
        Key1:=wsOut.Range("I2"), Order1:=xlDescending, Header:=xlNo

    AutoFitAllColumns wsOut
    Call ApplyROIFormatting(wsOut, outRow - 1)

    LogInfo "INFO", "M2.ComputeROIForAllSKUs", (outRow - 2) & " SKUs processed"
    MsgBox "Da tinh ROI cho " & (outRow - 2) & " SKUs. Xem sheet '" & SHEET_ROI_OUTPUT & "'.", vbInformation
End Sub

Private Sub ApplyROIFormatting(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim r As Long
    For r = 2 To lastRow
        Dim savings As Double
        savings = CDbl(ws.Cells(r, 9).Value)
        If savings > 0 Then
            ws.Cells(r, 9).Interior.Color = RGB(56, 142, 60)
            ws.Cells(r, 9).Font.Color = RGB(255, 255, 255)
        Else
            ws.Cells(r, 9).Interior.Color = RGB(211, 47, 47)
            ws.Cells(r, 9).Font.Color = RGB(255, 255, 255)
        End If

        Dim rec As String: rec = CStr(ws.Cells(r, 11).Value)
        Select Case rec
            Case "STRONG_VMI_CANDIDATE"
                ws.Cells(r, 11).Interior.Color = RGB(56, 142, 60)
                ws.Cells(r, 11).Font.Color = RGB(255, 255, 255)
            Case "MODERATE_VMI_CANDIDATE"
                ws.Cells(r, 11).Interior.Color = RGB(25, 118, 210)
                ws.Cells(r, 11).Font.Color = RGB(255, 255, 255)
            Case "WEAK_VMI_CANDIDATE"
                ws.Cells(r, 11).Interior.Color = RGB(245, 124, 0)
                ws.Cells(r, 11).Font.Color = RGB(255, 255, 255)
        End Select
        ws.Cells(r, 11).Font.Bold = True
    Next r
End Sub
