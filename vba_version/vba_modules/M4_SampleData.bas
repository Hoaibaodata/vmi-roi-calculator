Attribute VB_Name = "M4_SampleData"
Option Explicit

' ============================================================
' M4_SampleData - Generate 50 sample SKUs
' ============================================================

Public Sub GenerateSampleSKUMaster()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(SHEET_SKU_MASTER)
    ws.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "sku_id", "category", "abc_class", "xyz_class", "abc_xyz", _
        "avg_weekly_demand", "sigma_weekly_demand", "unit_price", _
        "lead_time_days", "sigma_lead_time_days", "is_critical" _
    )
    WriteHeader ws, headers

    Dim categories As Variant: categories = Array("Electronics", "Mechanical", "Plastics", "Chemicals", "Packaging")
    Dim abcClasses As Variant: abcClasses = Array("A", "B", "C")
    Dim abcWeights As Variant: abcWeights = Array(0.1, 0.3, 0.6)  ' cumulative
    Dim xyzClasses As Variant: xyzClasses = Array("X", "Y", "Z")
    Dim xyzWeights As Variant: xyzWeights = Array(0.4, 0.4, 0.2)

    Randomize 42

    Dim i As Long
    Dim numSKUs As Long: numSKUs = 50

    For i = 1 To numSKUs
        Dim r As Double: r = Rnd()
        Dim abc As String
        If r < abcWeights(0) Then
            abc = abcClasses(0)
        ElseIf r < abcWeights(0) + abcWeights(1) Then
            abc = abcClasses(1)
        Else
            abc = abcClasses(2)
        End If

        r = Rnd()
        Dim xyz As String
        If r < xyzWeights(0) Then
            xyz = xyzClasses(0)
        ElseIf r < xyzWeights(0) + xyzWeights(1) Then
            xyz = xyzClasses(1)
        Else
            xyz = xyzClasses(2)
        End If

        Dim avgDemand As Double, unitPrice As Double
        Select Case abc
            Case "A"
                avgDemand = 500 + Rnd() * 4500
                unitPrice = 10 + Rnd() * 90
            Case "B"
                avgDemand = 100 + Rnd() * 400
                unitPrice = 5 + Rnd() * 45
            Case "C"
                avgDemand = 5 + Rnd() * 95
                unitPrice = 1 + Rnd() * 19
        End Select

        Dim cv As Double
        Select Case xyz
            Case "X": cv = 0.1 + Rnd() * 0.3
            Case "Y": cv = 0.4 + Rnd() * 0.4
            Case "Z": cv = 0.8 + Rnd() * 1.2
        End Select
        Dim sigmaDemand As Double: sigmaDemand = avgDemand * cv

        Dim ltOptions As Variant: ltOptions = Array(7, 14, 21, 28, 35)
        Dim ltDays As Long: ltDays = ltOptions(Int(Rnd() * 5))
        Dim sigmaLT As Double: sigmaLT = ltDays * (0.1 + Rnd() * 0.2)

        Dim isCritical As Boolean
        isCritical = (abc = "A" And avgDemand > 1000)

        Dim cat As String: cat = categories(Int(Rnd() * 5))

        With ws
            .Cells(i + 1, 1).Value = "SKU" & Format(i, "00000")
            .Cells(i + 1, 2).Value = cat
            .Cells(i + 1, 3).Value = abc
            .Cells(i + 1, 4).Value = xyz
            .Cells(i + 1, 5).Value = abc & xyz
            .Cells(i + 1, 6).Value = Round(avgDemand, 1)
            .Cells(i + 1, 7).Value = Round(sigmaDemand, 1)
            .Cells(i + 1, 8).Value = Round(unitPrice, 2)
            .Cells(i + 1, 9).Value = ltDays
            .Cells(i + 1, 10).Value = Round(sigmaLT, 2)
            .Cells(i + 1, 11).Value = isCritical
        End With
    Next i

    ws.Range("H2:H" & (numSKUs + 1)).NumberFormat = "$0.00"
    AutoFitAllColumns ws

    LogInfo "INFO", "M4.GenerateSampleSKUMaster", numSKUs & " sample SKUs generated"
    MsgBox "Da tao " & numSKUs & " SKU mau. Bay gio bam 'Compute ROI'.", vbInformation
End Sub
