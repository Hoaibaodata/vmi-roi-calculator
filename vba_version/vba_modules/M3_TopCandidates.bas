Attribute VB_Name = "M3_TopCandidates"
Option Explicit

' ============================================================
' M3_TopCandidates - Filter + rank top VMI candidates
' ============================================================

Public Sub BuildTopCandidatesList()
    Dim wsROI As Worksheet, wsCand As Worksheet
    On Error Resume Next
    Set wsROI = ThisWorkbook.Sheets(SHEET_ROI_OUTPUT)
    On Error GoTo 0
    If wsROI Is Nothing Then
        MsgBox "Chua tinh ROI. Run 'Compute ROI' truoc.", vbExclamation
        Exit Sub
    End If

    Set wsCand = GetOrCreateSheet(SHEET_CANDIDATES)
    wsCand.Cells.Clear

    Dim headers As Variant
    headers = Array( _
        "rank", "sku_id", "category", "abc_xyz", _
        "annual_revenue", "lead_time_days", "annual_savings", _
        "payback_months", "recommendation", "candidacy_score" _
    )
    WriteHeader wsCand, headers

    Dim lastR As Long: lastR = LastRow(wsROI, 1)
    Dim r As Long, rank As Long: rank = 1
    Dim outRow As Long: outRow = 2

    ' Build temp data to sort
    Dim tempData As Object
    Set tempData = CreateObject("Scripting.Dictionary")

    For r = 2 To lastR
        Dim recommendation As String: recommendation = CStr(wsROI.Cells(r, 11).Value)
        If recommendation = "NOT_ELIGIBLE" Or recommendation = "NOT_ELIGIBLE_TOO_ERRATIC" Then GoTo NextR

        Dim savings As Double: savings = CDbl(wsROI.Cells(r, 9).Value)
        If savings <= 0 Then GoTo NextR

        Dim paybackStr As Variant: paybackStr = wsROI.Cells(r, 10).Value
        Dim payback As Double
        If IsNumeric(paybackStr) Then
            payback = CDbl(paybackStr)
        Else
            payback = 9999
        End If
        If payback > 24 Then GoTo NextR

        Dim abcClass As String, xyzClass As String
        Dim abcXyz As String: abcXyz = CStr(wsROI.Cells(r, 3).Value)
        abcClass = Left(abcXyz, 1)

        ' Candidacy score: savings/1000 - payback*10 + A_class bonus
        Dim cScore As Double
        cScore = savings / 1000 - payback * 10
        If abcClass = "A" Then cScore = cScore + 300
        If abcClass = "B" Then cScore = cScore + 100

        ' Store in dict using cScore as key (formatted for sorting)
        Dim entryKey As String
        entryKey = Format(cScore, "0000000.000") & "_" & CStr(wsROI.Cells(r, 1).Value)
        tempData.Add entryKey, r
NextR:
    Next r

    ' Get keys sorted descending
    Dim sortedKeys As Variant
    sortedKeys = SortDictionaryKeysDesc(tempData)

    Dim i As Long
    For i = LBound(sortedKeys) To UBound(sortedKeys)
        If i >= 20 Then Exit For  ' top 20
        Dim sourceRow As Long: sourceRow = tempData(sortedKeys(i))

        With wsCand
            .Cells(outRow, 1).Value = rank
            .Cells(outRow, 2).Value = wsROI.Cells(sourceRow, 1).Value
            .Cells(outRow, 3).Value = wsROI.Cells(sourceRow, 2).Value
            .Cells(outRow, 4).Value = wsROI.Cells(sourceRow, 3).Value
            .Cells(outRow, 5).Value = wsROI.Cells(sourceRow, 5).Value
            .Cells(outRow, 6).Value = wsROI.Cells(sourceRow, 6).Value
            .Cells(outRow, 7).Value = wsROI.Cells(sourceRow, 9).Value
            .Cells(outRow, 8).Value = wsROI.Cells(sourceRow, 10).Value
            .Cells(outRow, 9).Value = wsROI.Cells(sourceRow, 11).Value
            ' Re-compute candidacy for display
            Dim s As Double: s = CDbl(.Cells(outRow, 7).Value)
            Dim p As Double: p = CDbl(.Cells(outRow, 8).Value)
            Dim cs As Double: cs = s / 1000 - p * 10
            If Left(CStr(.Cells(outRow, 4).Value), 1) = "A" Then cs = cs + 300
            If Left(CStr(.Cells(outRow, 4).Value), 1) = "B" Then cs = cs + 100
            .Cells(outRow, 10).Value = Round(cs, 1)
        End With

        rank = rank + 1
        outRow = outRow + 1
    Next i

    FormatColumnAsCurrency wsCand, 5
    FormatColumnAsCurrency wsCand, 7
    AutoFitAllColumns wsCand

    LogInfo "INFO", "M3.BuildTopCandidatesList", (outRow - 2) & " top candidates"
    MsgBox "Top " & (outRow - 2) & " candidates ranked. Xem sheet '" & SHEET_CANDIDATES & "'.", vbInformation
End Sub

Private Function SortDictionaryKeysDesc(ByVal dict As Object) As Variant
    Dim keys As Variant
    keys = dict.Keys
    ' Bubble sort descending — OK for <500 items
    Dim n As Long: n = UBound(keys)
    Dim i As Long, j As Long, temp As Variant
    For i = 0 To n - 1
        For j = 0 To n - i - 1
            If StrComp(CStr(keys(j)), CStr(keys(j + 1))) < 0 Then
                temp = keys(j)
                keys(j) = keys(j + 1)
                keys(j + 1) = temp
            End If
        Next j
    Next i
    SortDictionaryKeysDesc = keys
End Function
