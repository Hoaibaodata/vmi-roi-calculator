Attribute VB_Name = "M1_CostFormulas"
Option Explicit

' ============================================================
' M1_CostFormulas - UDFs for safety stock + cost calculation
' Available as Excel formulas
' ============================================================

' --- Safety stock with full variability (Z * sqrt(LT*sigma_d^2 + d^2*sigma_LT^2)) ---
Public Function SafetyStock( _
    ByVal serviceLevel As Double, _
    ByVal sigmaDemandWeekly As Double, _
    ByVal ltWeeks As Double, _
    ByVal avgDemandWeekly As Double, _
    ByVal sigmaLTWeeks As Double) As Double
    Dim z As Double, variance As Double
    z = Application.WorksheetFunction.NormSInv(serviceLevel)
    variance = ltWeeks * (sigmaDemandWeekly ^ 2) + (avgDemandWeekly ^ 2) * (sigmaLTWeeks ^ 2)
    If variance < 0 Then variance = 0
    SafetyStock = z * Sqr(variance)
End Function

' --- Status quo annual total cost ---
Public Function StatusQuoCost( _
    ByVal avgWeeklyDemand As Double, _
    ByVal sigmaDemandWeekly As Double, _
    ByVal unitPrice As Double, _
    ByVal ltDays As Double, _
    ByVal sigmaLTDays As Double, _
    Optional ByVal holdingRate As Double = 0.25, _
    Optional ByVal stockoutMultiplier As Double = 3#, _
    Optional ByVal marginRate As Double = 0.3, _
    Optional ByVal orderCost As Double = 100#, _
    Optional ByVal serviceLevel As Double = 0.95) As Double
    Dim ltWeeks As Double: ltWeeks = ltDays / 7#
    Dim sigmaLTWeeks As Double: sigmaLTWeeks = sigmaLTDays / 7#
    Dim ss As Double
    ss = SafetyStock(serviceLevel, sigmaDemandWeekly, ltWeeks, avgWeeklyDemand, sigmaLTWeeks)
    Dim cycleStock As Double: cycleStock = ltWeeks * avgWeeklyDemand / 2#
    Dim avgInventory As Double: avgInventory = ss + cycleStock
    Dim holdingCost As Double: holdingCost = avgInventory * unitPrice * holdingRate
    Dim annualDemand As Double: annualDemand = avgWeeklyDemand * 52
    Dim expectedStockout As Double
    expectedStockout = annualDemand * (1 - serviceLevel) * 0.1
    Dim stockoutCostPerUnit As Double
    stockoutCostPerUnit = unitPrice * marginRate * stockoutMultiplier
    Dim stockoutCost As Double: stockoutCost = expectedStockout * stockoutCostPerUnit
    Dim ordersPerYear As Double
    ordersPerYear = 52 / Application.Max(ltWeeks, 1)
    Dim orderingCost As Double: orderingCost = ordersPerYear * orderCost
    StatusQuoCost = holdingCost + stockoutCost + orderingCost
End Function

' --- VMI annual total cost ---
Public Function VMICost( _
    ByVal avgWeeklyDemand As Double, _
    ByVal sigmaDemandWeekly As Double, _
    ByVal unitPrice As Double, _
    ByVal sigmaLTDays As Double, _
    Optional ByVal vmiLTDays As Double = 7#, _
    Optional ByVal holdingRate As Double = 0.25, _
    Optional ByVal stockoutMultiplier As Double = 3#, _
    Optional ByVal marginRate As Double = 0.3, _
    Optional ByVal orderCost As Double = 100#, _
    Optional ByVal vmiFeeRate As Double = 0.06, _
    Optional ByVal serviceLevel As Double = 0.95) As Double
    Dim ltWeeksVMI As Double: ltWeeksVMI = vmiLTDays / 7#
    Dim sigmaLTWeeksVMI As Double: sigmaLTWeeksVMI = (sigmaLTDays / 7#) * 0.3  ' reduced under VMI
    Dim ss As Double
    ss = SafetyStock(serviceLevel, sigmaDemandWeekly, ltWeeksVMI, avgWeeklyDemand, sigmaLTWeeksVMI)
    Dim cycleStock As Double: cycleStock = ltWeeksVMI * avgWeeklyDemand / 2#
    Dim avgInventoryBuyer As Double: avgInventoryBuyer = ss + cycleStock
    Dim holdingCost As Double: holdingCost = avgInventoryBuyer * unitPrice * holdingRate
    Dim annualDemand As Double: annualDemand = avgWeeklyDemand * 52
    Dim expectedStockout As Double
    expectedStockout = annualDemand * (1 - serviceLevel) * 0.02  ' 2% of original risk
    Dim stockoutCostPerUnit As Double
    stockoutCostPerUnit = unitPrice * marginRate * stockoutMultiplier
    Dim stockoutCost As Double: stockoutCost = expectedStockout * stockoutCostPerUnit
    Dim orderingCost As Double
    orderingCost = (52 / Application.Max(ltWeeksVMI, 1)) * orderCost * 0.3  ' 70% admin reduction
    Dim vmiStock As Double: vmiStock = 4 * avgWeeklyDemand
    Dim vmiFee As Double: vmiFee = vmiStock * unitPrice * vmiFeeRate
    VMICost = holdingCost + stockoutCost + orderingCost + vmiFee
End Function

' --- Annual savings = status quo - VMI ---
Public Function VMISavings( _
    ByVal avgWeeklyDemand As Double, _
    ByVal sigmaDemandWeekly As Double, _
    ByVal unitPrice As Double, _
    ByVal ltDays As Double, _
    ByVal sigmaLTDays As Double) As Double
    VMISavings = StatusQuoCost(avgWeeklyDemand, sigmaDemandWeekly, unitPrice, ltDays, sigmaLTDays) _
               - VMICost(avgWeeklyDemand, sigmaDemandWeekly, unitPrice, sigmaLTDays)
End Function

' --- Payback months given setup cost + annual savings ---
Public Function VMIPaybackMonths( _
    ByVal setupCost As Double, _
    ByVal annualSavings As Double) As Double
    If annualSavings <= 0 Then
        VMIPaybackMonths = 9999
    Else
        VMIPaybackMonths = setupCost / (annualSavings / 12)
    End If
End Function

' --- Recommendation classification from payback ---
Public Function VMIRecommendation( _
    ByVal paybackMonths As Double, _
    ByVal abcXyz As String) As String
    ' Eligibility filter
    If paybackMonths > 24 Or paybackMonths <= 0 Then
        VMIRecommendation = "NOT_ELIGIBLE"
        Exit Function
    End If
    Dim xyzClass As String
    xyzClass = Right(abcXyz, 1)
    If xyzClass = "Z" Then
        VMIRecommendation = "NOT_ELIGIBLE_TOO_ERRATIC"
        Exit Function
    End If
    If paybackMonths <= 12 Then
        VMIRecommendation = "STRONG_VMI_CANDIDATE"
    ElseIf paybackMonths <= 18 Then
        VMIRecommendation = "MODERATE_VMI_CANDIDATE"
    Else
        VMIRecommendation = "WEAK_VMI_CANDIDATE"
    End If
End Function
