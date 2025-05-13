Option Explicit
Sub AddSlicers()
    Dim pt As PivotTable
    Set pt = ThisWorkbook.Worksheets("Dashboard_Pivot").PivotTables("pvTrend")

    '타임라인 (ATA_Month)
    On Error Resume Next
    ThisWorkbook.SlicerCaches("Timeline_ATA").Delete
    On Error GoTo 0
    Dim sc As SlicerCache: Set sc = ThisWorkbook.SlicerCaches.Add2(pt, "ATA_Month", "Timeline_ATA")
    sc.TimelineState.Active = False  '전체 기간

    'STEP, SITE 슬라이서
    Dim arrFld: arrFld = Array("STEP_NAME", "SITE")
    Dim i As Long
    For i = LBound(arrFld) To UBound(arrFld)
        On Error Resume Next
        ThisWorkbook.SlicerCaches(arrFld(i) & "_SL").Delete
        On Error GoTo 0
        Set sc = ThisWorkbook.SlicerCaches.Add2(pt, arrFld(i), arrFld(i) & "_SL")
        sc.SlicerItems(1).Selected = True 'All
    Next i
End Sub