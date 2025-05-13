Option Explicit

Public Sub CreatePivotTables()
    On Error GoTo ErrHandler
    Dim wsDash As Worksheet, wsData As Worksheet
    Dim pc As PivotCache, pt As PivotTable
    Dim tbl As ListObject, pvtName As String
    Dim startCell As Range

    Set wsDash = ThisWorkbook.Worksheets("Dashboard_Pivot") '데이터 모델용 숨김 시트
    wsDash.Cells.Clear

    '▼ ① TrendByMonth
    Set tbl = ThisWorkbook.Worksheets("STEP_FLOW").ListObjects("tblStepFlow")
    pvtName = "pvTrend"
    Set pc = ThisWorkbook.PivotCaches.Create(xlDatabase, tbl.Range.Address(True, True, xlR1C1, True))
    Set pt = pc.CreatePivotTable(wsDash.Range("A3"), pvtName)

    With pt
        .PivotFields("ATA_Month").Orientation = xlRowField
        .AddDataField .PivotFields("NO."), "CNT", xlCount
        .AddDataField .PivotFields("전체 리드타임"), "AVG LT", xlAverage
    End With
    pt.PivotFields("ATA_Month").AutoGroup

    '▼ ②~⑧ 나머지 피벗 동일 패턴… (간략화)
    '   STEP_NAME · SLA · Container · Forecast · Matrix 등 추가 구현
ExitSub:
    Exit Sub
ErrHandler:
    mUtils.LogError Err.Number, Err.Description, "CreatePivotTables"
    Resume ExitSub
End Sub