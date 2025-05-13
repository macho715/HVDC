Option Explicit
Sub GenerateCharts()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("대시보드")
    Dim co As ChartObject, pt As PivotTable

    '-------------- ChartBox1 : 월별 처리량&평균 LT --------------
    Set pt = ThisWorkbook.Worksheets("Dashboard_Pivot").PivotTables("pvTrend")
    Set co = ws.ChartObjects("ChartBox1")
    With co.Chart
        .SetSourceData pt.TableRange1
        .ChartType = xlColumnClustered
        .SeriesCollection(1).ChartType = xlLine
        .SeriesCollection(1).AxisGroup = xlSecondary
        .HasTitle = True: .ChartTitle.Text = "월별 처리량 & 평균 LT"
        .Axes(xlCategory).HasTitle = True: .Axes(xlCategory).AxisTitle.Text = "Month"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "건수"
    End With

    '-------------- ChartBox2 : 공정별 평균 LT --------------
    Set pt = ThisWorkbook.Worksheets("Dashboard_Pivot").PivotTables("pvStep")
    Set co = ws.ChartObjects("ChartBox2")
    With co.Chart
        .SetSourceData pt.TableRange1
        .ChartType = xlBarClustered
        .HasTitle = True: .ChartTitle.Text = "공정별 평균 리드타임"
        .Axes(xlValue).HasTitle = True: .Axes(xlValue).AxisTitle.Text = "일"
        .SeriesCollection(1).ApplyDataLabels
    End With
    '-------------- 차트 3~8 동일 패턴 --------------
ExitSub: Exit Sub
ErrHandler: mUtils.LogError Err.Number, Err.Description, "GenerateCharts": Resume ExitSub
End Sub