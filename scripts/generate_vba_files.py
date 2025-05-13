# scripts/generate_vba_files.py

vba_modules = {
    "mInit.bas": '''
Option Explicit

Sub Auto_Open()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    ' 1) 데이터 새로 고침
    ThisWorkbook.RefreshAll

    ' 2) 피벗 재생성
    Call mPivot.CreatePivotTables

    ' 3) 차트 업데이트
    Call mChart.GenerateCharts

    ' 4) 슬라이서/타임라인 재연결
    Call mSlicer.AddSlicers

    ' 5) 예측 ARIMA 호출
    Call mPython.Call_ARIMA

ExitHandler:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    mUtils.LogError Err.Number, Err.Description, "Auto_Open"
    Resume ExitHandler
End Sub
''',
    "mPivot.bas": '''
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
''',
    "mChart.bas": '''
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
''',
    "mSlicer.bas": '''
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
''',
    "mPython.bas": '''
Option Explicit
Sub Call_ARIMA()
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    RunPython ("import fcast; fcast.run_arima(r'" & wb.FullName & "')")
ExitSub: Exit Sub
ErrHandler: mUtils.LogError Err.Number, Err.Description, "Call_ARIMA": Resume ExitSub
End Sub
''',
    "mUtils.bas": '''
Option Explicit
Sub LogError(num As Long, msg As String, src As String)
    With ThisWorkbook.Worksheets("Dashboard_Log")
        Dim r As Long: r = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(r, 1).Value = Now
        .Cells(r, 2).Value = src
        .Cells(r, 3).Value = num & " - " & msg
    End With
End Sub

Sub ExportAsPDF(sheetName As String, fileName As String)
    ThisWorkbook.Worksheets(sheetName).ExportAsFixedFormat _
        Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\" & fileName & ".pdf", _
        Quality:=xlQualityStandard
End Sub
''',
    "ThisWorkbook.txt": '''
Private Sub Workbook_Open()
    mInit.Auto_Open
End Sub
''',
}

fcast_code = '''
import pandas as pd
from pathlib import Path
import xlwings as xw
from statsmodels.tsa.arima.model import ARIMA

def run_arima(xlsm_path: str):
    wb = xw.Book(xlsm_path)
    df = pd.read_excel(xlsm_path, sheet_name="STEP_FLOW", usecols=["ATA", "전체 리드타임"])
    df["ATA"] = pd.to_datetime(df["ATA"])
    ts = df.set_index("ATA")["전체 리드타임"].resample("M").mean().fillna(method="ffill")

    model = ARIMA(ts, order=(1,1,1))
    fcast = model.fit().forecast(steps=3)
    f_df = fcast.reset_index()
    f_df.columns = ["Forecast_Month", "Pred LT"]

    sht = wb.sheets["DataModel"]  # 예: 테이블 넣을 시트
    sht.range("A1").options(index=False, header=True).value = f_df

    wb.save()
'''

readme = '''
# HVDC Dashboard 자동화

## 사용법 요약

1. Excel 파일을 .xlsm로 저장
2. ALT+F11 → 각 VBA 모듈(.bas) Import
3. Power Query 테이블/관계 연결
4. Python 의존성 설치
   ```
   pip install xlwings statsmodels pandas
   xlwings addin install
   ```
5. fcast.py 파일을 같은 폴더에 저장
6. 파일 열기만 하면 자동 갱신/예측

## Troubleshooting
- 매크로 차단: 신뢰할 수 있는 위치에 저장
- xlwings not found: 가상환경 활성화 후 pip 재설치
- PQ 오류: PQ 편집기에서 열/시트명 동기화
'''

# 파일 생성
for fname, code in vba_modules.items():
    with open(fname, "w", encoding="utf-8") as f:
        f.write(code.strip())

with open("fcast.py", "w", encoding="utf-8") as f:
    f.write(fcast_code.strip())

with open("README.md", "w", encoding="utf-8") as f:
    f.write(readme.strip())

print("✅ VBA .bas, fcast.py, README.md 파일 생성 완료") 