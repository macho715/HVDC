'===== mInit.bas =====
Option Explicit

Sub Auto_Open()
    '1) Power Query 새로 고침
    ThisWorkbook.RefreshAll
    
    '2) 피벗 재생성
    mPivot.CreatePivotTables
    
    '3) 차트 업데이트
    mChart.GenerateCharts
    
    '4) 슬라이서 연결
    mSlicer.AddSlicers
    
    '5) (선택) Python 예측
    'mPython.Call_ARIMA
End Sub