Attribute VB_Name = "mPython"
Option Explicit

Sub Call_ARIMA()
    On Error GoTo ErrHandler
    RunPython "import fcast; fcast.run_forecast(r'" & ThisWorkbook.FullName & "')"
ExitSub:
    Exit Sub
ErrHandler:
    mUtils.LogError Err.Number, Err.Description, "Call_ARIMA"
    Resume ExitSub
End Sub