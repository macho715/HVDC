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
        Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "" & fileName & ".pdf", _
        Quality:=xlQualityStandard
End Sub