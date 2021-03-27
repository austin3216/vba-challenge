Sub WorksheetLoop()

Dim WS_Count As Integer
Dim I As Integer

' Set WS_Count equal to the number of worksheets in the active workbook
WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop
For I = 1 To WS_Count

    ' Activate stock ticker
    ActiveWorkbook.Worksheets(I).Activate
  
    StockTicker

Next I

MsgBox ("completed")

End Sub