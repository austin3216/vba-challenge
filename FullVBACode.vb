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

Sub StockTicker()

        ' Variable to hold ticker
        Dim Ticker As String
        
        ' Variable to hold volume per ticker
        Dim Volume As Double
        Volume = 0
        
        ' Variable to hold open value at start of year
        Dim openv As Double
        openv = Cells(2, 3).Value

        ' Variable to hold close value at end of year
        Dim closev As Double
        closev = 0

        ' Variable to hold yearly change bwtn close/open
        Dim change As Double
        change = 0
        
        ' Variable to hold percent change btwn close/open
        Dim percent As Double
        percent = 0
        
        ' Keep track of location for each ticker in summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        ' Find Last Row
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).row

        ' Loop through all tickers/rows
        Dim row As Long
            For row = 2 To LastRow
                   
                ' Check if within the credit ticker, if it is not...
                If Cells(row, 1).Value <> Cells(row + 1, 1).Value Then
                
                ' Set the Ticker
                Ticker = Cells(row, 1).Value
                
                ' Get close_value
                closev = Cells(row, 6).Value
                
                ' Calculate yearly change
                change = closev - openv
                
                If openv = 0 Then
                percent = 0
                 
                Else
                
                ' Calulate percent change
                percent = (closev - openv) / openv
                
                End If
                
                ' Get open_value
                openv = Cells(row + 1, 3).Value
                
                ' Add to the Volume Total
                Volume = Volume + Cells(row, 7).Value
                
                ' Create Summary Table
                Range("J1").Value = "Ticker"
                Range("K1").Value = "Yearly_Change"
                Range("L1").Value = "Percent_Change"
                Range("M1").Value = "Total_Volume"
                Range("J1:M1").Font.Bold = True
                
                ' Print Ticker Name in Summary Table
                Range("J" & Summary_Table_Row).Value = Ticker
                
                ' Print yearly change in Summary Table
                Range("K" & Summary_Table_Row).Value = change
                
                ' Conditional formatting of yearly change
                Dim MyRange As Range
                Set MyRange = Range("K" & Summary_Table_Row)
                
                ' Delete any existing formatting
                MyRange.FormatConditions.Delete
                
                ' Apply Conditional Formatting
                
                ' add first rule for negative values
                MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                    Formula1:="=0"
                MyRange.FormatConditions(1).Interior.Color = vbRed
                
                ' add second rule for 0 and positive values
                MyRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                    Formula1:="=0"
                MyRange.FormatConditions(2).Interior.Color = vbGreen
                
                ' Print percent change in Summary Table
                Range("L" & Summary_Table_Row).Value = percent
                Range("L" & Summary_Table_Row).NumberFormat = "0.00%"

                ' Print Volume Total Amount in Summary Table
                Range("M" & Summary_Table_Row).Value = Volume
                
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset the Volume Total
                Volume = 0
                
                ' Reset yearly_change
                yearly_change = 0
                
                ' Reset percent_chg
                percent_chg = 0
                
                ' If cell immediately following a row is the same ticker
                Else
                
                ' Add to the Volume Total
                Volume = Volume + Cells(row, 7).Value
                
                If openv = 0 Then
                openv = Cells(row, 3).Value
                    
                End If
                
                End If

            Next row
            
End Sub