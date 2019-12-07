Attribute VB_Name = "M—dulo1"
Sub StockMarketForMultipleYears()

' Define the type of variables

   Dim TotalStockVolume As Double
   Dim TickerName As String
   Dim YearlyChange As Double
   Dim YearlyOpenPrice As Double
   Dim YearlyClosePrice As Double
   Dim PercentChange As Double
   Dim LastRow As Double
   Dim SummaryTickerRow As Double
   Dim i As Double
   Dim OpenAmount As Double
   
' Iterate through all worksheets in the Stock Market dataset

    For Each ws In Worksheets
       
        ws.Cells(1, 9).Value = "Ticker" ' cell I1
        ws.Cells(1, 10).Value = "Yearly Change"  'cell J1
        ws.Cells(1, 11).Value = "Percent Change" ' cell K1
        ws.Cells(1, 12).Value = "Total Stock Volume" ' cell L1
        ws.Cells(2, 15).Value = "Greatest % Increase" ' cell O1
        ws.Cells(3, 15).Value = "Greatest % Decrease" ' cell O2
        ws.Cells(4, 15).Value = "Greatest Total Volume" ' cell O3
        ws.Cells(1, 16).Value = "Ticker" ' cell P1
        ws.Cells(1, 17).Value = "Value" ' cell Q1
        
        
  ' Initialize
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  SummaryTickerRow = 2
  OpenAmount = 2
  TotalStockVolume = 0
  
  For i = 2 To LastRow
     TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
      TickerName = ws.Cells(i, 1).Value
                ' Print The Ticker Name In The Summary Table
      ws.Range("I" & SummaryTickerRow).Value = TickerName
                ' Print The Ticker Total Amount To The Summary Table
      ws.Range("L" & SummaryTickerRow).Value = TotalStockVolume
                ' Reset Ticker Total
      TotalStockVolume = 0
     
     ' Get Yearly Open, Yearly Close and Yearly Change
                
            YearlyOpenPrice = ws.Range("C" & OpenAmount)
            YearlyClosePrice = ws.Range("F" & i)
            YearlyChange = YearlyClosePrice - YearlyOpenPrice
            ws.Range("J" & SummaryTickerRow).Value = YearlyChange
                
    
       '  The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
                If YearlyOpenPrice = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpenPrice = ws.Range("C" & OpenAmount)
                    PercentChange = YearlyChange / YearlyOpenPrice
                End If
                ' Format to include % symbol and two decimal places
                ws.Range("K" & SummaryTickerRow).NumberFormat = "0.00%"
                ' Assign % Change
                ws.Range("K" & SummaryTickerRow).Value = PercentChange

                ' Conditional Formatting Assign (Green) to positive OR Assign (Red) to negative
                If ws.Range("J" & SummaryTickerRow).Value >= 0 Then
                    ws.Range("J" & SummaryTickerRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTickerRow).Interior.ColorIndex = 3
                End If
            
                ' Add one to Summary Stock Row variable
                SummaryTickerRow = SummaryTickerRow + 1
                OpenAmount = i + 1
                End If
    Next i
        
    Next ws
    
    Call IncreaseDecrease
    
End Sub


 ' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
         
        Sub IncreaseDecrease()
        
        Dim max_increase As Double
        Dim min_increase As Double
        Dim max_tot_value As Double
        Dim LastRow1 As Integer
        Dim LastRow2 As Integer
        Dim TickerMin As String
        Dim TickerMax As String
        Dim TickerMaxTot As String
        Dim i As Integer
        Dim Rng As Variant
        Dim Rng1 As Variant
        
         ' Iterate through all worksheets in the Stock Market dataset
     For Each ws In Worksheets
        ' Set the range for column
          Set Rng = ws.Columns(11)
           ' Get the max, min and total value in order to compare later
           max_increase = ws.Application.WorksheetFunction.Max(Rng)
           ws.Range("Q2").Value = max_increase
    
          min_decrease = ws.Application.WorksheetFunction.Min(Rng)
          ws.Range("Q3") = min_decrease
          Set Rng1 = ws.Columns(12)
          max_tot_value = ws.Application.WorksheetFunction.Max(Rng1)
          ws.Range("Q4").Value = max_tot_value
          
            ' MsgBox (max_tot_value)
          ' Gets the total of rows to start the loop and compare max and min numbers to assigne the corresponding ticker
          LastRow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
       For i = 2 To LastRow2
            If ws.Cells(i, 11).Value = min_decrease Then
             TickerMin = ws.Cells(i, 9).Value
           End If
           If ws.Cells(i, 11).Value = max_increase Then
             TickerMax = ws.Cells(i, 9).Value
           End If
           If ws.Cells(i, 12).Value = max_tot_value Then
           ' MsgBox (ws.Cells(i, 12).Value)
             TickerMaxTot = ws.Cells(i, 9).Value
      End If
            
      Next i
         ' MsgBox (TickerMaxTot)
      ' Write the max, min and total numbers to P2, P3 and P4 cells/positions
    ws.Range("P2") = TickerMax
    ws.Range("P3") = TickerMin
    ws.Range("P4") = TickerMaxTot
        
     ' Finish Formatting to include % symbol and two decimal places

     ws.Range("Q2").NumberFormat = "0.00%"
     ws.Range("Q3").NumberFormat = "0.00%"
            
        ' Formatting columns to Autofit
    ws.Columns("I:Q").AutoFit
        
Next ws

End Sub



