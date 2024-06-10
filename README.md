# VBA-Challenge


Sub Multiple_year_stock_dataPKA()
 
 
 ' Set an initial variable for holding the ticker name
 Dim Ticker As String
 
          
          ' Set an initial variable for holding the Ticker Total Volume per Ticker

    
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
    
    Dim Open_Price As Double
    
    Open_Price = Cells(2, 3).Value
    
    Dim Close_Price As Double
    
    Dim Quarterly_Change As Double
    
    Dim Percent_Change As Double
    
    Dim Summary_Table_Row As Double
    
    Summary_Table_Row = 2
    
          ' Set headers for summary table
          
    Cells(1, 9).Value = "Ticker"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 10).Value = "Quarterly Change"
    Cells(1, 11).Value = "Percent Change"

        ' Determine the last Row
        
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
             ' check for same ticker symbol
             
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
            Ticker = Cells(i, 1).Value
    
    'Accumulate total stock volume
    
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
   
            ' calculate and print quarterly change
            Close_Price = Cells(i, 6).Value
    
            Quarterly_Change = Close_Price - Open_Price
            
            If Open_Price = 0 Then
                Percent_Change = 0
                
                      ' calculate and print percent change
                      
            Else
                Percent_Change = (Quarterly_Change / Open_Price)
                
            End If
            
        
            Range("J" & Summary_Table_Row).Value = Quarterly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
    
    
            Summary_Table_Row = Summary_Table_Row + 1
            Total_Stock_Volume = 0
            Open_Price = Cells(i + 1, 3).Value
        
        Else
    
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
        
        End If
    
    Next i
    
    Summary_Table_Last_Row = Cells(Rows.Count, 9).End(xlUp).Row
    
        
        For i = 2 To Summary_Table_Last_Row
        
        ' conditional Formatting
        
        
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            ElseIf Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.ColorIndex = 3
                
Else
Cells(i, 10).Interior.ColorIndex = 0
            
            End If
        
        Next i
        
    '## Bonus
    
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest total volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    
        For i = 2 To Summary_Table_Last_Row
            
            If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & Summary_Table_Last_Row)) Then
                Cells(2, 15).Value = Cells(i, 9).Value
                Cells(2, 16).Value = Cells(i, 11).Value
                Cells(2, 16).NumberFormat = "0.00%"
            
            ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & Summary_Table_Last_Row)) Then
                Cells(3, 15).Value = Cells(i, 9).Value
                Cells(3, 16).Value = Cells(i, 11).Value
                Cells(3, 16).NumberFormat = "0.00%"
                
            ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & Summary_Table_Last_Row)) Then
                Cells(4, 15).Value = Cells(i, 9).Value
                Cells(4, 16).Value = Cells(i, 12).Value
                
            End If
            
        Next i
        
        
    
End Sub

ALPHABETICAL TESTING

Sub alphabetical_testingPKA()

    'loop through all sheets
For Each ws In Worksheets

'set the header as the requasted
ws.Cells(1, "i").Value = "Ticker"
ws.Cells(1, "j").Value = "Quarterly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"
ws.Cells(2, "O").Value = "Greatest % increase"
ws.Cells(3, "O").Value = "Greatest % decrease"
ws.Cells(4, "O").Value = "Greatest total volume"
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"

'identify all variables in part I
Dim current_ticker As String
Dim output_index As Double
output_index = 2
Dim QuarterlyChange As Double
Dim OpeningPrice As Double
OpeningPrice = ws.Cells(2, 3).Value
Dim ClosingPrice As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double

'loop for part I by determining the last row
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow

'comparing cells
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'introduce the ticker name
current_ticker = ws.Cells(i, 1).Value

'calculating the changes over a year for each tickers
ClosingPrice = ws.Cells(i, 6).Value
QuarterlyChange = (ClosingPrice - OpeningPrice)

'calculating the change percent for each tickers by using conditions
If OpeningPrice = 0 Then
PercentChange = 0
Else
PercentChange = (QuarterlyChange / OpeningPrice)
End If

'finding the total stock volume for each tickers
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)

'printing all result in identified cells
ws.Cells(output_index, 9).Value = current_ticker
ws.Cells(output_index, 10).Value = QuarterlyChange
ws.Cells(output_index, 11).Value = PercentChange
ws.Cells(output_index, 11).NumberFormat = "0.00%"
ws.Cells(output_index, 12).Value = TotalStockVolume

'Add a new row
output_index = output_index + 1

'reset the total stock volume
TotalStockVolume = 0

'add new opening price
OpeningPrice = ws.Cells(i + 1, 3).Value
Else
TotalStockVolume = TotalStockVolume + ws.Cells(i, 7)
End If

'conditional formatting for Quarterly change
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
Else
ws.Cells(i, 10).Interior.ColorIndex = 0
End If
Next i

'finding the maximum of percent changing and its similar ticker
max_change = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
max_ticker = WorksheetFunction.Match(max_change, ws.Range("K2:K" & lastRow), 0)
ws.Cells(2, 17).Value = max_change
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(2, 16).Value = ws.Cells(max_ticker + 1, 9)

'finding the minimum of percent changing and its similar ticker
min_change = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
min_ticker = WorksheetFunction.Match(min_change, ws.Range("K2:K" & lastRow), 0)
ws.Cells(3, 17).Value = min_change
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = ws.Cells(min_ticker + 1, 9)

'finding the maximum of total stock volume and its similar ticker
max_volume = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
max_total_ticker = WorksheetFunction.Match(max_volume, ws.Range("L2:L" & lastRow), 0)
ws.Cells(4, 17).Value = max_volume
ws.Cells(4, 16).Value = ws.Cells(max_total_ticker + 1, 9)

Next ws

End Sub



