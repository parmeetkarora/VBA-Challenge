Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data_corrected()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Total_Stock_Volume As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    Dim Summary_Table_Row As Long
    Dim Summary_Table_Last_Row As Long
    
    Set wb = ThisWorkbook
    
    For Each ws In wb.Worksheets
        ' Reset variables for each sheet
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
        
   
          ' Set headers for summary table
          
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
    
            ' Determine the last Row
            
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
                 ' check for same ticker symbol
                 
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                Ticker = ws.Cells(i, 1).Value
        
        'Accumulate total stock volume
        
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        
       
                ' calculate and print quarterly change
                Close_Price = ws.Cells(i, 6).Value
        
                Quarterly_Change = Close_Price - Open_Price
                
                If Open_Price = 0 Then
                    Percent_Change = 0
                    
                          ' calculate and print percent change
                          
                Else
                    Percent_Change = (Quarterly_Change / Open_Price)
                    
                End If
                
            
                ws.Range("J" & Summary_Table_Row).Value = Quarterly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        
                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
                Open_Price = ws.Cells(i + 1, 3).Value
            
            Else
        
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        
            
            End If
        
        Next i
        
        Summary_Table_Last_Row = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
            
            For i = 2 To Summary_Table_Last_Row
            
            ' conditional Formatting
            
            
                If ws.Cells(i, 10).Value > 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(i, 10).Value < 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                    
    Else
    ws.Cells(i, 10).Interior.ColorIndex = 0
                
                End If
            
            Next i
            
        ' Bonus
        
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest total volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
            For i = 2 To Summary_Table_Last_Row
                
                If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Last_Row)) Then
                    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).NumberFormat = "0.00%"
                
                ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Last_Row)) Then
                    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).NumberFormat = "0.00%"
                    
                ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Last_Row)) Then
                    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
                    
                End If
                
            Next i
        
 Next ws
    
End Sub






