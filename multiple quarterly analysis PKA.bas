Attribute VB_Name = "Module1"
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

