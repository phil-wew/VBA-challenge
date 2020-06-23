Sub yearly_stock()

Dim ws As Worksheet

'loop through every ws
For Each ws In ActiveWorkbook.Worksheets

    
        
' Assigning Variables
Dim stock_ticker As String
    
Dim total_volume As Double
        total_volume = 0
    
Dim yearly_change As Double
        yearly_change = 0

Dim closing_price As Double
        closing_price = 0
    
    
Dim summary_table As Double
        summary_table = 2
        
        
Dim opening_price As Double
        opening_price = Cells(2, 3)
        Range("L" & summary_table).Value = opening_price
        
Dim percent_change As Double
        percent_change = 0

Dim greatest_increase As Double
        greatest_increase = 0
        
Dim increasec_ticker As String
        increase_ticker = ""
        

Dim greatest_decrease As Double
        greatest_decrease = 0
        
Dim decrease_inc_ticker As String
        decrease_ticker = ""
        
Dim greatest_volume As Double
        greatest_volume = 0

Dim volume_ticker As String
        volume_ticker = ""
        

        
     ' Setting Row/column Titles
    ws.Cells(1, 10) = "stock_ticker"
    ws.Cells(1, 11) = "closing_price"
    ws.Cells(1, 12) = "opening_price"
    ws.Cells(1, 13) = "total_volume"
    ws.Cells(1, 14) = "yearly_change"
    ws.Cells(1, 15) = "percent_change"
    ws.Cells(2, 18) = "greatest_incease"
    ws.Cells(3, 18) = "greatest_decrease"
    ws.Cells(4, 18) = "greatest_volume"
    ws.Cells(1, 19) = "ticker"
    ws.Cells(1, 20) = "value"
    
    
        
    ' Get row number of last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
For I = 2 To lastrow
        
    'Check if new ticker
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
        
             'set ticker name
                 stock_ticker = ws.Cells(I, 1).Value
            

                  'Retrieve Values
                        ws.Range("J" & summary_table).Value = stock_ticker
                    
                    
                    
            
    'set closing price
    closing_price = ws.Cells(I, 6).Value
            
        'Print the closing price to the Summary Table
        ws.Range("K" & summary_table).Value = closing_price
        
    
    
    
    'Sum total volume
    total_volume = total_volume + ws.Cells(I, 7).Value
            
        'retrieve value to put in summary table
        ws.Range("M" & summary_table).Value = total_volume
        
            'reset total volume
            total_volume = 0
            
            
            
    'set opening price
    opening_price = ws.Cells(I + 1, 3)
    
         'retrieve value to put in summary table
        ws.Range("L" & summary_table + 1).Value = opening_price
                
        
                
    'Calc yearly change
    yearly_change = closing_price - opening_price
        
        'place yearly change in table
        ws.Range("N" & summary_table).Value = yearly_change
        
        
        
    If opening_price > 0 Then
    
    'calculate % change
    percent_change = (closing_price - opening_price) / opening_price
                        
        'place % change in table
        ws.Range("O" & summary_table).Value = percent_change
        
            'Format for %
            ws.Range("O" & summary_table).NumberFormat = "0.00%"
            
    Else
        percent_change = 0
        
    End If
    
            
    'Next row in summary table
    summary_table = summary_table + 1
        
        
      'If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      total_volume = total_volume + ws.Cells(I, 3).Value
    
            
    End If


If ws.Range("N" & summary_table).Value < 0 Then
    
    ws.Range("N" & summary_table).Interior.ColorIndex = 3
    
    Else
    
    ws.Range("N" & summary_table).Interior.ColorIndex = 4
    
    End If
    
    'if i is greater than current greatest set current as GI
    If ws.Cells(I, 14).Value > greatest_increase Then
    
        greatest_increase = ws.Cells(I, 14).Value
            
            
             'place Greatest in table
                ws.Cells(2, 20) = greatest_increase
        
        
        
        'ticker associated with greatest increase
        increase_ticker = ws.Cells(I, 10).Value
        
        
             'place ticker in table
              ws.Cells(2, 19) = increase_ticker
              
        
    End If
    
    
    'if i is greater than current greatest set current as GI
    If ws.Cells(I, 14).Value < greatest_decrease Then
    
        greatest_decrease = ws.Cells(I, 14).Value
            
            
             'place Greatest in table
                ws.Cells(3, 20) = greatest_decrease
        
        
        
        'ticker associated with greatest increase
        decrease_ticker = ws.Cells(I, 10).Value
        
        
             'place ticker in table
              ws.Cells(3, 19) = decrease_ticker
              
    End If
    

   
   'if i is greater than current greatest make current
    If ws.Cells(I, 13).Value > greatest_volume Then
    
        greatest_volume = ws.Cells(I, 13).Value
            
            
             'place Greatest in table
                ws.Cells(4, 20) = greatest_volume
        
        
        
        'ticker associated with greatest increase
        volume_ticker = ws.Cells(I, 10).Value
        
        
             'place ticker in table
              ws.Cells(4, 19) = volume_ticker
              
   
    End If
    
    Next I
    
    Next ws
    
    
    
    End Sub
