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
    Cells(1, 10) = "stock_ticker"
    Cells(1, 11) = "closing_price"
    Cells(1, 12) = "opening_price"
    Cells(1, 13) = "total_volume"
    Cells(1, 14) = "yearly_change"
    Cells(1, 15) = "percent_change"
    Cells(2, 18) = "greatest_incease"
    Cells(3, 18) = "greatest_decrease"
    Cells(4, 18) = "greatest_volume"
    Cells(1, 19) = "ticker"
    Cells(1, 20) = "value"
    
    
        
    ' Get row number of last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
For I = 2 To lastrow
        
    'Check if new ticker
    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
             'set ticker name
                 stock_ticker = Cells(I, 1).Value
            

                  'Retrieve Values
                        Range("J" & summary_table).Value = stock_ticker
                    
                    
                    
            
    'set closing price
    closing_price = Cells(I, 6).Value
            
        'Print the closing price to the Summary Table
        Range("K" & summary_table).Value = closing_price
        
    
    
    
    'Sum total volume
    total_volume = total_volume + Cells(I, 7).Value
            
        'retrieve value to put in summary table
        Range("M" & summary_table).Value = total_volume
        
            'reset total volume
            total_volume = 0
            
            
            
    'set opening price
    opening_price = Cells(I + 1, 3)
    
         'retrieve value to put in summary table
        Range("L" & summary_table + 1).Value = opening_price
                
                
                
    'Calc yearly change
    yearly_change = closing_price - opening_price
        
        'place yearly change in table
        Range("N" & summary_table).Value = yearly_change
        
        
            
    'calculate % change
    percent_change = (closing_price - opening_price) / opening_price
                        
        'place % change in table
        Range("O" & summary_table).Value = percent_change
        
            'Format for %
            Range("O" & summary_table).NumberFormat = "0.00%"
            
    'Next row in summary table
    summary_table = summary_table + 1
        
        
      'If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      total_volume = total_volume + Cells(I, 3).Value
    
            
    End If


If Range("N" & summary_table).Value < 0 Then
    
    Range("N" & summary_table).Interior.ColorIndex = 3
    
    Else
    
    Range("N" & summary_table).Interior.ColorIndex = 4
    
    End If
    
    'if i is greater than current greatest set current as GI
    If Cells(I, 14).Value > greatest_increase Then
    
        greatest_increase = Cells(I, 14).Value
            
            
             'place Greatest in table
                Cells(2, 20) = greatest_increase
        
        
        
        'ticker associated with greatest increase
        increase_ticker = Cells(I, 10).Value
        
        
             'place ticker in table
              Cells(2, 19) = increase_ticker
              
        
    End If
    
    
    'if i is greater than current greatest set current as GI
    If Cells(I, 14).Value < greatest_decrease Then
    
        greatest_decrease = Cells(I, 14).Value
            
            
             'place Greatest in table
                Cells(3, 20) = greatest_decrease
        
        
        
        'ticker associated with greatest increase
        decrease_ticker = Cells(I, 10).Value
        
        
             'place ticker in table
              Cells(3, 19) = decrease_ticker
              
    End If
    

   
   'if i is greater than current greatest make current
    If Cells(I, 13).Value > greatest_volume Then
    
        greatest_volume = Cells(I, 13).Value
            
            
             'place Greatest in table
                Cells(4, 20) = greatest_volume
        
        
        
        'ticker associated with greatest increase
        volume_ticker = Cells(I, 10).Value
        
        
             'place ticker in table
              Cells(4, 19) = volume_ticker
              
   
    End If
    
    Next I
    
    Next ws
    
    
    
    End Sub

