Sub stock_price_hw()

'loop through all sheets
For Each ws In Worksheets

    'declare variables: ticker, summary_table_index, open_price, closing_price
        
    Dim last_row As Long
    Dim ticker As String
    Dim ticker_count As Integer
    Dim summary_table_index As Integer
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    
    
    'iterate through all rows
           
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_count = 0
    summary_table_index = 2
    total_volume = 0
        
   
    
    'initialize first open_price and ticker
    
    open_price = ws.Cells(2, 3).Value
    ticker = ws.Cells(2, 1).Value
    
    
    For i = 2 To last_row
        
       close_price = ws.Cells(i, 6).Value
       total_volume = total_volume + ws.Cells(i, 7).Value
       
     
        'check if the cell below is the same? no action: then i have a new ticker
        If ticker <> ws.Cells(i + 1, 1).Value Then
            yearly_change = close_price - open_price
          
            If open_price <> 0 Then
                percent_change = yearly_change / open_price
            
            Else
                percent_change = 100
            
            End If
            
            
            
            ws.Range("I" & summary_table_index).Value = ticker
            ws.Range("J" & summary_table_index).Value = yearly_change
                If yearly_change >= 0 Then
                    ws.Range("J" & summary_table_index).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & summary_table_index).Interior.ColorIndex = 3
                End If
            ws.Range("K" & summary_table_index).Value = percent_change
            ws.Range("K" & summary_table_index).NumberFormat = "0.00%"
            ws.Range("L" & summary_table_index).Value = total_volume
            
            
            'move summary_table_index to the next row
            
            summary_table_index = summary_table_index + 1
            
            'reset the open price and ticker for the next ticker
            
            open_price = ws.Cells(i + 1, 3).Value
            ticker = ws.Cells(i + 1, 1).Value
            total_volume = 0
            
        End If
        
        
        
    Next i
    
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    

Next ws

    
    
End Sub


