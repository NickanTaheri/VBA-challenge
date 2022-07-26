Sub wallstreet()

Dim LastRow, i, j As Integer

Dim Headers(3) As String
Headers(0) = "Ticker"
Headers(1) = "Yearly Change"
Headers(2) = "Percent Change"
Headers(3) = "Total Stock Volume"



For Each ws In Worksheets

    'Clean Previous results
    ws.Range("J:M").Value = Null
    
    'Add Header
    Dim c As Range
    
    i = 0
    For Each c In ws.[J1:M1]
    c.Value = Headers(i)
    c.Font.Bold = True
    i = i + 1
    Next c

    j = 2
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim ticker_opening_price, ticker_closing_price, total_stock_vol As Double
    
    total_stock_vol = 0
    
    For i = 2 To LastRow
    
        If i = 2 Then
                ticker_opening_price = ws.Cells(i, 3).Value
                
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'The ticker symbol.
            ws.Cells(j, 10) = ws.Cells(i, 1).Value
            
            'Yearly change from opening price at the beginning
            'of a given year to the closing price at the end of that year.
            
            ticker_closing_price = ws.Cells(i, 6).Value
            
            
            ws.Cells(j, 11) = ticker_closing_price - ticker_opening_price
            
            
            
            If ws.Cells(j, 11).Value < 0 Then
                ws.Cells(j, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(j, 11).Interior.ColorIndex = 4
            End If
            
            'The percent change from opening price at the beginning of a given year
            'to the closing price at the end of that year.
            
            ws.Cells(j, 12) = ws.Cells(j, 11) / ticker_opening_price
            ws.Cells(j, 12).NumberFormat = "0.00%"
            
            
            ws.Cells(j, 13) = total_stock_vol + ws.Cells(i, 7).Value
            
            total_stock_vol = 0
            
            ticker_opening_price = ws.Cells(i + 1, 3).Value
            
            j = j + 1
            
            
        
        
        Else


            total_stock_vol = total_stock_vol + ws.Cells(i, 7).Value
      
        
        End If
  Next i
  'MsgBox (ws.Name)
Next ws

'MsgBox (LastRow)

End Sub

