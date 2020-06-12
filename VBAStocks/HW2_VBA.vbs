Option Explicit

Sub Stonks()

 'title new columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'define variables
    Dim ticker_symbol As String
    Dim ticker_row As Double
    Dim ann_price_change As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim i As Double
    ticker_row = 2
    
    
    For i = 2 To 70926

        'Set cell value to variable
        ticker_symbol = Cells(i, 1).Value
        
        'if ticker symbol in prior row does not equal ticker symbol in current row, print to ticker symbol column
        If Cells(i + 1, 1).Value <> ticker_symbol Then
            Cells(ticker_row, 9).Value = ticker_symbol
            open_price = Cells(i-261, 3).Value
            close_price = Cells(i, 6).Value
            ann_price_change = close_price - open_price
            Cells(ticker_row, 10).Value = ann_price_change
            'make sure new ticker symbol prints in next row
            ticker_row = ticker_row + 1
            
        End If
        
    Next i


End Sub
