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
    Dim percent_change as Double
    Dim stock_volume as Variant
    Dim i As Double
    ticker_row = 2
    
    
    For i = 2 To 70926

        'Set cell value to variable
        ticker_symbol = Cells(i, 1).Value
        
        'if ticker symbol in prior row does not equal ticker symbol in current row, print to ticker symbol column
        If Cells(i + 1, 1).Value <> ticker_symbol Then
            Cells(ticker_row, 9).Value = ticker_symbol

            'calculate open price for current stock
            open_price = Cells(i-261, 3).Value
            'calculate close price for current stock
            close_price = Cells(i, 6).Value
            'determine annual price change
            ann_price_change = close_price - open_price
            'print price change to corresponding ticker symbol
            Cells(ticker_row, 10).Value = ann_price_change

                'conditional formatting: +change = green, -change = red
                If ann_price_change < 0 Then
                    Cells(ticker_row, 10).Interior.ColorIndex = 3
                Else
                    Cells(ticker_row, 10).Interior.ColorIndex = 4
                End If
            'determine percent change
            percent_change = ann_price_change / open_price
            'print percent change to corresponding ticker symbol
            Cells(ticker_row, 11).Value = Format(percent_change, "Percent")
            
            'add last sotck volume to counter
            stock_volume = stock_volume + Cells(i, 7).Value
            'print stock volume to corresponding ticker symbol
            Cells(ticker_row, 12).Value = stock_volume

            'make sure new ticker symbol prints in next row
            ticker_row = ticker_row + 1

            'reset stock volume counter for next stock
            stock_volume = 0
        Else
            'counting volume for current stock
            stock_volume = CDec(stock_volume + Cells(i, 7).Value)
            
        End If
        
    Next i


End Sub
