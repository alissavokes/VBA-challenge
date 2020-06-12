Option Explicit

Sub Stonks()

    'define variables
    Dim ticker_symbol As String
    Dim ticker_row As Long
    Dim last_row As Long
    Dim ws As Worksheet
    Dim ann_price_change As Double
    Dim array_counter As Long
    Dim open_price() As Double
    Dim close_price As Double
    Dim percent_change As Variant
    Dim stock_volume As Variant
    Dim i As Long


    For Each ws In Worksheets
        ws.activate

        'title new columns
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
    
        'set ticker row and array counter
        ticker_row = 2
        array_counter = 0
        ReDim open_price(0 to array_counter)

        'Determine the Last Row
        last_row = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To last_row

            'Set cell value to variable
            ticker_symbol = Cells(i, 1).Value
            
            'if ticker symbol in prior row does not equal ticker symbol in current row, print to ticker symbol column
            If Cells(i + 1, 1).Value <> ticker_symbol Then
                Cells(ticker_row, 9).Value = ticker_symbol

                'calculate close price for current stock
                close_price = Cells(i, 6).Value
                'determine annual price change
                ann_price_change = close_price - open_price(0)
                'print price change to corresponding ticker symbol
                Cells(ticker_row, 10).Value = ann_price_change

                    'conditional formatting: +change = green, -change = red
                    If ann_price_change < 0 Then
                        Cells(ticker_row, 10).Interior.ColorIndex = 3
                    Else
                        Cells(ticker_row, 10).Interior.ColorIndex = 4
                    End If
                'determine percent change
                    If open_price(0) <> 0 then
                        percent_change = CDec(ann_price_change / open_price(0))
                    Else
                        percent_change = 0
                    End If

                'print percent change to corresponding ticker symbol
                Cells(ticker_row, 11).Value = Format(percent_change, "Percent")
                
                'add last sotck volume to counter
                stock_volume = stock_volume + Cells(i, 7).Value
                'print stock volume to corresponding ticker symbol
                Cells(ticker_row, 12).Value = stock_volume

                'make sure new ticker symbol prints in next row
                ticker_row = ticker_row + 1

                'reset stock volume/array counter for next stock
                stock_volume = 0
                array_counter = 0
                ReDim open_price(0 to array_counter)
                

            Else
                'counting volume for current stock
                stock_volume = CDec(stock_volume + Cells(i, 7).Value)

                If array_counter > 0 Then   
                    ReDim Preserve open_price(0 to array_counter)
                End If
                
                open_price(array_counter) = Cells(i, 3).Value
                array_counter = array_counter + 1
                
            End If
            
        Next i

    Next ws

End Sub

