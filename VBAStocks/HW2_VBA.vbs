Option Explicit

Sub BudgetChecker()

 'title new columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'define variables
    Dim ticker_symbol As String
    Dim i As Long
    Dim ticker_row As Double
    Dim column As Integer
    column = 1
    ticker_row = 2
    
    For i = 2 To 70926

        'Set cell value to variable
        ticker_symbol = Cells(i, column).Value
        
        'if ticker symbol in prior row does not equal ticker symbol in current row, print to ticker symbol column
        If Cells(i + 1, column).Value <> ticker_symbol Then
            Cells(ticker_row, 9).Value = ticker_symbol
            'make sure new ticker symbol prints in next row
            ticker_row = ticker_row + 1
        End If
        
    Next i


End Sub