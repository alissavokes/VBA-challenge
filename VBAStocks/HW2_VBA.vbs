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
    
    For i = 2 To 70926

        'Set cell value to variable
        ticker_symbol = Cells(i, 1).Value
        Cells(i, 9).Value = ticker_symbol
    Next i


End Sub
