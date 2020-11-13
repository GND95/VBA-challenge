Sub StockAnalysis()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Dim tickerSymbol As String
    Dim openingValue, closingValue As Double
    Dim isOpeningValue As Boolean
    Dim startingRow As Integer
    isOpeningValue = True
    startingRow = 2
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            If isOpeningValue = True Then
                openingValue = Cells(i, 3)
                isOpeningValue = False
                ' store the value of the opening price BUT only do it once so it's not overwritten by future prices that are not the first opening price
            End If
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            isOpeningValue = True
            tickerSymbol = Cells(i, 1).Value
            Cells(startingRow, 9).Value = tickerSymbol
            startingRow = startingRow + 1
            ' if cells are different then i am going to want to write the ticker symbol out to the "I" column and then increment the row counter variable
            ' and calculate the difference between the start of the year and now, since in the else block this is the last instance of the prior ticker symbol appearing
        End If
    Next i
End Sub


