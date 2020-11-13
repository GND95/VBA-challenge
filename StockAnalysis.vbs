Sub StockAnalysis()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Dim tickerSymbol As String
    Dim openingValue, closingValue, yearlyChangeValue, yearlyChangePercent As Double
    Dim isOpeningValue As Boolean
    Dim startingRow As Integer
    isOpeningValue = True
    startingRow = 2
    
    For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            If isOpeningValue = True Then ' store the value of the opening price BUT only do it once so it's not overwritten by future prices that are not the first opening price
                openingValue = Cells(i, 3)
                isOpeningValue = False
            End If
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerSymbol = Cells(i, 1).Value
            closingValue = Cells(i, 6).Value ' set the closing value of the stock price because in the else block this will be the last instance of the prior ticker symbol appearing
            yearlyChangeValue = closingValue - openingValue ' calculate the difference in stock price between the start of the year and now
            yearlyChangePercent = (closingValue - openingValue) / openingValue ' using the percent change math formula
            If yearlyChangeValue < 0 Then ' if it's less than 0 then it's a negative number and should be red
                Cells(startingRow, 10).Interior.ColorIndex = 3
            ElseIf yearlyChangeValue > 0 Then ' if it's greater than 0 then it's a positive number and should be green; 0.00 values will not be colored as they are neither positive nor negative numbers
                Cells(startingRow, 10).Interior.ColorIndex = 10
            End If
            Cells(startingRow, 9).Value = tickerSymbol ' if cells are different then i am going to want to write the ticker symbol out to the "I" column and then increment the row counter variable
            Cells(startingRow, 10).Value = Round(yearlyChangeValue, 2) ' round number down to two decimal places and put value in column j
            Cells(startingRow, 11).Value = yearlyChangePercent
            Cells(startingRow, 11).NumberFormat = "0.00%" ' format value to be a percent with two decimal places
            startingRow = startingRow + 1 'increment the row counter variable
            isOpeningValue = True ' set the bool back to true since the next iteration of the loop will be the opening value for the next stock
        End If
    Next i
End Sub



