Sub StockAnalysis()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Dim tickerSymbol As String
    Dim openingValue, closingValue, yearlyChangeValue, yearlyChangePercent, maxPercentChange, minPercentChange, maxVolume, totalVolume As Double
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
            totalVolume = totalVolume + Cells(i, 7).Value ' for every iteration of the loop set the stock volume equal to itself plus the volume of the current row of volume data for the ticker
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            tickerSymbol = Cells(i, 1).Value
            closingValue = Cells(i, 6).Value ' set the closing value of the stock price because in the else block this will be the last instance of the prior ticker symbol appearing
            totalVolume = totalVolume + Cells(i, 7).Value ' set the stock volume equal to itself plus the volume of the current (final) row of volume data for the ticker
            yearlyChangeValue = closingValue - openingValue ' calculate the difference in stock price between the start of the year and now
            If openingValue <> 0 Then ' account for the cases when the opening value of a stock is 0 by skipping this stock as you cannot divide by zero to calculate percent change
                yearlyChangePercent = (closingValue - openingValue) / openingValue ' using the percent change math formula
            End If
            If yearlyChangeValue < 0 Then ' if it's less than 0 then it's a negative number and should be red
                Cells(startingRow, 10).Interior.ColorIndex = 3
            ElseIf yearlyChangeValue > 0 Then ' if it's greater than 0 then it's a positive number and should be green; 0.00 values will not be colored as they are neither positive nor negative numbers
                Cells(startingRow, 10).Interior.ColorIndex = 10
            End If
            Cells(startingRow, 9).Value = tickerSymbol ' if cells are different then i am going to want to write the ticker symbol out to the "I" column and then increment the row counter variable
            Cells(startingRow, 10).Value = Round(yearlyChangeValue, 2) ' round number down to two decimal places and put value in column j
            Cells(startingRow, 11).Value = yearlyChangePercent
            Cells(startingRow, 11).NumberFormat = "0.00%" ' format value to be a percent with two decimal places
            Cells(startingRow, 12).Value = totalVolume
            startingRow = startingRow + 1 'increment the row counter variable
            isOpeningValue = True ' set the bool back to true since the next iteration of the loop will be the opening value for the next stock
            totalVolume = 0 ' reset total volume variable back to zero so the next stock volume can be tracked
        End If
    Next i
    maxVolume = WorksheetFunction.Max(Range("L2:L" + CStr(Cells(Rows.Count, 12).End(xlUp).Row))) 'getting the largest stock volume value from the column of stock volumes
    minPercentChange = WorksheetFunction.Min(Range("K2:K" + CStr(Cells(Rows.Count, 11).End(xlUp).Row))) 'getting the smallest stock percent change value from the column of stock percent change
    maxPercentChange = WorksheetFunction.Max(Range("K2:K" + CStr(Cells(Rows.Count, 11).End(xlUp).Row))) 'getting the largest stock percent change value from the column of stock percent change
    For i = 2 To Cells(Rows.Count, 12).End(xlUp).Row ' loop through the volume column until we find a value that matches the max volume
        If Cells(i, 12) = maxVolume Then
            Range("P4").Value = Cells(i, 9) ' if the volume is a match go to the Ticker column of that same row to retrieve which stock ticker has the highest volume
            Range("Q4").Value = maxVolume ' if the volume is a match then put the volume value into the Value column
            Exit For ' break out of for loop if a volume value match is found
        End If
    Next i
    For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row ' loop through the percent change column until we find a value that matches the min percent change
        If Cells(i, 11) = minPercentChange Then
            Range("P3").Value = Cells(i, 9) ' if the min percent change is a match go to the Ticker column of that same row to retrieve which stock ticker has the lowest percent change
            Range("Q3").Value = minPercentChange ' if the percent change is a match then put the percent change value into the Value column
            Range("Q3").NumberFormat = "0.00%" ' format value to be a percent with two decimal places
            Exit For ' break out of for loop if a percent change value match is found
        End If
    Next i
    For i = 2 To Cells(Rows.Count, 11).End(xlUp).Row ' loop through the percent change column until we find a value that matches the max percent change
        If Cells(i, 11) = maxPercentChange Then
            Range("P2").Value = Cells(i, 9) ' if the max percent change is a match go to the Ticker column of that same row to retrieve which stock ticker has the highest percent change
            Range("Q2").Value = maxPercentChange ' if the percent change is a match then put the percent change value into the Value column
            Range("Q2").NumberFormat = "0.00%" ' format value to be a percent with two decimal places
            Exit For ' break out of for loop if a percent change value match is found
        End If
    Next i
End Sub