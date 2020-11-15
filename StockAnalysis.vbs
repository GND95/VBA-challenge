Sub StockAnalysis()
    For Each ws In Worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        Dim tickerSymbol As String
        Dim openingValue, closingValue, yearlyChangeValue, yearlyChangePercent, maxPercentChange, minPercentChange, maxVolume, totalVolume As Double
        Dim isOpeningValue As Boolean
        Dim startingRow As Integer
        isOpeningValue = True
        startingRow = 2
        
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
                If isOpeningValue = True Then ' store the value of the opening price BUT only do it once so it's not overwritten by future prices that are not the first opening price
                    openingValue = ws.Cells(i, 3)
                    isOpeningValue = False
                End If
                totalVolume = totalVolume + ws.Cells(i, 7).Value ' for every iteration of the loop set the stock volume equal to itself plus the volume of the current row of volume data for the ticker
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickerSymbol = ws.Cells(i, 1).Value
                closingValue = ws.Cells(i, 6).Value ' set the closing value of the stock price because in the else block this will be the last instance of the prior ticker symbol appearing
                totalVolume = totalVolume + ws.Cells(i, 7).Value ' set the stock volume equal to itself plus the volume of the current (final) row of volume data for the ticker
                yearlyChangeValue = closingValue - openingValue ' calculate the difference in stock price between the start of the year and now
                If openingValue <> 0 Then ' account for the cases when the opening value of a stock is 0 by skipping this stock as you cannot divide by zero to calculate percent change
                    yearlyChangePercent = (closingValue - openingValue) / openingValue ' using the percent change math formula
                End If
                If yearlyChangeValue < 0 Then ' if it's less than 0 then it's a negative number and should be red
                    ws.Cells(startingRow, 10).Interior.ColorIndex = 3
                ElseIf yearlyChangeValue > 0 Then ' if it's greater than 0 then it's a positive number and should be green; 0.00 values will not be colored as they are neither positive nor negative numbers
                    ws.Cells(startingRow, 10).Interior.ColorIndex = 10
                End If
                ws.Cells(startingRow, 9).Value = tickerSymbol ' if ws.Cells are different then i am going to want to write the ticker symbol out to the "I" column and then increment the row counter variable
                ws.Cells(startingRow, 10).Value = Round(yearlyChangeValue, 2) ' round number down to two decimal places and put value in column j
                ws.Cells(startingRow, 11).Value = yearlyChangePercent
                ws.Cells(startingRow, 11).NumberFormat = "0.00%" ' format value to be a percent with two decimal places
                ws.Cells(startingRow, 12).Value = totalVolume
                startingRow = startingRow + 1 'increment the row counter variable
                isOpeningValue = True ' set the bool back to true since the next iteration of the loop will be the opening value for the next stock
                totalVolume = 0 ' reset total volume variable back to zero so the next stock volume can be tracked
            End If
        Next i
        maxVolume = WorksheetFunction.Max(ws.Range("L2:L" + CStr(ws.Cells(Rows.Count, 12).End(xlUp).Row))) 'getting the largest stock volume value from the column of stock volumes
        minPercentChange = WorksheetFunction.Min(ws.Range("K2:K" + CStr(ws.Cells(Rows.Count, 11).End(xlUp).Row))) 'getting the smallest stock percent change value from the column of stock percent change
        maxPercentChange = WorksheetFunction.Max(ws.Range("K2:K" + CStr(ws.Cells(Rows.Count, 11).End(xlUp).Row))) 'getting the largest stock percent change value from the column of stock percent change
        For i = 2 To ws.Cells(Rows.Count, 12).End(xlUp).Row ' loop through the volume column until we find a value that matches the max volume
            If ws.Cells(i, 12) = maxVolume Then
                ws.Range("P4").Value = ws.Cells(i, 9) ' if the volume is a match go to the Ticker column of that same row to retrieve which stock ticker has the highest volume
                ws.Range("Q4").Value = maxVolume ' if the volume is a match then put the volume value into the Value column
                Exit For ' break out of for loop if a volume value match is found
            End If
        Next i
        For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row ' loop through the percent change column until we find a value that matches the min percent change
            If ws.Cells(i, 11) = minPercentChange Then
                ws.Range("P3").Value = ws.Cells(i, 9) ' if the min percent change is a match go to the Ticker column of that same row to retrieve which stock ticker has the lowest percent change
                ws.Range("Q3").Value = minPercentChange ' if the percent change is a match then put the percent change value into the Value column
                ws.Range("Q3").NumberFormat = "0.00%" ' format value to be a percent with two decimal places
                Exit For ' break out of for loop if a percent change value match is found
            End If
        Next i
        For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row ' loop through the percent change column until we find a value that matches the max percent change
            If ws.Cells(i, 11) = maxPercentChange Then
                ws.Range("P2").Value = ws.Cells(i, 9) ' if the max percent change is a match go to the Ticker column of that same row to retrieve which stock ticker has the highest percent change
                ws.Range("Q2").Value = maxPercentChange ' if the percent change is a match then put the percent change value into the Value column
                ws.Range("Q2").NumberFormat = "0.00%" ' format value to be a percent with two decimal places
                Exit For ' break out of for loop if a percent change value match is found
            End If
        Next i
    Next ws
End Sub
