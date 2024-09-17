Sub TickerData()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, startRow As Long
    Dim startPrice As Double, endPrice As Double, totalVolume As Double
    Dim percentChange As Double, change As Double
    Dim currentTicker As String, previousTicker As String
    Dim resultRow As Long
    Dim greatestPercentIncrease, greatestPercentDecrease, greatestVolume As Double
    Dim tickerGI, tickerGD, tickerGV As String
    Dim outputColumn As Integer

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        startRow = 2
        previousTicker = ws.Cells(startRow, 1).Value
        startPrice = ws.Cells(startRow, 3).Value
        totalVolume = 0
        outputColumn = 15

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"

        resultRow = 2
        greatestPercentIncrease = -100000
        greatestPercentDecrease = 100000
        greatestVolume = 0

        For i = 2 To lastRow + 1
            If i <= lastRow Then
                currentTicker = ws.Cells(i, 1).Value

                If currentTicker <> previousTicker Or i = lastRow Then
                    If i = lastRow And currentTicker = previousTicker Then
                        totalVolume = totalVolume + ws.Cells(i, 7).Value
                    End If

                    endPrice = ws.Cells(i - 1, 6).Value
                    change = endPrice - startPrice
                     If startPrice <> 0 Then
                        percentChange = (change / startPrice)
                    Else
                        percentChange = 0
                    End If

                     If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        tickerGreatestIncrease = previousTicker
                    End If

                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        tickerGreatestDecrease = previousTicker
                    End If

                    If totalVolume > greatestVolume Then
                        greatestVolume = totalVolume
                        tickerGreatestVolume = previousTicker
                    End If

                    ws.Cells(resultRow, 9).Value = previousTicker
                    ws.Cells(resultRow, 10).Value = change
                    If ws.Cells(resultRow, 10).Value > 0 Then
                        ws.Cells(resultRow, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(resultRow, 10).Value < 0 Then
                        ws.Cells(resultRow, 10).Interior.ColorIndex = 3
                    End If
                    ws.Cells(resultRow, 11).Value = percentChange
                    ws.Cells(resultRow, 11).NumberFormat = "0.00%"
                    ws.Cells(resultRow, 12).Value = totalVolume

                    resultRow = resultRow + 1

                    If i < lastRow Then
                        startPrice = ws.Cells(i, 3).Value
                        totalVolume = 0
                    End If
                End If

                totalVolume = totalVolume + ws.Cells(i, 7).Value
                previousTicker = currentTicker
            Else
                ws.Cells(resultRow, 9).Value = previousTicker
                ws.Cells(resultRow, 10).Value = change
                ws.Cells(resultRow, 11).Value = percentChange
                ws.Cells(resultRow, 12).Value = totalVolume
            End If
        Next i

        ws.Cells(1, outputColumn + 1).Value = "Ticker"
        ws.Cells(1, outputColumn + 2).Value = "Value"

        ws.Cells(2, outputColumn).Value = "Greatest % Increase"
        ws.Cells(2, outputColumn + 1).Value = tickerGreatestIncrease
        ws.Cells(2, outputColumn + 2).Value = greatestIncrease
        ws.Cells(2, outputColumn + 2).NumberFormat = "0.00%"


        ws.Cells(3, outputColumn).Value = "Greatest % Decrease"
        ws.Cells(3, outputColumn + 1).Value = tickerGreatestDecrease
        ws.Cells(3, outputColumn + 2).Value = greatestDecrease
        ws.Cells(3, outputColumn + 2).NumberFormat = "0.00%"


        ws.Cells(4, outputColumn).Value = "Greatest Total Volume"
        ws.Cells(4, outputColumn + 1).Value = tickerGreatestVolume
        ws.Cells(4, outputColumn + 2).Value = greatestVolume

    Next ws
End Sub
