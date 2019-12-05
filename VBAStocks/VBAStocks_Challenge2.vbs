Sub VBAStocksCh2():

    'Challenge 2: Run VBA script across all worksheets

    Dim ticker As String
    Dim yearlyChange As Double
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim percentChange As Double
    Dim stockTotal As Double
    Dim lastRow As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    For Each ws In Worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        yearOpen = ws.Cells(2, 3).Value
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        For i = 2 To lastRow
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
                ticker = ws.Cells(i, 1).Value
                yearClose = ws.Cells(i, 6).Value
                yearlyChange = yearClose - yearOpen
                If yearOpen = 0 Then
                    percentChange = 0
                Else
                    percentChange = (yearlyChange / yearOpen)
                End If
                stockTotal = stockTotal + ws.Cells(i, 7).Value
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = yearlyChange
                ws.Cells(summaryRow, 10).NumberFormat = "0.00"
                If yearlyChange < 0 Then
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                End If
                ws.Cells(summaryRow, 11).Value = percentChange
                ws.Cells(summaryRow, 11).Style = "Percent"
                ws.Cells(summaryRow, 12).Value = stockTotal
                yearOpen = ws.Cells(i + 1, 3).Value
                summaryRow = summaryRow + 1
                stockTotal = 0
            Else 
                stockTotal = stockTotal + ws.Cells(i, 7).Value
            End If
        Next i

        For i = 2 To lastRow
            If ws.Cells(i, 11).Value > greatestIncrease Then
                greatestIncrease = ws.Cells(i, 11).Value
                tickerIncrease = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11).Value < greatestDecrease Then
                greatestDecrease = ws.Cells(i, 11).Value
                tickerDecrease = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 12).Value > greatestVolume Then
                greatestVolume = ws.Cells(i, 12).Value
                tickerVolume = ws.Cells(i, 9).Value
            End If
        Next i
    
        ws.Cells(2, 16).Value = tickerIncrease
        ws.Cells(2, 17).Value = greatestIncrease
        ws.Cells(2, 17).Style = "Percent"
        ws.Cells(3, 16).Value = tickerDecrease
        ws.Cells(3, 17).Value = greatestDecrease
        ws.Cells(3, 17).Style = "Percent"
        ws.Cells(4, 16).Value = tickerVolume
        ws.Cells(4, 17).Value = greatestVolume
    
        ws.Cells.EntireColumn.AutoFit

    Next ws

End Sub