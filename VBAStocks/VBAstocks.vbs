Sub VBAStocks():

    'Summary table for each ticker

    Dim ticker As String
    Dim yearlyChange As Double
    Dim yearOpen As Double
    Dim yearClose As Double
    Dim percentChange As Double
    Dim stockTotal As Double
    Dim lastRow As Long
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    summaryRow = 2
    yearOpen = Cells(2, 3).Value
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    For i = 2 To lastRow
        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
            ticker = Cells(i, 1).Value
            yearClose = Cells(i, 6).Value
            yearlyChange = yearClose - yearOpen
            percentChange = (yearlyChange / yearOpen)
            stockTotal = stockTotal + Cells(i, 7).Value
            Cells(summaryRow, 9).Value = ticker
            Cells(summaryRow, 10).Value = yearlyChange
            Cells(summaryRow, 10).NumberFormat = "0.00"
            If yearlyChange < 0 Then
                Cells(summaryRow, 10).Interior.ColorIndex = 3
            Else
                Cells(summaryRow, 10).Interior.ColorIndex = 4
            End If
            Cells(summaryRow, 11).Value = percentChange
            Cells(summaryRow, 11).Style = "Percent"
            Cells(summaryRow, 12).Value = stockTotal
            yearOpen = Cells(i + 1, 3).Value
            summaryRow = summaryRow + 1
            stockTotal = 0
        Else 
            stockTotal = stockTotal + Cells(i, 7).Value
        End If
    Next i

End Sub