Sub VBAStocksCh1():    

    'Challenge 1: Summary table of greatest increase, decrease, total
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    Dim lastRow As Long

    lastRow = Cells(Rows.Count, 9).End(xlUp).Row
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To lastRow
        If Cells(i, 11).Value > greatestIncrease Then
            greatestIncrease = Cells(i, 11).Value
            tickerIncrease = Cells(i, 9).Value
        ElseIf Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = Cells(i, 11).Value
            tickerDecrease = Cells(i, 9).Value
        ElseIf Cells(i, 12).Value > greatestVolume Then
            greatestVolume = Cells(i, 12).Value
            tickerVolume = Cells(i, 9).Value
        End If
    Next i
    
    Cells(2, 16).Value = tickerIncrease
    Cells(2, 17).Value = greatestIncrease
    Cells(2, 17).Style = "Percent"
    Cells(3, 16).Value = tickerDecrease
    Cells(3, 17).Value = greatestDecrease
    Cells(3, 17).Style = "Percent"
    Cells(4, 16).Value = tickerVolume
    Cells(4, 17).Value = greatestVolume

End Sub