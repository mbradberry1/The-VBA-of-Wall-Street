Sub StockTickerData()

For Each ws In Worksheets

    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"

    
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TickerSymbol As String
    Dim TotalVolume As Double
    TotalVolume = 0
    Dim Summary As Double
    Summary = 2
    Dim StockOpen As Double
    Dim StockClose As Double
    Dim lastrow As Double
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        TickerSymbol = ws.Cells(i, 1).Value
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value

          ws.Range("I" & Summary).Value = TickerSymbol
          ws.Range("L" & Summary).Value = TotalVolume

        TotalVolume = 0

        StockClose = ws.Cells(i, 6)
       
        If StockOpen = 0 Then
            YearlyChange = 0
            PercentChange = 0
        Else:
            YearlyChange = StockClose - StockOpen
            PercentChange = (StockClose - StockOpen) / StockOpen
        End If

    
            ws.Range("J" & Summary).Value = YearlyChange
            ws.Range("K" & Summary).Value = PercentChange
            ws.Range("K" & Summary).Style = "Percent"
            ws.Range("K" & Summary).NumberFormat = "0.00%"

            Summary = Summary + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         StockOpen = ws.Cells(i, 3)


    Else: TotalVolume = TotalVolume + ws.Cells(i, 7).Value

    End If


    Next i


For r = 2 To lastrow

    If ws.Range("J" & r).Value > 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 4

    ElseIf ws.Range("J" & r).Value < 0 Then
        ws.Range("J" & r).Interior.ColorIndex = 3
        
    End If

    Next r
    
    Next ws


End Sub