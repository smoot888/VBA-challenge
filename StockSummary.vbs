Sub StockSummary()
    'Iterate through each worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    
        'Initialize Variables
        Dim N As Long
        Dim summaryLocation As Integer
        Dim tickerStartLocation As Long
        Dim tickerEndLocation As Long
        Dim totalVolume As Double
        
        'Initial Variable values
        tickerStartLocation = 2
        tickerEndLocation = 2
        summaryLocation = 2
        totalVolume = 0
        N = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Create new headers in the excel sheet
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        
        'Loop through all of the rows and summarize each ticker
        For i = 2 To N
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                tickerEndLocation = i
                'Print Ticker Name
                ws.Cells(summaryLocation, 9) = ws.Cells(i, 1)
                'Print Yearly Change in dollars
                difference = ws.Cells(tickerEndLocation, 6) - ws.Cells(tickerStartLocation, 3)
                ws.Cells(summaryLocation, 10) = difference
                'Print Yearly Percent Change
                ws.Cells(summaryLocation, 11) = Round((difference / ws.Cells(tickerStartLocation, 3)), 4)
                'Print Total Stock Volume
                For j = tickerStartLocation To tickerEndLocation
                totalVolume = totalVolume + ws.Cells(j, 7)
                Next j
                ws.Cells(summaryLocation, 12) = totalVolume
                'Iterate variables
                summaryLocation = summaryLocation + 1
                tickerStartLocation = i + 1
                totalVolume = 0
            End If
            Next i
        'Conditional Formatting
        Dim rng As Range
        Set rng = ws.Range("J2", "J" & N)
        For Each Cell In rng
            If Cell.Value >= 0 Then
                Cell.Interior.ColorIndex = 4
            ElseIf Cell.Value < 0 Then
                Cell.Interior.ColorIndex = 3
        End If
        Next Cell
        
        Dim rng2 As Range
        Set rng2 = ws.Range("K2", "K" & N)
        For Each Cell In rng2
            If Cell.Value >= 0 Then
                Cell.Interior.ColorIndex = 4
            ElseIf Cell.Value < 0 Then
                Cell.Interior.ColorIndex = 3
        End If
        Next Cell
        
        'Additional summarization
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        
        Dim Max As Integer
        Dim Min As Integer
        Dim Greatest As Integer
        Max = 2
        Min = 2
        Greatest = 2
        For c = 2 To ws.Cells(Rows.Count, "A").End(xlUp).Row
            If ws.Cells(c, 11) > ws.Cells(Max, 11) Then
                Max = c
            End If
            If ws.Cells(c, 11) < ws.Cells(Min, 11) Then
                Min = c
            End If
            If ws.Cells(c, 12) > ws.Cells(Greatest, 12) Then
                Greatest = c
            End If
        Next c
        ws.Range("P2") = ws.Cells(Max, 9)
        ws.Range("Q2") = ws.Cells(Max, 11)
        ws.Range("P3") = ws.Cells(Min, 9)
        ws.Range("Q3") = ws.Cells(Min, 11)
        ws.Range("P4") = ws.Cells(Greatest, 9)
        ws.Range("Q4") = ws.Cells(Greatest, 12)
    Next ws
End Sub
