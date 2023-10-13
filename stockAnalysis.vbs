Sub stockAnalysis()
    
    ' Loop through all sheets
    For Each ws In Worksheets
    
        'Declare variables
        Dim tickerName, tickerLowestChange, tickerMaxStock As String
        Dim summaryTableRow As Integer
        Dim stockVolume, maxStock As Double
        Dim openPrice, closePrice, percentMaxInc, percentMinDec As Double

        'Initialize Variables
        stockVolume = 0
        summaryTableRow = 2
        openPrice = ws.Cells(2, 3).Value
        ws.Range("I1").Value = "Ticker"
        ws.Range("P1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Iterate through all the ticker values in the worksheet
        For rowCounter = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            If ws.Cells(rowCounter, 1).Value <> ws.Cells(rowCounter + 1, 1).Value Then
            
                'Set ticker name and other variables
                tickerName = ws.Cells(rowCounter, 1).Value
                stockVolume = stockVolume + ws.Cells(rowCounter, 7).Value
                closePrice = ws.Cells(rowCounter, 6).Value
                            
                ws.Range("I" & summaryTableRow).Value = tickerName
                ws.Range("J" & summaryTableRow).Value = closePrice - openPrice
                ws.Range("K" & summaryTableRow).Value = (closePrice - openPrice) / openPrice
                ws.Range("L" & summaryTableRow).Value = stockVolume
    
                ' Update variables/counters for the next ticker informaiton
                summaryTableRow = summaryTableRow + 1
                openPrice = ws.Cells(rowCounter + 1, 3).Value
                stockVolume = 0
                
             Else
                stockVolume = stockVolume + ws.Cells(rowCounter, 7).Value
            
            End If
    
    
        Next rowCounter
            
        'Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume" from the newly created summary table        
        percentMaxInc = ws.Cells(2, 11).Value
        percentMinDec = ws.Cells(2, 11).Value
        maxStock = ws.Cells(2, 12).Value
        tickerName = ws.Cells(2, 9).Value
        tickerLowestChange = ws.Cells(2, 9).Value
        tickerMaxStock = ws.Cells(2, 9).Value
    
        For summaryTableIndex = 2 To summaryTableRow - 1
            
            'Greatest % increase or decrease
            If percentMaxInc < ws.Cells(summaryTableIndex, 11).Value Then
               percentMaxInc = ws.Cells(summaryTableIndex, 11).Value
                tickerName = ws.Cells(summaryTableIndex, 9).Value
            
            ElseIf percentMinDec > ws.Cells(summaryTableIndex, 11).Value Then
                percentMinDec = ws.Cells(summaryTableIndex, 11).Value
                tickerLowestChange = ws.Cells(summaryTableIndex, 9).Value
                
            ElseIf maxStock < ws.Cells(summaryTableIndex, 12).Value Then
                maxStock = ws.Cells(summaryTableIndex, 12).Value
                tickerMaxStock = ws.Cells(summaryTableIndex, 9).Value
            
            End If
            
            'Conditional Formating of Yearly and Percentage Changes for each ticker
            If ws.Cells(summaryTableIndex, 10) < 0 Then
                ws.Cells(summaryTableIndex, 10).Interior.ColorIndex = 3
                ws.Cells(summaryTableIndex, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(summaryTableIndex, 10).Interior.ColorIndex = 4
                ws.Cells(summaryTableIndex, 11).Interior.ColorIndex = 4
            End If
            
        Next summaryTableIndex
        
        ws.Range("Q2").Value = percentMaxInc
        ws.Range("P2").Value = tickerName
        ws.Range("Q3").Value = percentMinDec
        ws.Range("P3").Value = tickerLowestChange
        ws.Range("Q4").Value = maxStock
        ws.Range("P4").Value = tickerMaxStock
        
        'Formatting the new summary table
        ws.Columns("I:Q").AutoFit
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ws

End Sub

