Attribute VB_Name = "Module4"
Sub ALLStockAnalysis()
    
    yearValue = InputBox("What year would you like to run the all stocks analysis on?")
    
    Worksheets("All Stock Analysis").Activate
    
    'insert user input using concatenation
    Cells(1, 1).Value = "All Stocks (" + yearValue + " )"
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '!!!create arrays for 12 tickers
    Dim tickers(12) As String
    'and assign names for each ticker by array's index(FROM 0!!!)
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'insert user input instead of 2018
    Worksheets(yearValue).Activate
    '!!!dim cannot be inside of loop, will overwrite
    Dim startingPrice As Single
    Dim endingPrice As Single
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'outer loop for through arrays 12 elements, assigned in variable ticker
    For i = 0 To 11
        tickerName = tickers(i)
        TotalVolume = 0
     
        '!!!!!inner loop also need to specific which sheet activated
        Worksheets(yearValue).Activate
        
        For J = 2 To lastRow
    
            If Cells(J, 1).Value = tickerName Then
                TotalVolume = TotalVolume + Cells(J, 8).Value
            End If
            
            If Cells(J, 1).Value = tickerName And Cells(J - 1, 1).Value <> tickerName And Cells(J, 6) <> 0 Then
                startingPrice = Cells(J, 6).Value
            End If
            
            If Cells(J, 1).Value = tickerName And Cells(J + 1, 1).Value <> tickerName Then
                endingPrice = Cells(J, 6).Value
            End If

        Next J
    
        Worksheets("All Stock Analysis").Activate
        Cells(i + 4, 1).Value = tickerName
        Cells(i + 4, 2).Value = TotalVolume
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
    Next i
    
    'formatting
    Worksheets("All Stock Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A1").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("c4:c15").NumberFormat = "0.0%"
    Columns(2).AutoFit
    
    'COLOR conditional formatting
    'use variable name as iterator
    Worksheets("All Stock Analysis").Activate
    dataRowEnd = Cells(Rows.Count, "C").End(xlUp).Row
    dataRowStart = 4
    
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3).Value > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3).Value < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
    Next i
    
End Sub
