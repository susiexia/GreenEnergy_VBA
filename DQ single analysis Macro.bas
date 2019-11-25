Attribute VB_Name = "Module1"
                                                           Sub MacroCheck()
    Dim textMessage As String
    testMessage = "Hello World"
    MsgBox (testMessage)
    
End Sub

Sub DQAnalysis()
    Worksheets("DQ analysis").Activate
    Range("a1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Range("C3").Value = "Return"
    
    'now! activate and move to sheet"2018"
    'set up variables and initial values
    Worksheets("2018").Activate
    Dim startingPrice As Double
    Dim endingPrice As Double
    TotalVolume = 0
    
    'check the last row number
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    'begin to for loop, using logical operator and conditional statement
    For i = 2 To lastRow
    
        If Cells(i, 1).Value = "DQ" Then
            TotalVolume = TotalVolume + Cells(i, 8).Value
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            startingPrice = Cells(i, 6).Value
        End If
        
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            endingPrice = Cells(i, 6).Value
        End If
    Next i
    'debugging
    'MsgBox (startingPrice)
    'MsgBox (endingPrice)
    
    'now! move back to sheet"dq analysis"
    'fill output into DQ analysis worksheet
    Worksheets("DQ analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = TotalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
    

End Sub
