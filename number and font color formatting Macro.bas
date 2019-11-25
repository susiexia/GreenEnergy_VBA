Attribute VB_Name = "Module2"
Sub FormatChange_AllStocksAnalysis()

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

Sub ClearWorksheet()
    Cells.Clear

End Sub
