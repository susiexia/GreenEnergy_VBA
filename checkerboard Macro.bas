Attribute VB_Name = "Module3"
Sub squareNumber_excercise()
    'Make a list of square numbers
    Worksheets("2.2.3 excercise").Activate
    
    For i = 1 To 10
    Cells(1, i).Value = i * i
    Next i
    'ptintout
    MsgBox (Cells(1, 5).Value)
End Sub

Sub looptest_excercise()
    '2.3.2 nest loop that puts a 1 in each cell from A1 to J10.
    Worksheets("2.3.2 excercise").Activate
    
    For i = 1 To 10
        For J = 1 To 10
            Cells(i, J).Value = i * J
        Next J
    Next i

End Sub

Sub square_checkerboard_excercise()

    '8x8 square cells with a checkerboard pattern(even or odd)
    'use mod fuction
    Worksheets("2.4.2 excercise").Activate
    
    For i = 1 To 20
        For J = 1 To 20
            Cells(i, J).Value = Int(100 * Rnd) + 1
            
            If Cells(i, J).Value Mod 2 = 0 Then
                Cells(i, J).Interior.Color = vbGreen
            ElseIf Cells(i, J).Value Mod 2 = 1 Then
                Cells(i, J).Interior.Color = vbRed
            End If
            
        Next J
    Next i
    
End Sub
