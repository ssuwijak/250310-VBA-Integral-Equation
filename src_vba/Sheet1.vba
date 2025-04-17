Public Lambda1 As Double, Lambda2 As Double

Private Sub GetLambdas()
    Dim a As Double, b As Double
    Dim j As Integer
    j = 2
    
    a = CDbl(Sheet1.Cells(j, 2))
    b = CDbl(Sheet1.Cells(j, 3))
    
    If a = 0 And b = 0 Then
        a = 0: b = a + 10
        MsgBox "Invalid lambda interval, so they were corrected to the default values.", vbExclamation
    End If
    
    If a = b Then
        b = a + 10
        MsgBox "Invalid lambda interval, so they were corrected to the default values.", vbExclamation
    End If
    
    If a > b Then
        x = a
        a = b
        b = x
        MsgBox "Invalid lambda interval, so they were swapped.", vbExclamation
    End If
    
    Sheet1.Cells(j, 2) = a
    Sheet1.Cells(j, 3) = b
    
    Lambda1 = a
    Lambda2 = b
End Sub

Sub FindLamda()
    GetLambdas
    FindLambda Lambda1, Lambda2
End Sub

Sub Plot()
    Dim dx As Double, x As Double, y As Double, a As Double, b As Double
    Dim i As Integer, j As Integer, iteration As Integer

    Clear_Plot
    GetLambdas
    a = Lambda1
    b = Lambda2
    
    iteration = 100
    dx = (b - a) / iteration
    
    x = a
    i = 0
    
    On Error Resume Next
    
    For i = 0 To iteration
        x = a + i * dx
        
        j = i + 9
        Sheet1.Cells(j, 1) = i
        Sheet1.Cells(j, 2) = x
        
        y = CalEquation(x)
        
        If Err.Number = 0 Then
            Sheet1.Cells(j, 3) = y
        Else
            Err.Clear
        End If
    Next
    
   ' MsgBox "Plot done", vbInformation
End Sub

Sub Clear_Plot()
    Range("A9:C9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A9").Select
End Sub