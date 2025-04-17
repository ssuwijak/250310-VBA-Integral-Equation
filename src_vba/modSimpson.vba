Private Function Fx(u As Double, lambda As Double) As Double
    Fx = (u ^ (1 / 7)) * Cosh(Pi4 * u / lambda)
    'Fx = u ^ 3 ' for testing only
End Function

Public Function SumByNumericalMethod(a As Double, b As Double, n As Integer, lambda As Double) As Double
    ' Numerical integration (Simpson's rule)
    ' https://www.youtube.com/watch?v=7EqRRuh-5Lk
    Dim i As Integer
    Dim dx As Double, x As Double, sum As Double
    
    If a = b Then
        SumByNumericalMethod = 0
        Exit Function
    End If
    
    If a > b Then
        x = a
        a = b
        b = a
    End If

    ' Number of intervals (adjust as needed for accuracy)
    If n < 2 Or n > 2000 Then n = 100

    dx = (b - a) / n
    
    'On Error GoTo err1
    
    sum = Fx(a, lambda) + Fx(b, lambda)
   
    For i = 1 To n - 1
        x = a + dx * i
        
        If i Mod 2 = 0 Then
            ' even term
            sum = sum + 2 * Fx(x, lambda)
        Else
            ' odd term
            sum = sum + 4 * Fx(x, lambda)
        End If
    Next
    
    sum = sum * dx / 3
    
    SumByNumericalMethod = sum
    Exit Function
    
err1:
    If Err.Number = 6 Then
        SumByNumericalMethod = CDbl(IIf(sum < 0, -10000, 10000))
    Else
        MsgBox Err.Description, vbCritical
        SumByNumericalMethod = 0
    End If
End Function

Private Sub Test_SumByNumericalMethod()
    ' https://www.youtube.com/watch?v=7EqRRuh-5Lk
    Dim y As Double
    y = SumByNumericalMethod(2, 10, 4, 0) ' Fx = x ^ 3 ; integrate from 2 to 10 and no lambda
    MsgBox "The answer is " & y, vbInformation
End Sub
