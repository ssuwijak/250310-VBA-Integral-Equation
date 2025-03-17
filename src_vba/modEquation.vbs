' Given parameters
Public Const T As Double = 10.04
Public Const g As Double = 9.81
Public Const d As Double = 481.18
Public Const Uc0 As Double = 1.33

Private m_lambda As Double
Private msg As String

Private Function Term1()
    Term1 = m_lambda / T
End Function

Private Function Term2()
    Term2 = m_lambda / Sqr(Pi2 * m_lambda / (g * Tanh(Pi2 * d / m_lambda)))
End Function

Private Function Term3()
    Dim integral_val As Double
    integral_val = SumByNumericalMethod(0, d, 1000, m_lambda) / (d ^ (1 / 7))

    Term3 = ((Pi4 / m_lambda) / Sinh(Pi4 * d / m_lambda)) * Uc0 * integral_val
End Function

Public Function CalEquation(lambda As Double) As Double
    If lambda <= 0 Then
        CalEquation = 0 ' pre-define value for identifying the lambda is zero.
        msg = "The lambda cannot be less than zero.\n\nTry to adjust new ones of the lower and/or upper lambda."
        MsgBox Replace(msg, "\n", vbCrLf), vbCritical
        End
        Exit Function
    End If
    
    m_lambda = lambda
    CalEquation = Term1 - Term2 - Term3
End Function

Sub FindLambda(x_lower As Double, x_upper As Double)
    ' Example usage (find root using a simple bisection method)
    ' https://www.youtube.com/watch?v=_M3C0RfowIg
    Dim lambda_lower As Double, lambda_upper As Double, lambda_mid As Double
    Dim f_lower As Double, f_upper As Double, f_mid As Double
    
    ' Set initial bounds for lambda (adjust as needed)
    lambda_lower = x_lower '100
    lambda_upper = x_upper '500
    
    On Error GoTo err1
    ' Evaluate function at bounds
    f_lower = CalEquation(lambda_lower)
    f_upper = CalEquation(lambda_upper)
    
    ' Check if root exists within bounds
    If f_lower * f_upper > 0 Then
        msg = "No root found within the specified lambda range.\n\nTry to adjust the new ones of the lower and/or upper lambda."
        MsgBox Replace(msg, "\n", vbCrLf), vbExclamation
        Exit Sub
    End If
    
    ' Bisection method
    Dim tolerance As Double, iteration_limit As Integer, i As Integer
    tolerance = 0.001 ' Adjust as needed
    iteration_limit = 1000
    i = 0
    
    Do While (lambda_upper - lambda_lower) > tolerance And i < iteration_limit
        lambda_mid = (lambda_lower + lambda_upper) / 2
        f_mid = CalEquation(lambda_mid)
        
        If f_lower * f_mid < 0 Then
            lambda_upper = lambda_mid
            f_upper = f_mid
        Else
            lambda_lower = lambda_mid
            f_lower = f_mid
        End If
        
        i = i + 1
    Loop
    
    ' Display result
    If i < iteration_limit Then
        MsgBox "Lambda = " & lambda_mid, vbInformation
    Else
        MsgBox "No root was found with iteration = " & i, vbExclamation
    End If
    Exit Sub
    
err1:
    msg = Err.Description & "\n\nThe F(x) cannot be calculated with the specified lambda range.\n\nClick the Plot button to find the proper lambda range."
    MsgBox Replace(msg, "\n", vbCrLf), vbCritical
End Sub

