Private Const x_max As Double = 20
Private Const y_scale As Integer = 1000

Public Const Pi1 As Double = 3.14159 'WorksheetFunction.PI
Public Const Pi2 As Double = 2 * Pi1
Public Const Pi4 As Double = 4 * Pi1

Public Function Sinh(x As Double) As Double
    'On Error GoTo err1
    Sinh = WorksheetFunction.Sinh(x)
    Exit Function
err1:
    If x = 0 Then
        Sinh = 0
    ElseIf x > x_max Or x < -x_max Then
        Sinh = y_scale * x
    Else
        Sinh = (Exp(x) - Exp(-x)) / 2
    End If
End Function

Public Function Cosh(x As Double) As Double
    'On Error GoTo err1
    Cosh = Excel.WorksheetFunction.Cosh(x)
    Exit Function
err1:
    If x = 0 Then
        Cosh = 1
    ElseIf x > x_max Or x < -x_max Then
        Cosh = y_scale * x
    Else
        Cosh = (Exp(x) + Exp(-x)) / 2
    End If
End Function

Public Function Tanh(x As Double) As Double
    'On Error GoTo err1
    Tanh = WorksheetFunction.Tanh(x)
    Exit Function
err1:
    If x = 0 Then
        Tanh = 0
    ElseIf x > x_max Then
        Tanh = 1 - 1 / y_scale
    ElseIf x < -x_max Then
        Tanh = -1 + 1 / y_scale
    Else
        Tanh = (Exp(x) - Exp(-x)) / (Exp(x) + Exp(-x))
    End If
End Function
