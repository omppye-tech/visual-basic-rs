Attribute VB_Name = "Calculator"
Declare Function FFIDivide Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long
Declare Function FFIMultiply Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long
Declare Function FFISubtract Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long
Declare Function FFISum Lib "vbrs.dll" (ByVal x As Long, ByVal y As Long) As Long

Function EnsureOperand(ByVal Expression As String) As String
    EnsureOperand = Replace(Replace(Replace(Replace(Expression, "+", ""), "-", ""), "*", ""), "/", "")
    Exit Function
End Function

Function Divide(ByVal x As Integer, ByVal y As Integer) As Integer
    Divide = FFIDivide(x, y)
    Exit Function
End Function

Function Multiply(ByVal x As Integer, ByVal y As Integer) As Integer
    Multiply = FFIMultiply(x, y)
    Exit Function
End Function

Function Subtract(ByVal x As Integer, ByVal y As Integer) As Integer
    Subtract = FFISubtract(x, y)
    Exit Function
End Function

Function Sum(ByVal x As Integer, ByVal y As Integer) As Integer
    Sum = FFISum(x, y)
    Exit Function
End Function
