VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' A simple calculator class for testing the translator

Private mLastResult As Double

Public Enum MathOperation
    opAdd = 1
    opSubtract = 2
    opMultiply = 3
    opDivide = 4
End Enum

Public Function Calculate(ByVal a As Double, ByVal b As Double, ByVal op As MathOperation) As Double
    Dim result As Double

    Select Case op
        Case opAdd
            result = a + b
        Case opSubtract
            result = a - b
        Case opMultiply
            result = a * b
        Case opDivide
            If b = 0 Then
                Err.Raise vbObjectError + 1, "Calculator", "Division by zero"
            End If
            result = a / b
    End Select

    mLastResult = result
    Calculate = result
End Function

Public Property Get LastResult() As Double
    LastResult = mLastResult
End Property

Public Sub Reset()
    mLastResult = 0
End Sub
