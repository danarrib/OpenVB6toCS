Attribute VB_Name = "MathHelpers"
Option Explicit

Public Function Clamp(ByVal value As Double, ByVal minVal As Double, ByVal maxVal As Double) As Double
    If value < minVal Then
        Clamp = minVal
    ElseIf value > maxVal Then
        Clamp = maxVal
    Else
        Clamp = value
    End If
End Function

Public Function RoundTo(ByVal value As Double, ByVal decimals As Integer) As Double
    Dim factor As Double
    factor = 10 ^ decimals
    RoundTo = Int(value * factor + 0.5) / factor
End Function
