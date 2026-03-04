Attribute VB_Name = "SampleModule"

Option Explicit

' ============================================================
' SampleModule.bas
' Purpose: exercises tokens specific to standard modules,
'          plus Global keyword and remaining coverage.
' ============================================================

' ── Global / Public / Private module-level declarations ──────
Global APP_VERSION  As String       ' Global = old-style Public
Public MODULE_NAME  As String
Private mCallCount  As Long

' ── Public Const ─────────────────────────────────────────────
Public Const PI         As Double   = 3.14159265358979
Public Const E          As Double   = 2.71828182845905
Public Const EMPTY_STR  As String   = ""

' ── Public Enum in a module ───────────────────────────────────
Public Enum LogLevel
    llDebug = 0
    llInfo = 1
    llWarning = 2
    llError = 3
End Enum

' ── Public Type in a module ───────────────────────────────────
Public Type Point
    X As Double
    Y As Double
End Type

' ── Initializer (called at startup) ──────────────────────────

Public Sub InitModule()
    APP_VERSION = "1.0.0"
    MODULE_NAME = "SampleModule"
    mCallCount = 0
End Sub

' ── String utility functions ──────────────────────────────────

Public Function Repeat(ByVal s As String, ByVal n As Integer) As String
    Dim i       As Integer
    Dim result  As String
    result = ""
    For i = 1 To n
        result = result & s
    Next i
    Repeat = result
End Function

Public Function IsNullOrEmpty(ByVal s As String) As Boolean
    IsNullOrEmpty = (Len(s) = 0)
End Function

' ── Math utilities ────────────────────────────────────────────

Public Function Clamp(ByVal value As Double, _
                      ByVal minVal As Double, _
                      ByVal maxVal As Double) As Double
    If value < minVal Then
        Clamp = minVal
    ElseIf value > maxVal Then
        Clamp = maxVal
    Else
        Clamp = value
    End If
End Function

Public Function Distance(ByVal p1 As Point, ByVal p2 As Point) As Double
    Dim dx As Double
    Dim dy As Double
    dx = p2.X - p1.X
    dy = p2.Y - p1.Y
    Distance = Sqr(dx ^ 2 + dy ^ 2)
End Function

' ── Bitwise / logical coverage ────────────────────────────────

Public Function BitwiseDemo(ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    r = a And b
    r = a Or b
    r = a Xor b
    r = Not a
    r = a Eqv b
    r = a Imp b
    r = a Mod b
    BitwiseDemo = r
End Function

' ── Comparison operators coverage ────────────────────────────

Public Function CompareDemo(ByVal a As Long, ByVal b As Long) As String
    Dim result As String
    If a = b Then
        result = "equal"
    ElseIf a <> b Then
        result = "not equal"
    ElseIf a < b Then
        result = "less"
    ElseIf a > b Then
        result = "greater"
    ElseIf a <= b Then
        result = "less or equal"
    ElseIf a >= b Then
        result = "greater or equal"
    End If
    CompareDemo = result
End Function

' ── Do While variant ──────────────────────────────────────────

Public Function DigitCount(ByVal n As Long) As Integer
    Dim count As Integer
    count = 0
    Do While n > 0
        n = n \ 10
        count = count + 1
    Loop
    DigitCount = count
End Function

' ── Object: Is Nothing check ─────────────────────────────────

Public Function IsSet(ByVal obj As Object) As Boolean
    IsSet = Not (obj Is Nothing)
End Function

' ── Date literal and Date type ────────────────────────────────

Public Function DaysSinceEpoch(ByVal d As Date) As Long
    Dim epoch As Date
    epoch = #1970-01-01#
    DaysSinceEpoch = DateDiff("d", epoch, d)
End Function

' ── Multi-line expression with line continuation ──────────────

Public Function LongExpression(ByVal a As Double, _
                               ByVal b As Double, _
                               ByVal c As Double) As Double
    LongExpression = (a * b) + _
                     (b * c) + _
                     (a * c)
End Function

' ── Static variable in module procedure ──────────────────────

Public Function NextId() As Long
    Static lastId As Long
    lastId = lastId + 1
    NextId = lastId
End Function

' ── Error handling in a module ───────────────────────────────

Public Function SafeSqrt(ByVal n As Double) As Double
    On Error GoTo ErrHandler
    If n < 0 Then
        Err.Raise vbObjectError + 100, "SampleModule.SafeSqrt", _
                  "Cannot take square root of negative number"
    End If
    SafeSqrt = Sqr(n)
    Exit Function
ErrHandler:
    SafeSqrt = 0
    Resume Next
End Function

' ── Bang (!) operator — Collection-style default member access

Public Sub BangExample()
    Dim col As Object
    Dim val As Variant
    Set col = CreateObject("Scripting.Dictionary")
    col.Add "key", "value"
    val = col!key       ' bang access
    Set col = Nothing
End Sub
