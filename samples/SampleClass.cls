VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

' ============================================================
' SampleClass.cls
' Purpose: exercises every VB6 token kind for lexer testing.
' ============================================================

Implements IOrderService

' ── Data types ───────────────────────────────────────────────
Private mBool       As Boolean
Private mByte       As Byte
Private mInt        As Integer
Private mLong       As Long
Private mSingle     As Single
Private mDouble     As Double
Private mCurrency   As Currency
Private mDecimal    As Decimal
Private mDate       As Date
Private mString     As String
Private mObject     As Object
Private mVariant    As Variant

' ── Constants ────────────────────────────────────────────────
Private Const MAX_RETRIES   As Integer = 3
Private Const APP_NAME      As String = "SampleApp"
Private Const TAX_RATE      As Double = 0.085
Private Const HEX_MASK      As Long = &HFF00FF       ' hex literal
Private Const OCT_VAL       As Long = &O777           ' octal literal
Private Const CURRENCY_FEE  As Currency = 9.99@       ' @ type suffix

' ── Enum ─────────────────────────────────────────────────────
Public Enum OrderStatus
    osNew = 0
    osPending = 1
    osApproved = 2
    osRejected = 3
End Enum

' ── User-defined Type ─────────────────────────────────────────
Public Type Address
    Street  As String
    City    As String
    Zip     As String
    Country As String
End Type

' ── Module-level state ───────────────────────────────────────
Private mStatus     As OrderStatus
Private mRetries    As Integer
Private mCreatedOn  As Date
Private mAddress    As Address

' ── Properties ───────────────────────────────────────────────

Public Property Get Status() As OrderStatus
    Status = mStatus
End Property

Public Property Let Status(ByVal newStatus As OrderStatus)
    mStatus = newStatus
End Property

Public Property Get Address() As Address
    Address = mAddress
End Property

Public Property Set ObjectRef(ByVal obj As Object)
    Set mObject = obj
End Property

' ── Static local variable example ────────────────────────────

Public Function CallCount() As Long
    Static nCalls As Long
    nCalls = nCalls + 1
    CallCount = nCalls
End Function

' ── All literal types ─────────────────────────────────────────

Public Sub LiteralExamples()
    Dim s       As String
    Dim i       As Integer
    Dim d       As Double
    Dim dt      As Date
    Dim b       As Boolean
    Dim cur     As Currency

    s   = "Hello, ""World"""           ' string literal with escaped quote
    i   = 42                            ' integer literal
    d   = 3.14159                       ' double literal
    d   = 1.5E+10                       ' double with exponent
    dt  = #1990-06-15#                  ' date literal
    b   = True                          ' bool literal
    b   = False                         ' bool literal
    cur = 19.99@                        ' currency literal
    i   = &H1F                          ' hex literal
    i   = &O17                          ' octal literal
End Sub

' ── All operators ─────────────────────────────────────────────

Public Function OperatorExamples(ByVal x As Double, ByVal y As Double) As String
    Dim result  As Double
    Dim s       As String
    Dim b       As Boolean

    ' Arithmetic
    result = x + y
    result = x - y
    result = x * y
    result = x / y
    result = x \ y          ' integer division (backslash)
    result = x ^ y          ' exponentiation (caret)
    result = x Mod y        ' modulo

    ' String concatenation
    s = "Value: " & result

    ' Comparison
    b = x = y               ' equals
    b = x <> y              ' not equal
    b = x < y               ' less than
    b = x > y               ' greater than
    b = x <= y              ' less or equal
    b = x >= y              ' greater or equal

    ' Logical / bitwise
    b = (x > 0) And (y > 0)
    b = (x > 0) Or (y > 0)
    b = Not b
    b = (x > 0) Xor (y > 0)
    b = (x > 0) Eqv (y > 0)
    b = (x > 0) Imp (y > 0)

    OperatorExamples = s
End Function

' ── Control flow: If / ElseIf / Else ─────────────────────────

Public Function Classify(ByVal n As Long) As String
    If n < 0 Then
        Classify = "negative"
    ElseIf n = 0 Then
        Classify = "zero"
    ElseIf n < 100 Then
        Classify = "small"
    Else
        Classify = "large"
    End If
End Function

' ── Control flow: Select Case ─────────────────────────────────

Public Function DescribeStatus(ByVal s As OrderStatus) As String
    Select Case s
        Case osNew
            DescribeStatus = "New"
        Case osPending, osApproved
            DescribeStatus = "In Progress"
        Case osRejected
            DescribeStatus = "Rejected"
        Case Else
            DescribeStatus = "Unknown"
    End Select
End Function

' ── Control flow: For / Next / Step ──────────────────────────

Public Function SumTo(ByVal n As Long) As Long
    Dim i       As Long
    Dim total   As Long

    total = 0
    For i = 1 To n Step 1
        total = total + i
    Next i

    SumTo = total
End Function

' ── Control flow: For Each / In ──────────────────────────────

Public Function JoinStrings(ByVal items As Variant) As String
    Dim item    As Variant
    Dim result  As String

    result = ""
    For Each item In items
        result = result & item & ", "
    Next item

    JoinStrings = result
End Function

' ── Control flow: While / Wend ────────────────────────────────

Public Function CountDown(ByVal start As Integer) As Integer
    Dim n As Integer
    n = start
    While n > 0
        n = n - 1
    Wend
    CountDown = n
End Function

' ── Control flow: Do / Loop / Until ──────────────────────────

Public Function FirstPositive(ByVal items As Variant) As Double
    Dim i   As Integer
    Dim val As Double

    i = 0
    Do
        val = items(i)
        i = i + 1
    Loop Until val > 0 Or i >= 10

    FirstPositive = val
End Function

' ── Control flow: With ────────────────────────────────────────

Public Sub SetAddress(ByVal street As String, ByVal city As String, _
                      ByVal zip As String, ByVal country As String)
    With mAddress
        .Street  = street
        .City    = city
        .Zip     = zip
        .Country = country
    End With
End Sub

' ── GoSub / Return ───────────────────────────────────────────

Public Function GoSubExample(ByVal n As Integer) As String
    Dim msg As String

    If n > 0 Then GoSub Positive Else GoSub NonPositive
    GoSubExample = msg
    Exit Function

Positive:
    msg = "positive"
    Return

NonPositive:
    msg = "non-positive"
    Return
End Function

' ── GoTo / error label ────────────────────────────────────────

Public Sub GoToExample(ByVal n As Integer)
    If n = 0 Then GoTo Done
    mRetries = mRetries + 1
Done:
End Sub

' ── Error handling: On Error GoTo ────────────────────────────

Public Function SafeDivide(ByVal a As Double, ByVal b As Double) As Double
    On Error GoTo ErrHandler

    If b = 0 Then
        Err.Raise vbObjectError + 1, "SampleClass.SafeDivide", "Division by zero"
    End If

    SafeDivide = a / b
    Exit Function

ErrHandler:
    SafeDivide = 0
    Resume Next
End Function

' ── Error handling: On Error Resume Next ─────────────────────

Public Function TryParse(ByVal s As String) As Long
    On Error Resume Next
    TryParse = CLng(s)
    If Err.Number <> 0 Then
        TryParse = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ── Object keywords: Set / New / Nothing / Me / Is / TypeOf ──

Public Sub ObjectExamples()
    Dim obj     As Object
    Dim b       As Boolean

    Set obj = New SampleClass
    b = obj Is Nothing
    b = obj Is Me
    b = TypeOf obj Is SampleClass
    Set obj = Nothing
End Sub

' ── Like operator ─────────────────────────────────────────────

Public Function MatchesPattern(ByVal s As String) As Boolean
    MatchesPattern = s Like "Order_*_[0-9]"
End Function

' ── Optional and ParamArray parameters ───────────────────────

Public Function Greet(ByVal name As String, _
                      Optional ByVal greeting As String = "Hello") As String
    Greet = greeting & ", " & name & "!"
End Function

Public Function SumAll(ParamArray values() As Variant) As Double
    Dim v       As Variant
    Dim total   As Double
    total = 0
    For Each v In values
        total = total + CDbl(v)
    Next v
    SumAll = total
End Function

' ── ByRef parameter ───────────────────────────────────────────

Public Sub Swap(ByRef a As Long, ByRef b As Long)
    Dim tmp As Long
    tmp = a
    a = b
    b = tmp
End Sub

' ── Friend method ─────────────────────────────────────────────

Friend Sub InternalReset()
    mRetries = 0
    mStatus = osNew
End Sub

' ── Inline colon separator ────────────────────────────────────

Public Sub InlineStatements()
    Dim x As Integer : Dim y As Integer
    x = 1 : y = 2 : x = x + y
End Sub

' ── Call statement ────────────────────────────────────────────

Public Sub CallExample()
    Call InternalReset
    Call Swap(mInt, mRetries)
End Sub

' ── Let statement (explicit) ──────────────────────────────────

Public Sub LetExample()
    Let mInt = 99
    Let mString = "explicit let"
End Sub

' ── IOrderService implementation (Implements) ─────────────────

Public Sub IOrderService_Process(ByVal orderId As Long)
    mStatus = osPending
End Sub
