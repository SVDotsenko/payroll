Attribute VB_Name = "TestHelpers"
' Minimal assertion helpers for the payroll tests.
' Usage: call BeginTests, run the test procedures, then call EndTests to print the summary.
' Results go to the Immediate window (Ctrl+G in the VBA editor).
Option Private Module

Private passedCount As Long
Private failedCount As Long

Sub BeginTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String(60, "-")
    Debug.Print "Tests started " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' Prints the summary and returns the number of failed assertions (0 means all passed).
Function EndTests() As Long
    Debug.Print "Passed: " & passedCount & ", failed: " & failedCount & _
        IIf(failedCount = 0, "  -> OK", "  -> FAILED")
    EndTests = failedCount
End Function

' Exact comparison: use for strings, integers and other non-monetary values.
' Numbers are compared by value regardless of their type (0 equals 0#); a number never equals a string.
Sub AssertEqual(ByVal expected, ByVal actual, ByVal testName As String)
    On Error GoTo NotComparable
    If IsNumber(expected) And IsNumber(actual) Then
        isEqual = (expected = actual)
    ElseIf VarType(expected) = VarType(actual) And Not IsObject(expected) And Not IsArray(expected) Then
        isEqual = (expected = actual)
    Else
        isEqual = False
    End If

    If isEqual Then
        RecordPass
    Else
        RecordFail testName, "expected " & Describe(expected) & ", got " & Describe(actual)
    End If
    Exit Sub

NotComparable:
    RecordFail testName, "values cannot be compared: expected " & Describe(expected) & ", got " & Describe(actual)
End Sub

' Comparison with a tolerance: use for money and other Double results.
' The default tolerance of half a cent accepts any value that rounds to the same cent.
Sub AssertNear(ByVal expected As Double, ByVal actual, ByVal testName As String, _
               Optional ByVal tolerance As Double = 0.005)
    If Not IsNumber(actual) Then
        RecordFail testName, "expected " & Describe(expected) & ", got non-numeric " & Describe(actual)
    ElseIf Abs(expected - actual) <= tolerance Then
        RecordPass
    Else
        RecordFail testName, "expected " & Describe(expected) & ", got " & Describe(actual) & _
            " (tolerance " & tolerance & ")"
    End If
End Sub

Sub AssertTrue(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        RecordPass
    Else
        RecordFail testName, "condition is False"
    End If
End Sub

Private Sub RecordPass()
    passedCount = passedCount + 1
End Sub

Private Sub RecordFail(ByVal testName As String, ByVal message As String)
    failedCount = failedCount + 1
    Debug.Print "FAIL: " & testName & ": " & message
End Sub

' True only for real numeric types; unlike IsNumeric, rejects numeric strings and Empty.
Private Function IsNumber(ByVal value) As Boolean
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, 20 ' 20 = vbLongLong (64-bit only)
            IsNumber = True
    End Select
End Function

' Renders a value for failure messages so that, for example, "12" and 12 are distinguishable.
Private Function Describe(ByVal value) As String
    If IsObject(value) Then
        Describe = "<" & TypeName(value) & ">"
    ElseIf IsArray(value) Then
        Describe = "<array>"
    ElseIf IsEmpty(value) Then
        Describe = "<Empty>"
    ElseIf IsNull(value) Then
        Describe = "<Null>"
    ElseIf IsError(value) Then
        Describe = "<" & CStr(value) & ">"
    ElseIf VarType(value) = vbString Then
        Describe = """" & value & """"
    Else
        Describe = CStr(value) & " (" & TypeName(value) & ")"
    End If
End Function
