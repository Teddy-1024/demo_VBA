' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit


Function input_Bool(ByVal bool_str As String) As Boolean
' FUNCTION
    ' Evaluate if String is equivalent to Boolean True
' ARGUMENTS
    ' String bool_str
' VARIABLE DECLARATION
    Dim L As Long
' VARIABLE INSTANTIATION
    bool_str = UCase(bool_str)
    L = Len(bool_str)
    input_Bool = False
' METHODS
    Select Case bool_str
        Case "YES", "Y", "YEAH", "YESH", "YEESH", "YEES", "TRUE", "T":
            input_Bool = True
        Case Else:
            If is_numeric(bool_str) Then
                input_Bool = True
            ElseIf L > 1 Then
                If Left(bool_str, 1) = "_" And is_numeric(Mid(bool_str, 2)) Then
                    input_Bool = True
                End If
            End If
    End Select
End Function


Function input_Long(ByVal long_str As String, Optional Default As Long = 0) As Long
' FUNCTION
    ' Evaluate if String is equivalent to Long
' ARGUMENTS
    ' String long_str
    ' Long default
' METHODS
    If is_numeric(long_str) Then
        input_Long = CLng(long_str)
    Else
        input_Long = Default
    End If
End Function


Function max_Long(ByRef a As Long, ByVal b As Long) As Long
' FUNCTION
    ' Maximum of two Long values
' ARGUMENTS
    ' Long a
    ' Long b
' METHODS
    If a < b Then
        max_Long = b
    Else
        max_Long = a
    End If
End Function


Function min_Long(ByRef a As Long, ByVal b As Long) As Long
' FUNCTION
    ' Maximum of two Long values
' ARGUMENTS
    ' Long a
    ' Long b
' METHODS
    min_Long = -max_Long(-a, -b)
End Function


Function max_Double(ByRef a As Double, ByVal b As Double) As Double
' FUNCTION
    ' Maximum of two Double values
' ARGUMENTS
    ' Double a
    ' Double b
' METHODS
    If a < b Then
        max_Double = b
    Else
        max_Double = a
    End If
End Function


Function min_Double(ByRef a As Double, ByVal b As Double) As Double
' FUNCTION
    ' Maximum of two Double values
' ARGUMENTS
    ' Double a
    ' Double b
' METHODS
    min_Double = -max_Double(-a, -b)
End Function


Function valid_coordinate(ByRef row As Long, ByVal col As Long) As Boolean
' FUNCTION
    ' Is coordinate on Worksheet?
' ARGUMENTS
    ' Long a
    ' Long b
' RETURNS
    valid_coordinate = Not ((row < 1) Or (row > 1048576) Or (col < 1) Or (col > 16384))
End Function


Function valid_range_String(ByVal range_str As String) As Boolean
' FUNCTION
    ' Validate format of range string, e.g. "ABC123:BUD420"
' ARGUMENTS
    ' String range_str
' VARIABLE DECLARATION
    Dim Temps() As String
    Dim Phrases() As String
    Dim i As Long
    Dim N As Long
' VARIABLE INSTANTIATION
    valid_range_String = False
    Temps = Split(range_str, ":")
    Phrases = change_array_base_String(Temps)
    N = SizeArrayDim_String(Phrases)
' METHODS
    If N <> 2 Then
        If (N = 1) Then
            valid_range_String = valid_range_String_segment(range_str)
        End If
        Exit Function
    End If
    valid_range_String = True
    For i = 1 To 2
        valid_range_String = valid_range_String And valid_range_String_segment(Phrases(i))
        If Not valid_range_String Then Exit Function
    Next
End Function


Function valid_range_String_segment(ByVal range_str_segment As String) As Boolean
' FUNCTION
    ' Validate format of segment of range string, e.g. "ABC123"
' ARGUMENTS
    ' String range_str_segment
' VARIABLE DECLARATION
    Dim i As Long
    Dim N As Long
    Dim end_of_lets As Boolean
    Dim temp As String
' VARIABLE INSTANTIATION
    N = Len(range_str_segment)
    end_of_lets = False
    valid_range_String_segment = False
' METHODS
    If N < 2 Then Exit Function
    For i = 1 To N
        temp = Mid(range_str_segment, i, 1)
        If end_of_lets Then
            If Not is_numeric(temp) Then Exit Function
        ElseIf Not is_alphabetic(temp) Then
            If is_numeric(temp) Then
                end_of_lets = True
            Else
                Exit Function
            End If
        End If
    Next
' RETURNS
    valid_range_String_segment = end_of_lets
End Function


Function error_msg(ByVal v As Variant, ByVal name As String, ByVal v_type As String, Optional v_expected As Variant = Nothing) As String
' FUNCTION
    ' Error message string for invalid argument
' ARGUMENTS
    ' Variant v - erroneous argument
    ' String name - name of argument
    ' String v_type - argument data type
    ' Variant v_expected - expected value
' VARIABLE DECLARATION
    Dim str_v As String
    Dim str_exp As String
' VARIABLE INSTANTIATION
    If v Is Nothing Then
        str_v = "Nothing"
    Else
        str_v = CStr(v)
    End If
    If v_expected Is Nothing Then
        str_exp = "Nothing"
    Else
        str_exp = CStr(v_expected)
    End If
' RETURNS
    error_msg = "Invalid " & v_type & " " & name & "." & vbCrLf & "Value = " & str_v & vbCrLf & "Expected value = " & str_exp
End Function
