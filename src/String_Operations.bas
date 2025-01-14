' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit


Function is_numeric(ByVal numeric_chars As String) As Boolean
' FUNCTION
    ' Evaluate if all characters in numeric_chars are numeric
' ARGUMENTS
    ' String numeric_chars
' VARIABLE DECLARATION
    Dim i As Long
    Dim asci_i As Long
    Dim valid As Boolean
' ARGUMENT VALIDATION
    If numeric_chars = "" Then Exit Function
' VARIABLE INSTANTIATION
    valid = True
' METHODS
    For i = 1 To Len(numeric_chars)
        asci_i = Asc(Mid(numeric_chars, i, 1))
        valid = valid And (asci_i >= 48 And asci_i <= 57)
    Next
' RETURNS
    is_numeric = valid
End Function

Function is_alphabetic(ByVal alphabetic_chars As String) As Boolean
' FUNCTION
    ' Evaluate if all characters in alphabetic_chars are alpabetic
' ARGUMENTS
    ' String numeric_chars
' VARIABLE DECLARATION
    Dim i As Long
    Dim asci_i As Long
    Dim valid As Boolean
' ARGUMENT VALIDATION
    If alphabetic_chars = "" Then Exit Function
' VARIABLE INSTANTIATION
    valid = True
' METHODS
    For i = 1 To Len(alphabetic_chars)
        asci_i = Asc(Mid(alphabetic_chars, i, 1))
        valid = valid And ((asci_i >= 65 And asci_i <= 90) Or (asci_i >= 97 And asci_i <= 122))
    Next
' RETURNS
    is_alphabetic = valid
End Function

Function get_col_str(ByVal col_ID As Long) As String
' FUNCTION
    ' Return column ID as String from Long
' ARGUMENTS
    ' Long col_ID
' VARIABLE DECLARATION
    Dim strID As String
    Dim i As Double
    Dim N As Double
    Dim temp As Long
    Dim remainder As Double
' ARGUMENT VALIDATION
    If col_ID < 1 Then Exit Function
' VARIABLE INSTANTIATION
    N = max_Double(0, Div_Double(Log(col_ID), Log(26)))
    If Not Mod_Double(Log(col_ID), Log(26)) = 0# Or col_ID = 1 Then N = N + 1
' METHODS
    For i = 1 To N
        temp = ((col_ID - 1) Mod (26 ^ i)) \ (26 ^ (i - 1)) + 1
        strID = Chr(64 + temp) & strID
        col_ID = col_ID - temp * (26 ^ (i - 1))
    Next
' RETURNS
    get_col_str = strID
End Function

Function get_col_ID(ByVal col_str As String) As Long
' FUNCTION
    ' Return column ID as Long from String
' ARGUMENTS
    ' String col_ID
' VARIABLE DECLARATION
    Dim col_ID As Long
    Dim i As Long
    Dim N As Long
    Dim temp As Long
    Dim remainder As Double
' ARGUMENT VALIDATION
    If Not is_alphabetic(col_str) Then Exit Function
' VARIABLE INSTANTIATION
    N = Len(col_str)
' METHODS
    For i = 1 To N
        col_ID = col_ID + (Asc(Mid(col_str, i, 1)) - 64) * 26 ^ (N - i)
    Next
' RETURNS
    get_col_ID = col_ID
End Function

Function Range_String_Coords(ByVal range_str As String) As Long()
' FUNCTION
    ' Array of coordinates in range_str (row_1, column_1, [row_2, column_2])
' ARGUMENTS
    ' String range_str
' VARIABLE DECLARATION
    Dim coords() As Long
    Dim Phrases() As String
    Dim Tmps() As String
    Dim i As Long
    Dim N As Long
' ARGUMENT VALIDATION
    If Not valid_range_String(range_str) Then Exit Function
' VARIABLE INSTANTIATION
    Tmps = Split(range_str, ":")
    Phrases = change_array_base_String(Tmps)
    N = SizeArrayDim_String(Phrases)
    ReDim coords(N)
' METHODS
    For i = 1 To N
        If i Mod 2 = 0 Then
            coords(i) = get_col_ID(Phrases(i))
        Else
            coords(i) = CLng(Phrases(i))
        End If
    Next
End Function

Function CMoney_String(ByVal money_D As Double) As String
' FUNCTION
    ' Get string equivalent of double to 2 decimal places
' ARGUMENTS
    ' Double money_D
' VARIABLE DECLARATION
    ' Dim money_L As Long
    Dim money_S As String
    Dim iDot As Long
' VARIABLE INSTANTIATION
    ' money_D = 100 * money_D
    ' money_L = money_D
    ' money_D = money_L / 100
    money_S = CStr(Round(money_D, 2))
    iDot = InStr(1, money_S, ".")
' RETURNS
    If iDot < 1 Then
        CMoney_String = "'" & money_S & ".00"
    ElseIf iDot = Len(money_S) - 1 Then
        CMoney_String = "'" & money_S & "0"
    Else
        CMoney_String = "'" & money_S
    End If
End Function

Function Path2Name(ByVal FilePath As String) As String
' FUNCTION
    ' Get file name from path
' ARGUMENTS
    ' String FilePath
' VARIABLE DECLARATION
    Dim iSlash As Long
' VARIABLE INSTANTIATION
    Path2Name = FilePath
    iSlash = InStr(1, Path2Name, "/")
    If iSlash < 1 Then iSlash = InStr(1, Path2Name, "/")
' METHODS
    Do While iSlash > 0
        Path2Name = Mid(Path2Name, iSlash + 1)
        iSlash = InStr(1, Path2Name, "/")
        If iSlash < 1 Then iSlash = InStr(1, Path2Name, "/")
    Loop
End Function
