' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit


Function Range_1D_String(ByRef ws As Worksheet, ByVal range_str As String) As String()
' FUNCTION
    ' Get range of worksheet as 1D String array
' ARGUMENTS
    ' Worksheet ws
    ' String range_str
' VARIABLE DECLARATION
    Dim my_range() As Variant
    Dim sz_x As Long
    Dim sz_y As Long
    Dim i As Long
    Dim N As Long
    Dim strs() As String
    Dim out_err() As String
' ARGUMENT VALIDATION
    ReDim out_err(1)
    out_err(1) = "Error"
    If Not valid_range_String(range_str) Or ws Is Nothing Then
        Range_1D_String = out_err
        Exit Function
    End If
' VARIABLE INSTANTIATION
    my_range = ws.Range(range_str).value
    sz_x = SizeArrayDim_Variant(my_range, 2)
    sz_y = SizeArrayDim_Variant(my_range, 1)
    If (sz_x = 0 Or sz_y = 0 Or (sz_x > 1 And sz_y > 1)) Then
        Range_1D_String = out_err
        Exit Function
    End If
    N = max_Long(sz_x, sz_y)
    ReDim strs(N)
' METHODS
    For i = 1 To N
        If sz_x > sz_y Then
            strs(i) = my_range(1, i)
        Else
            strs(i) = my_range(i, 1)
        End If
    Next
' RETURNS
    Range_1D_String = strs
End Function

Function Range_String(ByVal col_min As Long, ByVal col_max As Long, ByVal row_min As Long, ByVal row_max As Long) As String
' FUNCTION
    ' Create range string from minimum and maximum positions
' ARGUMENTS
    ' Long col_min
    ' Long col_max
    ' Long row_min
    ' Long row_max
' ARGUMENT VALIDATION
    Range_String = "Error: Invalid coordinates"
    If row_max < 1 Then row_max = 1048576
    If col_max < 1 Then col_max = 16384
    If Not (valid_coordinate(row_min, col_min) And valid_coordinate(row_max, col_max)) Then Exit Function
' RETURNS
    Range_String = get_col_str(col_min) & CStr(row_min) & ":" & get_col_str(col_max) & CStr(row_max)
End Function
