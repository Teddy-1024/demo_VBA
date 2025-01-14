' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit


Sub ReDimPreserve_String(ByRef arr() As String, ByVal dimension As Long, ByVal newbound As Long, ByVal NDim As Long)
' FUNCTION
    ' Redimensionalise array without losing data (unless shrinking)
' ARGUMENTS
    ' String Array arr
    ' Long dimension - dimension of arr to change
    ' Long newbound - new size of dimension
    ' Long ndim - number of dimensions of arr
' VARIABLE DECLARATION
    Dim x() As Long ' Iterables for each dimension
    Dim N() As Long ' Size of each dimension
    Dim i As Long
    Dim j As Long
    Dim iterate As Boolean
    Dim Outs() As String
    Dim Nold() As Long
    Dim minbound As Long
' VARIABLE INSTANTIATION
    ReDim x(NDim)
    ReDim N(NDim)
    iterate = True
' METHODS
    ' Populate N
    For i = 1 To NDim
        x(i) = 1
        N(i) = SizeArrayDim_String(arr, i)
    Next
    Nold = N
    N(dimension) = newbound
    minbound = min_Long(newbound, Nold(dimension))
    ' Redimensionalise outputs
    Select Case NDim
    Case 1
        ReDim Outs(N(1))
    Case 2
        ReDim Outs(N(1), N(2))
    Case 3
        ReDim Outs(N(1), N(2), N(3))
    Case 4
        ReDim Outs(N(1), N(2), N(3), N(4))
    Case 5
        ReDim Outs(N(1), N(2), N(3), N(4), N(5))
    Case Else
        MsgBox "Too many dimensions"
        Exit Sub
    End Select
    ' Fill values
    Do While iterate
        ' Fill value
        Select Case NDim
        Case 1
            Outs(x(1)) = arr(x(1))
        Case 2
            Outs(x(1), x(2)) = arr(x(1), x(2))
        Case 3
            Outs(x(1), x(2), x(3)) = arr(x(1), x(2), x(3))
        Case 4
            Outs(x(1), x(2), x(3), x(4)) = arr(x(1), x(2), x(3), x(4))
        Case 5
            Outs(x(1), x(2), x(3), x(4), x(5)) = arr(x(1), x(2), x(3), x(4), x(5))
        Case Else
            MsgBox "Too many dimensions"
            Exit Sub
        End Select
        ' Iterate position
        For i = NDim To 1 Step -1
            If (i = dimension) Then
                If (x(dimension) < minbound) Then
                    x(dimension) = x(dimension) + 1
                    If (dimension < NDim) Then
                        For j = dimension + 1 To NDim
                            x(j) = 1
                        Next
                    End If
                    Exit For
                End If
            ElseIf (x(i) < N(i)) Then
                x(i) = x(i) + 1
                If (i < NDim) Then
                    For j = i + 1 To NDim
                        x(j) = 1
                    Next
                End If
                Exit For
            End If
            If (i = 1) Then
                iterate = False
            End If
        Next
    Loop
' RETURNS
    arr = Outs
End Sub


Sub ReDimPreserve_Long(ByRef arr() As Long, ByVal dimension As Long, ByVal newbound As Long, ByVal NDim As Long)
' FUNCTION
    ' Redimensionalise array without losing data (unless shrinking)
' ARGUMENTS
    ' Long Array arr
    ' Long dimension - dimension of arr to change
    ' Long newbound - new size of dimension
    ' Long ndim - number of dimensions of arr
' VARIABLE DECLARATION
    Dim x() As Long ' Iterables for each dimension
    Dim N() As Long ' Size of each dimension
    Dim i As Long
    Dim j As Long
    Dim iterate As Boolean
    Dim Outs() As Long
    Dim Nold() As Long
    Dim minbound As Long
' VARIABLE INSTANTIATION
    ReDim x(NDim)
    ReDim N(NDim)
    iterate = True
' METHODS
    ' Populate N
    For i = 1 To NDim
        x(i) = 1
        N(i) = SizeArrayDim_Long(arr, i)
    Next
    Nold = N
    N(dimension) = newbound
    minbound = min_Long(newbound, Nold(dimension))
    ' Redimensionalise outputs
    Select Case NDim
    Case 1
        ReDim Outs(N(1))
    Case 2
        ReDim Outs(N(1), N(2))
    Case 3
        ReDim Outs(N(1), N(2), N(3))
    Case 4
        ReDim Outs(N(1), N(2), N(3), N(4))
    Case 5
        ReDim Outs(N(1), N(2), N(3), N(4), N(5))
    Case Else
        MsgBox "Too many dimensions"
        Exit Sub
    End Select
    ' Fill values
    Do While iterate
        ' Fill value
        Select Case NDim
        Case 1
            Outs(x(1)) = arr(x(1))
        Case 2
            Outs(x(1), x(2)) = arr(x(1), x(2))
        Case 3
            Outs(x(1), x(2), x(3)) = arr(x(1), x(2), x(3))
        Case 4
            Outs(x(1), x(2), x(3), x(4)) = arr(x(1), x(2), x(3), x(4))
        Case 5
            Outs(x(1), x(2), x(3), x(4), x(5)) = arr(x(1), x(2), x(3), x(4), x(5))
        Case Else
            MsgBox "Too many dimensions"
            Exit Sub
        End Select
        ' Iterate position
        For i = NDim To 1 Step -1
            If (i = dimension) Then
                If (x(dimension) < minbound) Then
                    x(dimension) = x(dimension) + 1
                    If (dimension < NDim) Then
                        For j = dimension + 1 To NDim
                            x(j) = 1
                        Next
                    End If
                    Exit For
                End If
            ElseIf (x(i) < N(i)) Then
                x(i) = x(i) + 1
                If (i < NDim) Then
                    For j = i + 1 To NDim
                        x(j) = 1
                    Next
                End If
                Exit For
            End If
            If (i = 1) Then
                iterate = False
            End If
        Next
    Loop
' RETURNS
    arr = Outs
End Sub


Sub ReDimPreserve_Variant(ByRef arr() As Variant, ByVal dimension As Long, ByVal newbound As Long, ByVal NDim As Long)
' FUNCTION
    ' Redimensionalise array without losing data (unless shrinking)
' ARGUMENTS
    ' Variant Array arr
    ' Long dimension - dimension of arr to change
    ' Long newbound - new size of dimension
    ' Long ndim - number of dimensions of arr
' VARIABLE DECLARATION
    Dim x() As Long ' Iterables for each dimension
    Dim N() As Long ' Size of each dimension
    Dim i As Long
    Dim j As Long
    Dim iterate As Boolean
    Dim Outs() As Variant
    Dim Nold() As Long
    Dim minbound As Long
' VARIABLE INSTANTIATION
    ReDim x(NDim)
    ReDim N(NDim)
    iterate = True
' METHODS
    ' Populate N
    For i = 1 To NDim
        x(i) = 1
        N(i) = SizeArrayDim_Variant(arr, i)
    Next
    Nold = N
    N(dimension) = newbound
    minbound = min_Long(newbound, Nold(dimension))
    ' Redimensionalise outputs
    Select Case NDim
    Case 1
        ReDim Outs(N(1))
    Case 2
        ReDim Outs(N(1), N(2))
    Case 3
        ReDim Outs(N(1), N(2), N(3))
    Case 4
        ReDim Outs(N(1), N(2), N(3), N(4))
    Case 5
        ReDim Outs(N(1), N(2), N(3), N(4), N(5))
    Case Else
        MsgBox "Too many dimensions"
        Exit Sub
    End Select
    ' Fill values
    Do While iterate
        ' Fill value
        Select Case NDim
        Case 1
            Outs(x(1)) = arr(x(1))
        Case 2
            Outs(x(1), x(2)) = arr(x(1), x(2))
        Case 3
            Outs(x(1), x(2), x(3)) = arr(x(1), x(2), x(3))
        Case 4
            Outs(x(1), x(2), x(3), x(4)) = arr(x(1), x(2), x(3), x(4))
        Case 5
            Outs(x(1), x(2), x(3), x(4), x(5)) = arr(x(1), x(2), x(3), x(4), x(5))
        Case Else
            MsgBox "Too many dimensions"
            Exit Sub
        End Select
        ' Iterate position
        For i = NDim To 1 Step -1
            If (i = dimension) Then
                If (x(dimension) < minbound) Then
                    x(dimension) = x(dimension) + 1
                    If (dimension < NDim) Then
                        For j = dimension + 1 To NDim
                            x(j) = 1
                        Next
                    End If
                    Exit For
                End If
            ElseIf (x(i) < N(i)) Then
                x(i) = x(i) + 1
                If (i < NDim) Then
                    For j = i + 1 To NDim
                        x(j) = 1
                    Next
                End If
                Exit For
            End If
            If (i = 1) Then
                iterate = False
            End If
        Next
    Loop
' RETURNS
    arr = Outs
End Sub


Function SizeArrayDim_String(ByRef arr() As String, Optional dimension As Long = 1) As Long
' FUNCTION
    ' Find size of dimension of arr
' VARIABLE INSTANTIATION
    dimension = max_Long(1, dimension)
' METHODS
    On Error GoTo errhand
    If Not ((Not arr) = -1) Then
        SizeArrayDim_String = UBound(arr, dimension) - LBound(arr, dimension) + 1
    Else
        SizeArrayDim_String = 0
    End If
    Exit Function
' ERROR HANDLING
errhand:
    SizeArrayDim_String = 0
End Function


Function SizeArrayDim_Long(ByRef arr() As Long, Optional dimension As Long = 1) As Long
' FUNCTION
    ' Find size of dimension of arr
' VARIABLE INSTANTIATION
    dimension = max_Long(1, dimension)
' METHODS
    On Error GoTo errhand
    If Not ((Not arr) = -1) Then
        SizeArrayDim_Long = UBound(arr, dimension) - LBound(arr, dimension) + 1
    Else
        SizeArrayDim_Long = 0
    End If
    Exit Function
' ERROR HANDLING
errhand:
    SizeArrayDim_Long = 0
End Function


Function SizeArrayDim_Variant(ByRef arr() As Variant, Optional dimension As Long = 1) As Long
' FUNCTION
    ' Find size of dimension of arr
' VARIABLE INSTANTIATION
    dimension = max_Long(1, dimension)
' METHODS
    On Error GoTo errhand
    If Not ((Not arr) = -1) Then
        SizeArrayDim_Variant = UBound(arr, dimension) - LBound(arr, dimension) + 1
    Else
        SizeArrayDim_Variant = 0
    End If
    Exit Function
' ERROR HANDLING
errhand:
    SizeArrayDim_Variant = 0
End Function


Function SizeArrayDim_Variant_0(ByVal arr As Variant) As Long
' FUNCTION
    ' Find size of dimension of arr
' METHODS
    On Error GoTo errhand
    If Not IsEmpty(arr) Then ' ((Not arr) = -1) Then
        SizeArrayDim_Variant_0 = UBound(arr) - LBound(arr) + 1
    Else
        SizeArrayDim_Variant_0 = 0
    End If
    Exit Function
' ERROR HANDLING
errhand:
    SizeArrayDim_Variant_0 = 0
End Function


Function create_1D_mat_Boolean(Optional value As Boolean = False, Optional N As Long = 1) As Boolean()
' FUNCTION
    ' Create 1D matrix (array) of size N, type Boolean and value
' ARGUMENTS
    ' Boolean value - for each element of array
    ' Long N - number of elements in array
' PROCESSING ACCELERATION
' CONSTANTS
' VARIABLE DECLARATION
    Dim Outs() As Boolean
    Dim i As Long
' ARGUMENT VALIDATION
    N = max_Long(1, N)
' VARIABLE INSTANTIATION
    ReDim Outs(N)
' METHODS
    For i = 1 To N
        Outs(i) = value
    Next
' RETURNS
    create_1D_mat_Boolean = Outs
End Function


Function create_1D_mat_Long(Optional value As Long = 0, Optional N As Long = 1) As Long()
' FUNCTION
    ' Create 1D matrix (array) of size N, type Long and value
' ARGUMENTS
    ' Long value - for each element of array
    ' Long N - number of elements in array
' PROCESSING ACCELERATION
' CONSTANTS
' VARIABLE DECLARATION
    Dim Outs() As Long
    Dim i As Long
' ARGUMENT VALIDATION
    N = max_Long(1, N)
' VARIABLE INSTANTIATION
    ReDim Outs(N)
' METHODS
    For i = 1 To N
        Outs(i) = value
    Next
' RETURNS
    create_1D_mat_Long = Outs
End Function


Function create_1D_mat_String(Optional value As String = False, Optional N As Long = 1) As String()
' FUNCTION
    ' Create 1D matrix (array) of size N, type String and value
' ARGUMENTS
    ' String value - for each element of array
    ' Long N - number of elements in array
' PROCESSING ACCELERATION
' CONSTANTS
' VARIABLE DECLARATION
    Dim Outs() As String
    Dim i As Long
' ARGUMENT VALIDATION
    N = max_Long(1, N)
' VARIABLE INSTANTIATION
    ReDim Outs(N)
' METHODS
    For i = 1 To N
        Outs(i) = value
    Next
' RETURNS
    create_1D_mat_String = Outs
End Function


Sub copy_N_mat_String(ByRef data_in() As String, ByRef data_out() As String, ByVal N_in As Long, ByVal N_out As Long)
' FUNCTION
    ' Copy from one N-dimensional matrix into another all overlapping elements
' ARGUMENTS
    ' String Matrix data_in - matrix to copy from
    ' String Matrix data_out - matrix receiving data
    ' Long N_in - number of dimensions in data_in
    ' Long N_out - number of dimensions in N_out
' VARIABLE DECLARATION
    Dim dims_in() As Long
    Dim dims_out() As Long
    Dim i As Long
    Dim dim_min As Long
    Dim x() As Long
    Dim N() As Long
' ARGUMENT VALIDATION
    If (N_in < 1 Or N_out <> N_in) Then Exit Sub
' VARIABLE INSTANTIATION
    dim_min = min_Long(N_in, N_out)
    ReDim dims_in(dim_min)
    ReDim dims_out(dim_min)
    ReDim x(dim_min)
    ReDim N(dim_min)
    For i = 1 To dim_min
        dims_in(i) = SizeArrayDim_String(data_in, i)
        dims_out(i) = SizeArrayDim_String(data_out, i)
        x(i) = 1
        N(i) = min_Long(dims_in(i), dims_out(i))
    Next
' METHODS
    Do While compare_all_iterators(x, N, dim_min)
        set_index_N_mat_String data_out, x, dim_min, get_index_N_mat_String(data_in, x, N_out)
        iterate_iterator x, N, dim_min
    Loop
End Sub


Sub copy_N_mat_Long(ByRef data_in() As Long, ByRef data_out() As Long, ByVal N_in As Long, ByVal N_out As Long)
' FUNCTION
    ' Copy from one N-dimensional matrix into another all overlapping elements
' ARGUMENTS
    ' Long Matrix data_in - matrix to copy from
    ' Long Matrix data_out - matrix receiving data
    ' Long N_in - number of dimensions in data_in
    ' Long N_out - number of dimensions in N_out
' VARIABLE DECLARATION
    Dim dims_in() As Long
    Dim dims_out() As Long
    Dim i As Long
    Dim dim_min As Long
    Dim x() As Long
    Dim N() As Long
' ARGUMENT VALIDATION
    If (N_in < 1 Or N_out <> N_in) Then Exit Sub
' VARIABLE INSTANTIATION
    dim_min = min_Long(N_in, N_out)
    ReDim dims_in(dim_min)
    ReDim dims_out(dim_min)
    ReDim x(dim_min)
    ReDim N(dim_min)
    For i = 1 To dim_min
        dims_in(i) = SizeArrayDim_Long(data_in, i)
        dims_out(i) = SizeArrayDim_Long(data_out, i)
        x(i) = 1
        N(i) = min_Long(dims_in(i), dims_out(i))
    Next
' METHODS
    Do While compare_all_iterators(x, N, dim_min)
        set_index_N_mat_Long data_out, x, dim_min, get_index_N_mat_Long(data_in, x, N_out)
        iterate_iterator x, N, dim_min
    Loop
End Sub


Sub copy_N_mat_Variant(ByRef data_in() As Variant, ByRef data_out() As Variant, ByVal N_in As Long, ByVal N_out As Long)
' FUNCTION
    ' Copy from one N-dimensional matrix into another all overlapping elements
' ARGUMENTS
    ' Variant Matrix data_in - matrix to copy from
    ' Variant Matrix data_out - matrix receiving data
    ' Long N_in - number of dimensions in data_in
    ' Long N_out - number of dimensions in N_out
' VARIABLE DECLARATION
    Dim dims_in() As Long
    Dim dims_out() As Long
    Dim i As Long
    Dim dim_min As Long
    Dim x() As Long
    Dim N() As Long
' ARGUMENT VALIDATION
    If (N_in < 1 Or N_out <> N_in) Then Exit Sub
' VARIABLE INSTANTIATION
    dim_min = min_Long(N_in, N_out)
    ReDim dims_in(dim_min)
    ReDim dims_out(dim_min)
    ReDim x(dim_min)
    ReDim N(dim_min)
    For i = 1 To dim_min
        dims_in(i) = SizeArrayDim_Variant(data_in, i)
        dims_out(i) = SizeArrayDim_Variant(data_out, i)
        x(i) = 1
        N(i) = min_Long(dims_in(i), dims_out(i))
    Next
' METHODS
    Do While compare_all_iterators(x, N, dim_min)
        set_index_N_mat_Variant data_out, x, dim_min, get_index_N_mat_Variant(data_in, x, N_out)
        iterate_iterator x, N, dim_min
    Loop
End Sub


Sub iterate_iterator(ByRef x() As Long, ByRef N() As Long, ByVal NDim As Long)
' FUNCTION
    ' Increment iterator x under limits N
' ARGUMENTS
    ' Long Array x - iterator
    ' Long Array SzDim1 - iterator limits
    ' Long NDim
' VARIABLE DECLARATION
    Dim i As Long
    Dim j As Long
' METHODS
    For i = NDim To 1 Step -1
        If (x(i) < N(i)) Then
            x(i) = x(i) + 1
            If (i < NDim) Then
                For j = i + 1 To NDim
                    x(j) = 1
                Next
            End If
            Exit Sub
        End If
    Next
' RETURNS
    x(1) = x(1) + 1 ' what is this
End Sub


Function compare_all_iterators(ByRef x() As Long, ByRef N() As Long, ByVal NDim As Long) As Boolean
' FUNCTION
    ' Are all x(i) <= N(i)
' ARGUMENTS
    ' Long Array x
    ' Long Array N
    ' Long ndim
' VARIABLE DECLARATION
    Dim i As Long
' ARGUMENT VALIDATION
    compare_all_iterators = False
    If (NDim < 1) Then Exit Function
' VARIABLE INSTANTIATION
    compare_all_iterators = True
' METHODS
    For i = 1 To NDim
        If Not (x(i) <= N(i)) Then
            compare_all_iterators = False
            Exit Function
        End If
    Next
End Function


Function create_N_mat_String(ByVal N As Long, ByRef dims() As Long) As String()
' FUNCTION
    ' Create N-dimensional String-type matrix
' ARGUMENTS
    ' Long N - number of dimensions
    ' Long Array dims - size of each dimension
' ARGUMENT VALIDATION
    If Not N <= SizeArrayDim_Long(dims) Then Exit Function
' VARIABLE INSTANTIATION
    Select Case N
        Case 1
            ReDim create_N_mat_String(max_Long(1, dims(1)))
        Case 2
            ReDim create_N_mat_String(max_Long(1, dims(1)), max_Long(1, dims(2)))
        Case 3
            ReDim create_N_mat_String(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)))
        Case 4
            ReDim create_N_mat_String(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)), max_Long(1, dims(4)))
        Case 5
            ReDim create_N_mat_String(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)), max_Long(1, dims(4)), max_Long(1, dims(5)))
        Case Else:
            Exit Function
        End Select
End Function


Function create_N_mat_Long(ByVal N As Long, ByRef dims() As Long) As Long()
' FUNCTION
    ' Create N-dimensional Long-type matrix
' ARGUMENTS
    ' Long N - number of dimensions
    ' Long Array dims - size of each dimension
' ARGUMENT VALIDATION
    If Not N <= SizeArrayDim_Long(dims) Then Exit Function
' VARIABLE INSTANTIATION
    Select Case N
        Case 1
            ReDim create_N_mat_Long(max_Long(1, dims(1)))
        Case 2
            ReDim create_N_mat_Long(max_Long(1, dims(1)), max_Long(1, dims(2)))
        Case 3
            ReDim create_N_mat_Long(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)))
        Case 4
            ReDim create_N_mat_Long(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)), max_Long(1, dims(4)))
        Case 5
            ReDim create_N_mat_Long(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)), max_Long(1, dims(4)), max_Long(1, dims(5)))
        Case Else:
            Exit Function
        End Select
End Function


Function create_N_mat_Variant(ByVal N As Long, ByRef dims() As Long) As Variant()
' FUNCTION
    ' Create N-dimensional Variant-type matrix
' ARGUMENTS
    ' Long N - number of dimensions
    ' Long Array dims - size of each dimension
' ARGUMENT VALIDATION
    If Not N <= SizeArrayDim_Long(dims) Then Exit Function
' VARIABLE INSTANTIATION
    Select Case N
        Case 1
            ReDim create_N_mat_Variant(max_Long(1, dims(1)))
        Case 2
            ReDim create_N_mat_Variant(max_Long(1, dims(1)), max_Long(1, dims(2)))
        Case 3
            ReDim create_N_mat_Variant(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)))
        Case 4
            ReDim create_N_mat_Variant(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)), max_Long(1, dims(4)))
        Case 5
            ReDim create_N_mat_Variant(max_Long(1, dims(1)), max_Long(1, dims(2)), max_Long(1, dims(3)), max_Long(1, dims(4)), max_Long(1, dims(5)))
        Case Else:
            Exit Function
        End Select
End Function


Function get_index_N_mat_String(ByRef nd_matrix() As String, ByRef position() As Long, ByVal NDim As Long) As String
' FUNCTION
    ' Get value from indexed cell of N-dimensional String-type matrix
' ARGUMENTS
    ' String Matrix nd_matrix
    ' Long Array position
    ' Long ndim - number of dimensions in matrix
' VARIABLE DECLARATION
    Dim x As Long
' ARGUMENT VALIDATION
    get_index_N_mat_String = "Error"
    If NDim < 1 Then Exit Function
    If Not SizeArrayDim_Long(position) = NDim Then Exit Function
    If Not SizeArrayDim_String(nd_matrix, NDim) >= 1 Then Exit Function
    For x = 1 To NDim
        If Not SizeArrayDim_String(nd_matrix, x) >= position(x) Then Exit Function
        If Not position(x) >= 1 Then Exit Function
    Next
' METHODS
    Select Case NDim
        Case 1:
            get_index_N_mat_String = nd_matrix(position(1))
        Case 2:
            get_index_N_mat_String = nd_matrix(position(1), position(2))
        Case 3:
            get_index_N_mat_String = nd_matrix(position(1), position(2), position(3))
        Case 4:
            get_index_N_mat_String = nd_matrix(position(1), position(2), position(3), position(4))
        Case 5:
            get_index_N_mat_String = nd_matrix(position(1), position(2), position(3), position(4), position(5))
        Case Else:
            Exit Function
    End Select
End Function


Function get_index_N_mat_Long(ByRef nd_matrix() As Long, ByRef position() As Long, ByVal NDim As Long) As Long
' FUNCTION
    ' Get value from indexed cell of N-dimensional Long-type matrix
' ARGUMENTS
    ' Long Matrix nd_matrix
    ' Long Array position
    ' Long ndim - number of dimensions in matrix
' VARIABLE DECLARATION
    Dim x As Long
' ARGUMENT VALIDATION
    get_index_N_mat_Long = "Error"
    If NDim < 1 Then Exit Function
    If Not SizeArrayDim_Long(position) = NDim Then Exit Function
    If Not SizeArrayDim_Long(nd_matrix, NDim) >= 1 Then Exit Function
    For x = 1 To NDim
        If Not SizeArrayDim_Long(nd_matrix, x) >= position(x) Then Exit Function
        If Not position(x) >= 1 Then Exit Function
    Next
' METHODS
    Select Case NDim
        Case 1:
            get_index_N_mat_Long = nd_matrix(position(1))
        Case 2:
            get_index_N_mat_Long = nd_matrix(position(1), position(2))
        Case 3:
            get_index_N_mat_Long = nd_matrix(position(1), position(2), position(3))
        Case 4:
            get_index_N_mat_Long = nd_matrix(position(1), position(2), position(3), position(4))
        Case 5:
            get_index_N_mat_Long = nd_matrix(position(1), position(2), position(3), position(4), position(5))
        Case Else:
            Exit Function
    End Select
End Function


Function get_index_N_mat_Variant(ByRef nd_matrix() As Variant, ByRef position() As Long, ByVal NDim As Long) As Variant
' FUNCTION
    ' Get value from indexed cell of N-dimensional Variant-type matrix
' ARGUMENTS
    ' Variant Matrix nd_matrix
    ' Long Array position
    ' Long ndim - number of dimensions in matrix
' VARIABLE DECLARATION
    Dim x As Long
' ARGUMENT VALIDATION
    get_index_N_mat_Variant = "Error"
    If NDim < 1 Then Exit Function
    If Not SizeArrayDim_Long(position) = NDim Then Exit Function
    If Not SizeArrayDim_Variant(nd_matrix, NDim) >= 1 Then Exit Function
    For x = 1 To NDim
        If Not SizeArrayDim_Variant(nd_matrix, x) >= position(x) Then Exit Function
        If Not position(x) >= 1 Then Exit Function
    Next
' METHODS
    Select Case NDim
        Case 1:
            get_index_N_mat_Variant = nd_matrix(position(1))
        Case 2:
            get_index_N_mat_Variant = nd_matrix(position(1), position(2))
        Case 3:
            get_index_N_mat_Variant = nd_matrix(position(1), position(2), position(3))
        Case 4:
            get_index_N_mat_Variant = nd_matrix(position(1), position(2), position(3), position(4))
        Case 5:
            get_index_N_mat_Variant = nd_matrix(position(1), position(2), position(3), position(4), position(5))
        Case Else:
            Exit Function
    End Select
End Function


Sub set_index_N_mat_String(ByRef nd_matrix() As String, ByRef position() As Long, ByVal NDim As Long, ByVal vNew As String)
' FUNCTION
    ' Get value from indexed cell of N-dimensional String-type matrix
' ARGUMENTS
    ' String Matrix nd_matrix
    ' Long Array position
    ' Long ndim - number of dimensions in matrix
' VARIABLE DECLARATION
    Dim x As Long
' ARGUMENT VALIDATION
    ' get_index_N_mat_String = "Error"
    If NDim < 1 Then Exit Sub
    If Not SizeArrayDim_Long(position) = NDim Then GoTo exitsub
    If Not SizeArrayDim_String(nd_matrix, NDim) >= 1 Then GoTo exitsub
    For x = 1 To NDim
        If Not SizeArrayDim_String(nd_matrix, x) >= position(x) Then GoTo exitsub
        If Not position(x) >= 1 Then GoTo exitsub
    Next
' METHODS
    Select Case NDim
        Case 1:
            nd_matrix(position(1)) = vNew
        Case 2:
            nd_matrix(position(1), position(2)) = vNew
        Case 3:
            nd_matrix(position(1), position(2), position(3)) = vNew
        Case 4:
            nd_matrix(position(1), position(2), position(3), position(4)) = vNew
        Case 5:
            nd_matrix(position(1), position(2), position(3), position(4), position(5)) = vNew
        Case Else:
            Exit Sub
    End Select
    Exit Sub
exitsub:
    MsgBox "Error"
End Sub


Sub set_index_N_mat_Long(ByRef nd_matrix() As Long, ByRef position() As Long, ByVal NDim As Long, ByVal vNew As Long)
' FUNCTION
    ' Get value from indexed cell of N-dimensional String-type matrix
' ARGUMENTS
    ' Long Matrix nd_matrix
    ' Long Array position
    ' Long ndim - number of dimensions in matrix
' VARIABLE DECLARATION
    Dim x As Long
' ARGUMENT VALIDATION
    ' get_index_N_mat_String = "Error"
    If NDim < 1 Then Exit Sub
    If Not SizeArrayDim_Long(position) = NDim Then GoTo exitsub
    If Not SizeArrayDim_Long(nd_matrix, NDim) >= 1 Then GoTo exitsub
    For x = 1 To NDim
        If Not SizeArrayDim_Long(nd_matrix, x) >= position(x) Then GoTo exitsub
        If Not position(x) >= 1 Then GoTo exitsub
    Next
' METHODS
    Select Case NDim
        Case 1:
            nd_matrix(position(1)) = vNew
        Case 2:
            nd_matrix(position(1), position(2)) = vNew
        Case 3:
            nd_matrix(position(1), position(2), position(3)) = vNew
        Case 4:
            nd_matrix(position(1), position(2), position(3), position(4)) = vNew
        Case 5:
            nd_matrix(position(1), position(2), position(3), position(4), position(5)) = vNew
        Case Else:
            Exit Sub
    End Select
    Exit Sub
exitsub:
    MsgBox "Error"
End Sub


Sub set_index_N_mat_Variant(ByRef nd_matrix() As Variant, ByRef position() As Long, ByVal NDim As Long, ByVal vNew As Variant)
' FUNCTION
    ' Get value from indexed cell of N-dimensional String-type matrix
' ARGUMENTS
    ' Variant Matrix nd_matrix
    ' Long Array position
    ' Long ndim - number of dimensions in matrix
' VARIABLE DECLARATION
    Dim x As Long
' ARGUMENT VALIDATION
    ' get_index_N_mat_String = "Error"
    If NDim < 1 Then GoTo exitsub
    If Not SizeArrayDim_Long(position) = NDim Then GoTo exitsub
    If Not SizeArrayDim_Variant(nd_matrix, NDim) >= 1 Then GoTo exitsub
    For x = 1 To NDim
        If Not SizeArrayDim_Variant(nd_matrix, x) >= position(x) Then GoTo exitsub
        If Not position(x) >= 1 Then GoTo exitsub
    Next
' METHODS
    Select Case NDim
        Case 1:
            nd_matrix(position(1)) = vNew
        Case 2:
            nd_matrix(position(1), position(2)) = vNew
        Case 3:
            nd_matrix(position(1), position(2), position(3)) = vNew
        Case 4:
            nd_matrix(position(1), position(2), position(3), position(4)) = vNew
        Case 5:
            nd_matrix(position(1), position(2), position(3), position(4), position(5)) = vNew
        Case Else:
            Exit Sub
    End Select
    Exit Sub
exitsub:
    MsgBox "Error"
End Sub


Sub convert_1D_Variant_2_String(ByRef input_array() As Variant, ByRef return_array() As String)
' FUNCTION
    ' Convert 1D Variant array to 1D String array
' ARGUMENTS
    ' Variant Array input_array
    ' String Array return_array
' VARIABLE DECLARATION
    Dim N As Long
    Dim i As Long
' ARGUMENT VALIDATION
    If ((Not input_array) = -1) Then
        ReDim return_array(1)
        return_array(1) = "Error"
        Exit Sub
    End If
' VARIABLE INSTANTIATION
    N = SizeArrayDim_Variant(input_array)
    ReDim return_array(N)
' METHODS
    For i = 1 To N
        return_array(i) = CStr(input_array(i))
    Next
End Sub


Sub convert_1D_Variant_2_Long(ByRef input_array() As Variant, ByRef return_array() As Long)
' FUNCTION
    ' Convert 1D Variant array to 1D Long array
' ARGUMENTS
    ' Variant Array input_array
    ' Long Array return_array
' VARIABLE DECLARATION
    Dim N As Long
    Dim i As Long
' ARGUMENT VALIDATION
    If ((Not input_array) = -1) Then
        ReDim return_array(1)
        return_array(1) = "Error"
        Exit Sub
    End If
' VARIABLE INSTANTIATION
    N = SizeArrayDim_Variant(input_array)
    ReDim return_array(N)
' METHODS
    For i = 1 To N
        return_array(i) = CStr(input_array(i))
    Next
End Sub


Sub convert_2D_Variant_2_String(ByRef input_array() As Variant, ByRef return_array() As String)
' FUNCTION
    ' Convert 2D matrix from Variant to String
' ARGUMENTS
    ' Variant Array input_array
    ' String Array return_array
' VARIABLE DECLARATION
    Dim N(2) As Long
    Dim i As Long
    Dim j As Long
' ARGUMENT VALIDATION
    If ((Not input_array) = -1) Then
        ReDim return_array(1)
        return_array(1) = "Error"
        Exit Sub
    End If
' VARIABLE INSTANTIATION
    ' ReDim N(2)
    N(1) = SizeArrayDim_Variant(input_array, 1)
    N(2) = SizeArrayDim_Variant(input_array, 2)
    ReDim return_array(N(1), N(2))
' METHODS
    For i = 1 To N(1)
        For j = 1 To N(2)
            return_array(i, j) = CStr(input_array(i, j))
        Next
    Next
End Sub


Sub convert_2D_Variant_2_Long(ByRef input_array() As Variant, ByRef return_array() As Long)
' FUNCTION
    ' Convert 2D matrix from Variant to Long
' ARGUMENTS
    ' Variant Array input_array
    ' Long Array return_array
' VARIABLE DECLARATION
    Dim N() As Long
    Dim i As Long
    Dim j As Long
' ARGUMENT VALIDATION
    If ((Not input_array) = -1) Then
        ReDim return_array(1)
        return_array(1) = "Error"
        Exit Sub
    End If
' VARIABLE INSTANTIATION
    ReDim N(2)
    N(1) = SizeArrayDim_Variant(input_array, 1)
    N(2) = SizeArrayDim_Variant(input_array, 2)
    ReDim return_array(N(1), N(2))
' METHODS
    For i = 1 To N(1)
        For j = 1 To N(2)
            return_array(i, j) = CLng(input_array(i, j))
        Next
    Next
End Sub


Function last_filled_cell(ByVal array1D As Variant, Optional start As Long = 1, Optional max_i As Long = -1, Optional gap_max As Long = 1, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Find last populated cell in 1D array from start to max_i with maximum number of consecutive empty cells gap_max
' ARGUMENTS
    ' Variant Array array1D
    ' Long start
    ' Long max_i
    ' Long gap_max
    ' dir_traverse dir_move
' VARIABLE DECLARATION
    Dim found_last As Boolean
    Dim temp As Variant
    Dim gap_count As Long
    Dim i As Long
' ARGUMENT VALIDATION
    If max_i < 1 Then
        max_i = SizeArrayDim_Variant_0(array1D)
    End If
    If (start < 0 Or start > max_i) Then Exit Function
' VARIABLE INSTANTIATION
    If dir_move = FORWARDS Then
        i = start - 1
    Else
        i = max_i
        max_i = start
        start = i
        i = max_i + 1
    End If
' METHODS
    Do While Not found_last
        i = i + 1
        If i = max_i Then
            found_last = True
        ElseIf IsEmpty(array1D(i)) Or array1D(i) = "" Then
            gap_count = gap_count + 1
        Else
            gap_count = 0
        End If
        If gap_count > gap_max Then
            found_last = True
        End If
    Loop
' RETURNS
    last_filled_cell = i - gap_count
End Function


Function max_n_Long(ByRef array1D() As Long, Optional N As Long = 0, Optional start As Long = 1, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Maximum Long value in array1D from start up to start + N - 1
' ARGUMENTS
    ' Long Array array1D
    ' Long N - number of elements to check
    ' Long start - first element to check
' VARIABLE DECLARATION
    Dim i As Long
    Dim max_i As Long
    Dim grand_max As Long
' ARGUMENT VALIDATION
    start = max_Long(1, start)
    max_i = SizeArrayDim_Long(array1D)
    If N > 0 Then
        If start + N - 1 > max_i Then
            MsgBox "Invalid search indices for array of size " & CStr(max_i)
        End If
    End If
' METHODS
    If dir_move = FORWARDS Then
        For i = start To max_i
            grand_max = max_Long(grand_max, array1D(i))
        Next
    Else
        For i = start To 1 Step -1
            grand_max = max_Long(grand_max, array1D(i))
        Next
    End If
' RETURNS
    max_n_Long = grand_max
End Function


Function min_n_Long(ByRef array1D() As Long, Optional N As Long = 0, Optional start As Long = 1, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Minimum Long value in array1D from start up to start + N - 1
' ARGUMENTS
    ' Long Array array1D
    ' Long N - number of elements to check
    ' Long start - first element to check
' VARIABLE DECLARATION
    Dim i As Long
    Dim max_i As Long
    Dim grand_min As Long
' ARGUMENT VALIDATION
    start = max_Long(1, start)
    max_i = SizeArrayDim_Long(array1D)
    If N > 0 Then
        If start + N - 1 > max_i Then
            MsgBox "Invalid search indices for array of size " & CStr(max_i)
        End If
    End If
' METHODS
    If dir_move = FORWARDS Then
        For i = start To max_i
            grand_min = min_Long(grand_min, array1D(i))
        Next
    Else
        For i = start To 1 Step -1
            grand_min = min_Long(grand_min, array1D(i))
        Next
    End If
' RETURNS
    min_n_Long = grand_min
End Function


Function match_Long(ByVal key As Long, ByRef array1D() As Long, Optional start As Long = 1, Optional max_i As Long = 0, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Find first or last instance of key in array1D
' ARGUMENTS
    ' Long key
    ' Long Array array1D
    ' Long start
    ' Long max_i
    ' dir_traverse dir_move
' VARIABLE DECLARATION
    Dim i As Long
' ARGUMENT VALIDATION
    If max_i < 1 Then
        max_i = SizeArrayDim_Long(array1D)
    End If
    If (start < 0 Or start > max_i) Then Exit Function
' METHODS
    If dir_move = FORWARDS Then
        For i = start To max_i
            If array1D(i) = key Then
                match_Long = i
                Exit Function
            End If
        Next
    Else
        For i = start To 1 Step -1
            If array1D(i) = key Then
                match_Long = i
                Exit Function
            End If
        Next
    End If
End Function


Function match_String(ByVal key As String, ByRef array1D() As String, Optional start As Long = 1, Optional max_i As Long = 0, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Find first or last instance of key in array1D
' ARGUMENTS
    ' String key
    ' String Array array1D
    ' Long start
    ' Long max_i
    ' dir_traverse dir_move
' VARIABLE DECLARATION
    Dim i As Long
' ARGUMENT VALIDATION
    If max_i < 1 Then
        max_i = SizeArrayDim_String(array1D)
    End If
    If (start < 0 Or start > max_i) Then Exit Function
' METHODS
    If dir_move = FORWARDS Then
        For i = start To max_i
            If array1D(i) = key Then
                match_String = i
                Exit Function
            End If
        Next
    Else
        For i = start To 1 Step -1
            If array1D(i) = key Then
                match_String = i
                Exit Function
            End If
        Next
    End If
End Function


Function match_Variant(ByVal key As Variant, ByRef array1D() As Variant, Optional start As Long = 1, Optional max_i As Long = 0, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Find first or last instance of key in array1D
' ARGUMENTS
    ' Variant key
    ' Variant Array array1D
    ' Long start
    ' Long max_i
    ' dir_traverse dir_move
' VARIABLE DECLARATION
    Dim i As Long
' ARGUMENT VALIDATION
    If max_i < 1 Then
        max_i = SizeArrayDim_Variant(array1D, 2)
    End If
    If (start < 0 Or start > max_i) Then Exit Function
' METHODS
    If dir_move = FORWARDS Then
        For i = start To max_i
            If array1D(i) = key Then
                match_Variant = i
                Exit Function
            End If
        Next
    Else
        For i = start To 1 Step -1
            If array1D(i) = key Then
                match_Variant = i
                Exit Function
            End If
        Next
    End If
End Function


Function change_array_base_String(ByRef array1D() As String, Optional base_new As Long = 1) As String()
' FUNCTION
    ' Change base of array1D to new_base, if necessary
' ARGUMENTS
    ' String Array array1D
    ' Long base_new
' VARIABLE DECLARATION
    Dim base_old As Long
    Dim i As Long
    Dim N As Long
    Dim array_out() As String
' ARGUMENT VALIDATION
    N = SizeArrayDim_String(array1D)
    If N < 1 Then Exit Function
' VARIABLE INSTANTIATION
    base_old = LBound(array1D)
    ReDim array_out(N)
' METHODS
    If base_old = base_new Then
        change_array_base_String = array1D
        Exit Function
    End If
    For i = 1 To N
        array_out(base_new + i - 1) = array1D(base_old + i - 1)
    Next
' RETURNS
    change_array_base_String = array_out
End Function


Function get_1D_strip_N_mat_String(ByRef N_mat() As String, ByVal dimension As Long, ByRef position() As Long) As String()
' FUNCTION
    ' Get 1D array of String-type N-dimensional matrix with all but one dimension fixed
' ARGUMENTS
    ' String Array N_mat
    ' Long dimension - dimension of N-mat to vary
    ' Long Array position - fixed coordinates
' VARIABLE DECLARATION
    Dim Outs() As String
    Dim x() As Long
    Dim i As Long
    Dim N As Long
    Dim x_max As Long
    Dim p_max As Long
' ARGUMENT VALIDATION
    If dimension < 1 Then
        MsgBox "Error: Dimension out of range."
        Exit Function
    End If
    p_max = SizeArrayDim_Long(position)
    If p_max = 0 Then
        MsgBox "Error: Position is nothing."
        Exit Function
    End If
    If (p_max < dimension - 1) Then
        MsgBox "Error: Insufficient quantity of dimensions in position for dimension " & CStr(dimension) & "."
        Exit Function
    End If
    x_max = SizeArrayDim_String(N_mat, N)
    If x_max = 0 Then
        MsgBox "Error: N-dimensional Matrix of insufficient quantity of dimensions."
        Exit Function
    End If
    x_max = SizeArrayDim_String(N_mat, N + 1)
    If x_max > 0 Then
        MsgBox "Error: N-dimensional Matrix of too many dimensions."
        Exit Function
    End If
' VARIABLE INSTANTIATION
    x_max = SizeArrayDim_String(N_mat, dimension)
    ReDim Outs(x_max)
    ReDim x(N)
    For i = 1 To p_max
        x(i) = position(i)
    Next
' METHODS
    For i = 1 To x_max
        x(dimension) = i
        Outs(i) = get_index_N_mat_String(N_mat, x, N)
    Next
' RETURNS
    get_1D_strip_N_mat_String = Outs
End Function


Function get_1D_strip_N_mat_Long(ByRef N_mat() As Long, ByVal dimension As Long, ByRef position() As Long) As Long()
' FUNCTION
    ' Get 1D array of Long-type N-dimensional matrix with all but one dimension fixed
' ARGUMENTS
    ' String Array N_mat
    ' Long dimension - dimension of N-mat to vary
    ' Long Array position - fixed coordinates
' VARIABLE DECLARATION
    Dim Outs() As Long
    Dim x() As Long
    Dim i As Long
    Dim N As Long
    Dim x_max As Long
    Dim p_max As Long
' ARGUMENT VALIDATION
    If dimension < 1 Then
        MsgBox "Error: Dimension out of range."
        Exit Function
    End If
    p_max = SizeArrayDim_Long(position)
    If p_max = 0 Then
        MsgBox "Error: Position is nothing."
        Exit Function
    End If
    If (p_max < dimension - 1) Then
        MsgBox "Error: Insufficient quantity of dimensions in position for dimension " & CStr(dimension) & "."
        Exit Function
    End If
    N = max_Long(dimension, p_max)
    x_max = SizeArrayDim_Long(N_mat, N)
    If x_max = 0 Then
        MsgBox "Error: N-dimensional Matrix of insufficient quantity of dimensions."
        Exit Function
    End If
    x_max = SizeArrayDim_Long(N_mat, N + 1)
    If x_max > 0 Then
        MsgBox "Error: N-dimensional Matrix of too many dimensions."
        Exit Function
    End If
' VARIABLE INSTANTIATION
    x_max = SizeArrayDim_Long(N_mat, dimension)
    ReDim Outs(x_max)
    ReDim x(N)
    For i = 1 To p_max
        x(i) = position(i)
    Next
' METHODS
    For i = 1 To x_max
        x(dimension) = i
        Outs(i) = get_index_N_mat_Long(N_mat, x, N)
    Next
' RETURNS
    get_1D_strip_N_mat_Long = Outs
End Function


Function get_1D_strip_N_mat_Variant(ByRef N_mat() As Variant, ByVal dimension As Long, ByRef position() As Long) As Variant()
' FUNCTION
    ' Get 1D array of Variant-type N-dimensional matrix with all but one dimension fixed
' ARGUMENTS
    ' String Array N_mat
    ' Long dimension - dimension of N-mat to vary
    ' Long Array position - fixed coordinates
' VARIABLE DECLARATION
    Dim Outs() As Variant
    Dim x() As Long
    Dim i As Long
    Dim N As Long
    Dim x_max As Long
    Dim p_max As Long
' ARGUMENT VALIDATION
    If dimension < 1 Then
        MsgBox "Error: Dimension out of range."
        Exit Function
    End If
    p_max = SizeArrayDim_Long(position)
    If p_max = 0 Then
        MsgBox "Error: Position is nothing."
        Exit Function
    End If
    If (p_max < dimension - 1) Then
        MsgBox "Error: Insufficient quantity of dimensions in position for dimension " & CStr(dimension) & "."
        Exit Function
    End If
    x_max = SizeArrayDim_Variant(N_mat, N)
    If x_max = 0 Then
        MsgBox "Error: N-dimensional Matrix of insufficient quantity of dimensions."
        Exit Function
    End If
    x_max = SizeArrayDim_Variant(N_mat, N + 1)
    If x_max > 0 Then
        MsgBox "Error: N-dimensional Matrix of too many dimensions."
        Exit Function
    End If
' VARIABLE INSTANTIATION
    x_max = SizeArrayDim_Variant(N_mat, dimension)
    ReDim Outs(x_max)
    ReDim x(N)
    For i = 1 To p_max
        x(i) = position(i)
    Next
' METHODS
    For i = 1 To x_max
        x(dimension) = i
        Outs(i) = get_index_N_mat_Variant(N_mat, x, N)
    Next
' RETURNS
    get_1D_strip_N_mat_Variant = Outs
End Function


Function Mod_Double(ByVal numerator As Double, ByVal denominator As Double) As Double
' FUNCTION
    ' Modulo method for doubles
' ARGUMENTS
    ' Double numerator
    ' Double denominator
' METHODS
    Do While numerator > denominator
        numerator = numerator - denominator
    Loop
    Do While numerator < 0
        numerator = numerator + denominator
    Loop
' RETURNS
    Mod_Double = numerator
End Function


Function Div_Double(ByVal numerator As Double, ByVal denominator As Double) As Long
' FUNCTION
    ' Modular division method for doubles
' ARGUMENTS
    ' Double numerator
    ' Double denominator
' VARIABLE DECLARATION
    Dim count As Long
' VARIABLE DECLARATION
    count = 0
' METHODS
    Do While numerator > denominator
        count = count + 1
        numerator = numerator - denominator
    Loop
    Do While numerator < 0
        count = count - 1
        numerator = numerator + denominator
    Loop
' RETURNS
    Div_Double = count
End Function

