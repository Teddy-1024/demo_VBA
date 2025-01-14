' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit


' PROPERTY DECLARATION
Private v_ws As Worksheet
Private v_name As String
Private v_dir As orientation
Private v_col_min As Long
Private v_col_max As Long
Private v_col_search As Long
Private v_row_min As Long
Private v_row_max As Long
Private v_row_search As Long
Private v_data() As String
Private v_headings() As String
Private v_col_IDs() As Long
Private v_mutable_headings As Boolean
Private v_gap_max As Long


' PROPERTY METHODS
Public Property Get Worksheet() As Worksheet
    Set Worksheet = v_ws
End Property

Public Property Get name() As String
    name = v_name
End Property
Public Property Let name(ByVal new_name As String)
    v_name = new_name
End Property

Public Property Get Cell(ByVal row As Long, ByVal col As Long) As Range
    Set Cell = v_ws.Cells(row, col)
End Property

Public Property Get Range(ByVal range_str As String) As Range
    Set Range = v_ws.Range(range_str)
End Property
Public Sub SetRange_Variant(ByVal range_str As String, ByRef new_data() As Variant)
' FUNCTION
    ' Set Range of v_ws
' ARGUMENTS
    ' String range_str
    ' Variant Matrix new_data
' VARIABLE DECLARATION
    Dim sz_x As Long
    Dim sz_y As Long
    Dim sz_z As Long
    Dim out_data() As String
    Dim range_coords() As Long
' VARIABLE INSTANTIATION
    sz_x = SizeArrayDim_Variant(new_data, 2)
    sz_y = SizeArrayDim_Variant(new_data, 1)
    sz_z = SizeArrayDim_Variant(new_data, 3)
' ARGUMENT VALIDATION
    ' string-literal decomposition of range string
    If Not valid_range_String(range_str) Then
        MsgBox "Invalid range string"
        Exit Sub
    End If
    range_coords = Range_String_Coords(range_str)
    ' Validate new coordinates
    If Not (valid_coordinate(range_coords(1), range_coords(2)) And valid_coordinate(range_coords(3), range_coords(4))) Then
        MsgBox "Invalid range string coordinates"
        Exit Sub
    End If
    ' Validate new_data dimensions against range_str dimensions
    If Not ((sz_x = range_coords(4) - range_coords(2) + 1) And (sz_x = range_coords(4) - range_coords(2) + 1) And (sz_z = 0)) Then
        MsgBox "Inconsistent range and matrix dimensions"
        Exit Sub
    End If
' METHODS
    ' Compare new dimensions against current dimensions
    If Not (range_coords(1) = v_row_min And range_coords(2) = v_col_min And range_coords(3) = v_row_max And range_coords(4) = v_col_max) Then
        ResizeLocal range_coords(2), range_coords(4), range_coords(1), range_coords(3), True
    End If
' RETURNS
    v_ws.Range(range_str).value = new_data
    convert_2D_Variant_2_String new_data, out_data
    Me.SetData out_data
End Sub
Public Sub SetRange_String(ByVal range_str As String, ByRef new_data() As String)
' FUNCTION
    ' Set Range of v_ws
' ARGUMENTS
    ' String range_str
    ' String Matrix new_data
' VARIABLE DECLARATION
    Dim sz_x As Long
    Dim sz_y As Long
    Dim sz_z As Long
    Dim range_coords() As Long
    Dim i As Long
    Dim j As Long
' VARIABLE INSTANTIATION
    sz_x = SizeArrayDim_String(new_data, 2)
    sz_y = SizeArrayDim_String(new_data, 1)
    sz_z = SizeArrayDim_String(new_data, 3)
' ARGUMENT VALIDATION
    ' string-literal decomposition of range string
    If Not valid_range_String(range_str) Then
        MsgBox "Invalid range string"
        Exit Sub
    End If
    range_coords = Range_String_Coords(range_str)
    ' Validate new coordinates
    If Not (valid_coordinate(range_coords(1), range_coords(2)) And valid_coordinate(range_coords(3), range_coords(4))) Then
        MsgBox "Invalid range string coordinates"
        Exit Sub
    End If
    ' Validate new_data dimensions against range_str dimensions
    If Not ((sz_x = range_coords(4) - range_coords(2) + 1) And (sz_x = range_coords(4) - range_coords(2) + 1) And (sz_z = 0)) Then
        MsgBox "Inconsistent range and matrix dimensions"
        Exit Sub
    End If
' METHODS
    ' Compare new dimensions against current dimensions
    If Not (range_coords(1) = v_row_min And range_coords(2) = v_col_min And range_coords(3) = v_row_max And range_coords(4) = v_col_max) Then
        ResizeLocal range_coords(2), range_coords(4), range_coords(1), range_coords(3), True
    End If
' RETURNS
    For i = v_col_min To v_col_max
        For j = v_row_min To v_row_max
            v_ws.Cells(j, i).value = new_data(j, i)
        Next
    Next
    Me.SetData new_data
End Sub

Public Sub GetData(ByRef data_out() As String)
    data_out = v_data
End Sub
Public Sub SetData(ByRef new_data() As String, Optional update_ws As Boolean = False)
' FUNCTION
    ' Set local data to new matrix, and optionally, update Worksheet
' ARGUMENTS
    ' String Matrix new_data
    ' Boolean update_ws
' VARIABLE DECLARATION
    Dim sz_x As Long
    Dim sz_y As Long
    Dim sz_z As Long
' VARIABLE INSTANTIATION
    sz_x = SizeArrayDim_String(new_data, 2)
    sz_y = SizeArrayDim_String(new_data, 1)
    sz_z = SizeArrayDim_String(new_data, 3)
' ARGUMENT VALIDATION
    If Not ((sz_x = v_col_max - v_col_min + 1) And (sz_y = v_row_max - v_row_min + 1) And (sz_z = 0)) Then
        MsgBox "Invalid dimensions to replace worksheet data"
        Exit Sub
    End If
' METHODS
    Erase v_data
' RETURNS
    v_data = new_data
    If update_ws Then ExportLocalData2WS
End Sub
Public Sub SetEmptyData(ByRef new_data() As String, Optional update_ws As Boolean = False)
' FUNCTION
    ' Set local data to new matrix, and optionally, update Worksheet. Only populate empty cells
' ARGUMENTS
    ' String Matrix new_data
    ' Boolean update_ws
' VARIABLE DECLARATION
    Dim sz_x As Long
    Dim sz_y As Long
    Dim sz_z As Long
    Dim i As Long
    Dim j As Long
' VARIABLE INSTANTIATION
    sz_x = SizeArrayDim_String(new_data, 2)
    sz_y = SizeArrayDim_String(new_data, 1)
    sz_z = SizeArrayDim_String(new_data, 3)
' ARGUMENT VALIDATION
    If Not ((sz_x = v_col_max - v_col_min + 1) And (sz_y = v_row_max - v_row_min + 1) And (sz_z = 0)) Then
        MsgBox "Invalid dimensions to replace worksheet data"
        Exit Sub
    End If
' METHODS
    For j = 1 To sz_y
        For i = 1 To sz_x
            If v_data(j, i) = "" Then v_data(j, i) = new_data(j, i)
        Next
    Next
' RETURNS
    If update_ws Then ExportLocalData2WS
End Sub

Public Function GetCellDataLocal(ByVal row As Long, ByVal col As Long) As String
' FUNCTION
    ' Get value from cell of local copy of Worksheet data
' ARGUMENTS
    ' Long row - relative row index in data table
    ' Long col - relative column index in data table
' VARIABLE INSTANTIATION
    If ((row > v_row_max - v_row_min + 1) Or (row < 1) Or (col > v_col_max - v_col_min + 1) Or (col < 1)) Then
        Debug.Print "Invalid cell index. Row = " & CStr(row) & ", column = " & CStr(col)
        Exit Function
    End If
' RETURNS
    GetCellDataLocal = v_data(row, col)
End Function
Public Sub SetCellDataLocal(ByVal row As Long, ByVal col As Long, ByVal new_value As String)
' FUNCTION
    ' Set value of cell of local copy of Worksheet data
' ARGUMENTS
    ' Long row - relative row index in data table
    ' Long col - relative column index in data table
' VARIABLE INSTANTIATION
    If ((row > v_row_max - v_row_min + 1) Or (row < 1) Or (col > v_col_max - v_col_min + 1) Or (col < 1)) Then
        Debug.Print "Invalid cell index. Row = " & CStr(row) & ", column = " & CStr(col)
        Exit Sub
    End If
' METHODS
    v_data(row, col) = new_value
End Sub

Public Property Get nHeadings() As Long
    nHeadings = SizeArrayDim_String(v_headings)
End Property

Public Property Get HeadingName(ByVal h_index As Long) As String
' FUNCTION
    ' Get heading text from index in v_headings
' ARGUMENTS
    ' Long h_index - index of desired heading in v_headings
' VARIABLE DECLARATION
    Dim n_h As Long
' VARIABLE INSTANTIATION
    n_h = nHeadings
    HeadingName = "Error: Invalid index"
' ARGUMENT VALIDATION
    If (h_index < 1 Or h_index > n_h) Then Exit Property
' RETURNS
    HeadingName = v_headings(h_index)
End Property

Public Property Get HeadingIndex(ByVal h_search As String, Optional ID_in_searchstr As Boolean = False) As Long
' FUNCTION
    ' Get index of h_search within v_headings, if it exists
' ARGUMENTS
    ' String h_search - heading to search for
    ' Boolean ID_in_searchstr - is index in h_search instead of heading name?
' VARIABLE DECLARATION
    Dim n_h As Long
    Dim h_i As Long
' ARGUMENT VALIDATION
    If h_search = "" Then Exit Property
    If Not ExistsHeading(h_search, ID_in_searchstr) Then Exit Property
' VARIABLE INSTANTIATION
    n_h = nHeadings
' METHODS
    If ID_in_searchstr Then
        h_i = CLng(h_search)
        If (h_i < 1 Or h_i > n_h) Then Exit Property
        HeadingIndex = h_i
    Else
        For h_i = 1 To n_h
            If (v_headings(h_i) = h_search) Then
                HeadingIndex = h_i
                Exit Property
            End If
        Next
    End If
End Property

Public Property Get ColumnID(ByVal h_search As String, Optional ID_in_searchstr As Boolean = False) As Long
' FUNCTION
    ' Get index of column containing h_search, if it exists within v_headings
' ARGUMENTS
    ' String h_search - heading to search for
    ' Boolean ID_in_searchstr - is index in h_search instead of heading name?
' VARIABLE DECLARATION
    Dim n_h As Long
    Dim h_i As Long
' ARGUMENT VALIDATION
    If h_search = "" Then Exit Property
    If Not ExistsHeading(h_search, ID_in_searchstr) Then Exit Property
' VARIABLE INSTANTIATION
    n_h = nHeadings
' METHODS
    If ID_in_searchstr Then
        h_i = CLng(h_search)
        If (h_i < 1 Or h_i > n_h) Then Exit Property
        ColumnID = v_col_IDs(CLng(h_search))
    Else
        For h_i = 1 To n_h
            If (v_headings(h_i) = h_search) Then
                ColumnID = v_col_IDs(h_i)
                Exit Property
            End If
        Next
    End If
End Property

Public Property Get SearchID() As Long
' FUNCTION
    ' Get index of search row / column
' METHODS
    If v_dir = ColumnHeaders Then
        SearchID = v_col_search
    Else
        SearchID = v_row_search
    End If
End Property

Public Property Get ColumnMin() As Long
    ColumnMin = v_col_min
End Property
Public Property Get ColumnMax() As Long
    ColumnMax = v_col_max
End Property
Public Property Get ColumnSearch() As Long
    ColumnSearch = v_col_search
End Property
Public Property Let ColumnSearch(ByVal vNew As Long)
    If Not valid_coordinate(1, vNew) Then
        MsgBox "Error: Invalid new search column index"
        Exit Property
    End If
    If vNew < v_col_min And v_dir = ColumnHeaders Then v_col_min = vNew
    If vNew > v_col_max And v_dir = ColumnHeaders Then v_col_max = vNew
    If vNew > v_col_min And v_dir = RowHeaders Then v_col_min = vNew + 1
    v_col_search = vNew
End Property

Public Property Get RowMin() As Long
    RowMin = v_row_min
End Property
Public Property Get RowMax() As Long
    RowMax = v_row_max
End Property
Public Property Get RowSearch() As Long
    RowSearch = v_row_search
End Property
Public Property Let RowSearch(ByVal vNew As Long)
    If Not valid_coordinate(vNew, 1) Then
        MsgBox "Error: Invalid new search column index"
        Exit Property
    End If
    If vNew < v_row_min And v_dir = RowHeaders Then v_row_min = vNew
    If vNew > v_row_max And v_dir = RowHeaders Then v_row_max = vNew
    If vNew > v_row_min And v_dir = ColumnHeaders Then v_row_min = vNew + 1
    v_row_search = vNew
End Property

Public Property Get Orient() As orientation
    Orient = v_dir
End Property

Public Property Get GapMax() As Long
    GapMax = v_gap_max
End Property


' METHODS
Public Sub Init(ByRef ws As Worksheet, ByVal name As String, ByRef headings() As String, Optional Orient As orientation = orientation.ColumnHeaders, Optional search_col As Long = 1, Optional search_col_is_heading_index As Boolean = False, Optional search_row As Long = 1, Optional search_row_is_heading_index As Boolean = False, Optional col_min As Long = 1, Optional col_max As Long = 0, Optional row_min As Long = 2, Optional row_max As Long = 0, Optional gap_max As Long = 1, Optional mutable_headings As Boolean = False)
' FUNCTION
    ' Instantiate worksheet container
' ARGUMENTS
    ' Worksheet ws
    ' String name
    ' String Array headings
    ' orientation Orient
    ' Long search_col
    ' Boolean search_col_is_heading_index
    ' Long search_row
    ' Boolean search_row_is_heading_index
    ' Long col_min
    ' Long col_max
    ' Long row_min
    ' Long row_max
    ' Long gap_max
    ' Boolean mutable_headings - can headings be over-written?
' PROCESSING ACCELERATION
' CONSTANTS
    Const uni_col_max As Long = 16384
    Const uni_row_max As Long = 1048576
' VARIABLE DECLARATION
    If ws Is Nothing Then Exit Sub
    Erase v_data
    Dim my_data() As Variant
    Dim col_min_str As String
    Dim col_max_str As String
    Dim col_search_str As String
    Dim n_h As Long
    Dim ws_data_in() As Variant
    Dim temp_col_min As String
    Dim temp_col_max As String
' ARGUMENT VALIDATION
' VARIABLE INSTANTIATION
    If (Orient = ColumnHeaders) Then
        v_col_search = min_Long(max_Long(1, search_col), uni_col_max)
        v_row_search = min_Long(max_Long(1, search_row), uni_row_max - 1)
        v_col_min = min_Long(max_Long(1, col_min), uni_col_max - 1)
        v_row_min = min_Long(max_Long(1 + v_row_search, row_min), uni_row_max)
    Else
        v_col_search = min_Long(max_Long(1, search_col), uni_col_max - 1)
        v_row_search = min_Long(max_Long(1, search_row), uni_row_max)
        v_col_min = min_Long(max_Long(1 + v_col_search, col_min), uni_col_max)
        v_row_min = min_Long(max_Long(1, row_min), uni_row_max - 1)
    End If
    col_min_str = get_col_str(v_col_min)
    n_h = SizeArrayDim_String(headings)
' METHODS
    Set v_ws = ws
    v_name = name
    ' Dimensions
    If (Orient = ColumnHeaders) Then
        If (col_max < v_col_min) Then
            v_col_max = last_filled_cell(Range_1D_String(ws, Range_String(1, 0, v_row_search, v_row_search)), v_col_min, -1, gap_max) ' ws.Range("A" & CStr(v_row_search) & ":XFD" & CStr(v_row_search)).Value,
            If v_col_max = v_col_min And IsEmpty(ws.Cells(v_row_search, v_col_max)) Then v_col_max = v_col_search
        Else
            v_col_max = min_Long(col_max, uni_col_max)
        End If
    ' Headings + Column IDs
        If n_h > 0 Then
            v_headings = headings
            GetColIDs
            If (search_col_is_heading_index And v_col_search <= n_h) Then
                If (v_col_IDs(v_col_search) > 0) Then v_col_search = v_col_IDs(v_col_search)
            End If
        End If
        If (row_max < v_row_min) Then
            v_row_max = last_filled_cell(Range_1D_String(ws, Range_String(v_col_search + v_col_min - 1, v_col_search + v_col_min - 1, 1, 0)), v_row_min, -1, gap_max)
            If v_row_max = v_row_min And IsEmpty(ws.Cells(v_row_max, v_col_search + v_col_min - 1)) Then v_row_max = v_row_search
        Else
            v_row_max = min_Long(v_row_max, uni_row_max)
        End If
    ' Get data
        my_data = ws.Range(Range_String(v_col_min, v_col_max, v_row_min, max_Long(v_row_max, v_row_min))).value
    Else
        If (row_max < v_row_min) Then
            v_row_max = last_filled_cell(Range_1D_String(ws, Range_String(1, 0, v_row_search + v_row_min - 1, v_row_search + v_row_min - 1)), v_row_min, -1, gap_max)
            If v_row_max = v_row_min And IsEmpty(ws.Cells(v_row_max, v_col_search)) Then v_row_max = v_row_search
        Else
            v_row_max = min_Long(v_row_max, uni_row_max)
        End If
    ' Headings + Column IDs
        If n_h > 0 Then
            v_headings = headings
            GetColIDs
            If (search_row_is_heading_index And v_row_search <= n_h) Then
                If (v_col_IDs(v_row_search) > 0) Then v_row_search = v_col_IDs(v_row_search)
            End If
        End If
        If (col_max < v_col_min) Then
            v_col_max = last_filled_cell(Range_1D_String(ws, Range_String(1, 0, v_row_search, v_row_search)), v_col_min, -1, gap_max) ' ws.Range("A" & CStr(v_row_search) & ":XFD" & CStr(v_row_search)).Value,
            If v_col_max = v_col_min And IsEmpty(ws.Cells(v_row_search, v_col_max)) Then v_col_max = v_col_search
        Else
            v_col_max = min_Long(col_max, uni_col_max)
        End If
    ' Get data
        my_data = ws.Range(Range_String(v_col_min, max_Long(v_col_max, v_col_min), v_row_min, v_row_max)).value
    End If
' RETURNS
    convert_2D_Variant_2_String my_data, v_data
    v_mutable_headings = mutable_headings
    v_gap_max = gap_max
End Sub

Private Sub GetColIDs()
' FUNCTION
    ' Find first instance of each of colnames within row RowSearch of ws
' VARIABLE DECLARATION
    Dim NHeading As Long
    Dim Search_S As String
    Dim iCol As Long
    Dim ColName As String
    Dim Outs() As Long
    Dim Temps() As String
    Dim iFail As Long
    Dim fail_S As String
' VARIABLE INSTANTIATION
    NHeading = nHeadings
    ReDim v_col_IDs(NHeading)
    If (v_dir = ColumnHeaders) Then
        ' Temps = v_ws.Range(Range_String(v_col_min, v_col_max, v_row_search, v_row_search)).value
        Temps = Range_1D_String(v_ws, Range_String(v_col_min, v_col_max, v_row_search, v_row_search))
    Else
        ' Temps = v_ws.Range(Range_String(v_col_search, v_col_search, v_row_min, v_row_max)).value
        Temps = Range_1D_String(v_ws, Range_String(v_col_min, v_col_max, v_row_search, v_row_search))
    End If
    iFail = 0
' METHODS
    For iCol = 1 To NHeading
        ColName = v_headings(iCol)
        If Not ColName = "" Then
            v_col_IDs(iCol) = match_String(ColName, Temps)
            If v_col_IDs(iCol) < 1 Then
                iFail = iFail + 1
                If iFail > 1 Then
                    fail_S = fail_S & ", " & ColName
                Else
                    fail_S = ColName
                End If
            End If
        Else
            v_col_IDs(iCol) = 0
        End If
    Next
' RETURNS
    If iFail > 0 Then
        MsgBox "Error: The below headings were not found on Worksheet " & v_name & vbCrLf & fail_S
    End If
End Sub

Public Sub ExportLocalData2WS()
' FUNCTION
    ' Export v_data to v_ws
' METHODS
    If v_dir = ColumnHeaders Then
        v_ws.Range(Range_String(v_col_min, v_col_max, v_row_min, max_Long(v_row_max, v_row_min))).value = v_data
    Else
        v_ws.Range(Range_String(v_col_min, max_Long(v_col_max, v_col_min), v_row_min, v_row_max)).value = v_data
    End If
End Sub

Public Sub AddCellComment(ByVal row As Long, ByVal col As Long, ByVal mycomment As String, Optional row_abs_not_rel As Boolean = False, Optional col_abs_not_rel As Boolean = False)
' FUNCTION
    ' Add comment to cell
' ARGUMENTS
    ' Long row
    ' Long col - cell index
    ' String mycomment
    ' Optional Boolean row_abs_not_rel
    ' Optional Boolean col_abs_not_rel - positioning on worksheet relative to v_data table
' METHODS
    v_ws.Range(Range_String(col, col, row, row)).AddComment mycomment
End Sub

Public Sub ResizeLocalAuto(Optional replace_local_data As Boolean = True)
' FUNCTION
    ' Find table size from Worksheet
' PROCESSING ACCELERATION
    ' Disable automatic spreadsheet calculation - prevents refreshing of whole ws on each cell entry
    Application.Calculation = xlCalculationManual
    ' Disable screen updating
    Application.ScreenUpdating = False
' CONSTANTS
    Const uni_col_max As Long = 16384
    Const uni_row_max As Long = 1048576
' VARIABLE DECLARATION
    Dim my_data() As Variant
' METHODS
    If (Orient = ColumnHeaders) Then
        v_row_max = last_filled_cell(Range_1D_String(v_ws, Range_String(v_col_search, v_col_search, 1, 0)), v_row_min, -1, v_gap_max)
        If v_row_max = v_row_min And IsEmpty(v_ws.Cells(v_row_max, v_col_search)) Then v_row_max = v_row_search
    ' Get data
        my_data = v_ws.Range(Range_String(v_col_min, v_col_max, v_row_min, max_Long(v_row_max, v_row_min))).value
    Else
        v_col_max = last_filled_cell(Range_1D_String(v_ws, Range_String(1, 0, v_row_search, v_row_search)), v_col_min, -1, v_gap_max) ' ws.Range("A" & CStr(v_row_search) & ":XFD" & CStr(v_row_search)).Value,
        If v_col_max = v_col_min And IsEmpty(v_ws.Cells(v_row_search, v_col_max)) Then v_col_max = v_col_search
    ' Get data
        my_data = v_ws.Range(Range_String(v_col_min, max_Long(v_col_max, v_col_min), v_row_min, v_row_max)).value
    End If
' RETURNS
    If replace_local_data Then
        Erase v_data
        convert_2D_Variant_2_String my_data, v_data
    End If
' PROCESSING DECELARATION
    ' Enable automatic spreadsheet calculation - prevents refreshing of whole ws on each cell entry
    Application.Calculation = xlCalculationAutomatic
    ' Enable screen updating
    Application.ScreenUpdating = True
End Sub


Public Sub ResizeLocal(ByVal col_min As Long, ByVal col_max As Long, ByVal row_min As Long, ByVal row_max As Long, Optional clear_data As Boolean = False)
' FUNCTION
    ' Resize local copy of ws table data
' ARGUMENTS
    ' Long col_min
    ' Long col_max
    ' Long row_min
    ' Long row_max
    ' Optional Boolean clear_data
' ARGUMENT VALIDATION
    If col_max < col_min Or row_max < row_min Or (Not valid_coordinate(row_min, col_min)) Or (Not valid_coordinate(row_max, col_max)) Then
        MsgBox "Error: Invalid new local data dimensions." & vbCrLf & error_msg(col_min, "col_min", "Long", "?") & vbCrLf & error_msg(col_max, "col_max", "Long", "?") & vbCrLf & error_msg(row_min, "row_min", "Long", "?") & vbCrLf & error_msg(row_max, "row_max", "Long", "?")
        Exit Sub
    End If
    If col_min = v_col_min And col_max = v_col_max And row_min = v_row_min And row_max = v_row_max Then
        Debug.Print "No change required to local data dimensions."
        Exit Sub
    End If
' METHODS
    If clear_data Then
        v_col_min = col_min
        v_col_max = col_max
        v_row_min = row_min
        v_row_max = row_max
        ClearData
    Else
        v_col_min = col_min
        v_col_max = col_max
        ReDimPreserve_String v_data, 2, v_col_max - v_col_min + 1, 2
        v_row_min = row_min
        v_row_max = row_max
        ReDimPreserve_String v_data, 1, v_row_max - v_row_min + 1, 2
    End If
End Sub


Public Function Match(ByVal v_search As String, Optional array_ID As Long = -1, Optional minimum As Long = 1, Optional maximum As Long = -1, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Find match of v_search in local data and return index
' ARGUMENTS
    ' String v_search
    ' Long array_ID
    ' Long minimum
    ' Long maximum
    ' dir_traverse dir_move
' CONSTANTS
    ' Const RowMaxUniversal As Long = 1048576
    ' Const ColMaxUniversal As Long = 16384
' VARIABLE DECLARATION
    Dim i As Long
    Dim di As Long
    Dim RowMaxUniversal As Long
    Dim ColMaxUniversal As Long
' ARGUMENT VALIDATION
    Match = -1
' VARIABLE INSTANTIATION
    RowMaxUniversal = SizeArrayDim_String(v_data, 1) + v_row_min - 1
    ColMaxUniversal = SizeArrayDim_String(v_data, 2) + v_col_min - 1
    If v_dir = ColumnHeaders Then
        maximum = min_Long(RowMaxUniversal - v_row_min + 1, maximum)
        If maximum < 1 Then maximum = RowMaxUniversal - v_row_min + 1
        minimum = max_Long(1, min_Long(minimum, maximum))
        If Not valid_coordinate(minimum, 1) Then minimum = 1
        If Not valid_coordinate(maximum, 1) Then maximum = v_row_max - v_row_min + 1
        If Not valid_coordinate(1, array_ID) Then array_ID = v_col_search
    Else
        maximum = min_Long(ColMaxUniversal - v_col_min + 1, maximum)
        If maximum < 1 Then maximum = ColMaxUniversal - v_col_min + 1
        minimum = max_Long(1, min_Long(minimum, maximum))
        If Not valid_coordinate(1, minimum) Then minimum = 1
        If Not valid_coordinate(1, maximum) Then maximum = v_col_max - v_col_min + 1
        If Not valid_coordinate(array_ID, 1) Then array_ID = v_row_search
    End If
    If maximum < minimum Then ' not possible
        MsgBox "Error: Invalid minimum (" & CStr(minimum) & ") and maximum (" & CStr(maximum) & ") positions."
        Exit Function
    End If
    If dir_move = BACKWARDS Then
        i = minimum
        minimum = maximum
        maximum = i
        di = -1
    Else
        di = 1
    End If
' METHODS
    For i = minimum To maximum Step di
        If v_dir = ColumnHeaders Then
            If v_data(i, array_ID) = v_search Then
                Match = i
                Exit For
            End If
        Else
            If v_data(array_ID, i) = v_search Then
                Match = i
                Exit For
            End If
        End If
    Next
End Function


Public Function Match_F(ByVal v_search As String, ByVal function_name As String, Optional array_ID As Long = -1, Optional minimum As Long = 1, Optional maximum As Long = -1, Optional dir_move As dir_traverse = dir_traverse.FORWARDS) As Long
' FUNCTION
    ' Find match of v_search in local data and return index
' ARGUMENTS
    ' String v_search
    ' String function_name - name of function which takes single argument of cell data
    ' Long array_ID
    ' Long minimum
    ' Long maximum
    ' dir_traverse dir_move
' VARIABLE DECLARATION
    Dim i As Long
    Dim di As Long
' ARGUMENT VALIDATION
    Match_F = -1
' VARIABLE INSTANTIATION
    i = max_Long(minimum, maximum)
    minimum = max_Long(1, min_Long(minimum, maximum))
    If v_dir = ColumnHeaders Then
        maximum = min_Long(1048576, i)
        If Not valid_coordinate(1, minimum) Then minimum = v_row_min
        If Not valid_coordinate(1, maximum) Then maximum = v_row_max
        If Not valid_coordinate(1, array_ID) Then array_ID = v_row_search
    Else
        maximum = min_Long(16384, i)
        If Not valid_coordinate(minimum, 1) Then minimum = v_col_min
        If Not valid_coordinate(maximum, 1) Then maximum = v_col_max
        If Not valid_coordinate(array_ID, 1) Then array_ID = v_col_search
    End If
    If maximum < minimum Then
        MsgBox "Error: Invalid minimum (" & CStr(minimum) & ") and maximum (" & CStr(maximum) & ") positions."
        Exit Function
    End If
    If dir_move = BACKWARDS Then
        i = minimum
        minimum = maximum
        maximum = i
        di = -1
    Else
        di = 1
    End If
' METHODS
    For i = minimum To maximum Step di
        If v_dir = ColumnHeaders Then
            If CStr(Application.Run(function_name, v_data(i, array_ID))) = v_search Then
                Match_F = i
                Exit For
            End If
        Else
            If CStr(Application.Run(function_name, v_data(array_ID, i))) = v_search Then
                Match_F = i
                Exit For
            End If
        End If
    Next
End Function


Public Function ExistsHeading(ByVal search_heading As String, Optional ID_in_str As Boolean = False) As Boolean
' FUNCTION
    ' Does search_heading exist in v_headings?
' ARGUMENTS
    ' String search_heading
    ' Boolean ID_in_str - is search_heading a heading index instead of name?
' VARIABLE DECLARATION
    Dim h_n As Long
    Dim h_i As Long
    Dim ID As String
' ARGUMENT VALIDATION
    ExistsHeading = False
    If search_heading = "" Then Exit Function
' VARIABLE INSTANTIATION
    h_n = nHeadings
' METHODS
    If ID_in_str Then
        If Not is_numeric(search_heading) Then Exit Function
        ID = CLng(search_heading)
        If ID < 1 Or ID > h_n Then Exit Function
        ExistsHeading = True
        Exit Function
    End If
    For h_i = 1 To h_n
        If v_headings(h_i) = search_heading Then
            ExistsHeading = True
            Exit Function
        End If
    Next
End Function


Public Sub ClearData()
' FUNCTION
    ' Clear contents of local copy of ws data but keep dimensions
' METHODS
    ReDim v_data(v_row_max - v_row_min + 1, v_col_max - v_col_min + 1)
End Sub


Public Function LastFilledCell(Optional search_ID As Long = -1, Optional start As Long = -1, Optional gap_max As Long = -1, Optional Orient As orientation, Optional i_max As Long = -1, Optional dir_move As dir_traverse) As Long
' FUNCTION
    ' Identify index of last filled cell in column or row from start in v_ws
' ARGUMENTS
    ' Optional Long search_ID
    ' Optional Long start
    ' Optional Long gap_max
    ' Optional orientation orient
    ' Optional Long i_max
    ' Optional dir_traverse dir_move
' CONSTANTS
    Const uni_row_max As Long = 1048576
    Const uni_col_max As Long = 16384
' VARIABLE DECLARATION
    Dim found_last As Boolean
    Dim temp As Variant
    Dim gap_n As Long
    Dim i As Long
' ARGUMENT VALIDATION
    If Orient = ColumnHeaders Then
        If i_max < 1 Then i_max = uni_row_max
        i_max = min_Long(max_Long(1, i_max), uni_row_max)
        If search_ID < 1 Then
            search_ID = v_col_search
        Else
            search_ID = min_Long(max_Long(1, search_ID), uni_col_max)
        End If
        gap_max = min_Long(max_Long(1, gap_max), i_max - 1)
        If start = -1 Then
            start = v_row_min
        Else
            start = min_Long(max_Long(1, v_row_min), max_Long(i_max - gap_max, 1))
        End If
    Else
        If i_max < 1 Then i_max = uni_col_max
        i_max = min_Long(max_Long(1, i_max), uni_col_max)
        If search_ID < 1 Then
            search_ID = v_row_search
        Else
            search_ID = min_Long(max_Long(1, search_ID), uni_row_max)
        End If
        gap_max = min_Long(max_Long(1, gap_max), i_max - 1)
        If start = -1 Then
            start = v_row_min
        Else
            start = min_Long(max_Long(1, v_row_min), max_Long(i_max - gap_max, 1))
        End If
    End If
' VARIABLE INSTANTIATION
    found_last = False
    If dir_move = FORWARDS Then
        i = start - 1
    Else
        i = i_max + 1
    End If
' METHODS
    Do While Not found_last
        i = i + dir_move
        If Orient = ColumnHeaders Then
            temp = Me.Cell(i, search_ID).value
        Else
            temp = Me.Cell(search_ID, i).value
        End If
        If (dir_move = FORWARDS And i = start) Or (dir_move = FORWARDS And i = start) Then
            found_last = True
        ElseIf IsEmpty(temp) Then
            gap_n = gap_n + 1
        Else
            gap_n = 0
        End If
        If gap_n > gap_max Then found_last = True
    Loop
' RETURNS
    LastFilledCell = i - gap_n * dir_move
End Function
