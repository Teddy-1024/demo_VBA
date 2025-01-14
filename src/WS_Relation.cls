' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit




' PROPERTY DECLARATION
Private v_wss As Collection
Private v_N_h As Long
Private v_name As String


' PROPERTY METHODS
Public Property Get nSheets() As Long
    nSheets = CountWorksheets
End Property

Public Property Get nHeadings() As Long
    nHeadings = v_N_h
End Property

Public Property Get wsc(ByVal sheet_name As String, Optional ID_in_name As Boolean = False) As ws_access
' FUNCTION
    ' Get Worksheet Container object by name or ID
' ARGUMENTS
    ' String sheet_name
    ' Boolean ID_in_name
' VARIABLE DECLARATION
    Dim sheet_ID As Long
    Dim n_ws As Long
    Dim w_i As Long
    Dim wsc_temp As ws_access
' VARIABLE INSTANTIATION
    n_ws = CountWorksheets
' METHODS
    If sheet_name = "" Then Exit Property
    If ID_in_name Then
        If Not is_numeric(sheet_name) Then Exit Property
        sheet_ID = CLng(sheet_name)
        If (sheet_ID < 1 Or sheet_ID > n_ws) Then
            Set wsc = v_wss.Item(sheet_ID)
            Exit Property
        End If
    End If
    For w_i = 1 To n_ws
        Set wsc_temp = v_wss.Item(w_i)
        If wsc_temp.name = sheet_name Then
            Set wsc = wsc_temp
            Exit Property
        End If
    Next
End Property

Public Property Get name() As String
    name = v_name
End Property
Public Property Let name(ByVal new_name As String)
    v_name = new_name
End Property


' METHODS
Public Function AddWSC(ByRef wsc As ws_access, Optional suppress_console_outputs As Boolean = False) As Boolean
' FUNCTION
    ' Get Workheet Container object by name or ID
' ARGUMENTS
    ' ws_access wsc
' VARIABLE DECLARATION
    Dim n_h As Long
    Dim n_ws As Long
' ARGUMENT VALIDATION
    AddWSC = False
    If wsc Is Nothing Then Exit Function
' VARIABLE INSTANTIATION
    n_h = wsc.nHeadings
    n_ws = CountWorksheets
' METHODS
    If n_ws = 0 Then
        v_N_h = n_h
        Set v_wss = New Collection
        v_wss.Add wsc
        If Not suppress_console_outputs Then Debug.Print "Worksheet Container " & wsc.name & " added to Worksheet Relation " & v_name & "."
    Else
        If v_N_h = n_h Then
            v_wss.Add wsc
            If Not suppress_console_outputs Then Debug.Print "Worksheet Container " & wsc.name & " added to Worksheet Relation " & v_name & "."
        Else
            MsgBox "Error: Unable to add Worksheet Container " & wsc.name & " to Worksheet Relation " & v_name & "." & vbCrLf & _
                CStr(v_N_h) & " headings in relation." & vbCrLf & _
                CStr(n_h) & " headings in container."
            Exit Function
        End If
    End If
' RETURNS
    AddWSC = True
End Function

Private Function CountWorksheets() As Long
    If v_wss Is Nothing Then
        CountWorksheets = 0
    Else
        CountWorksheets = v_wss.count
    End If
End Function

Public Sub Populate(Optional recipient_wsc As String = "1", Optional recipient_ID_in_str As Boolean = True, Optional provider_wsc As String = "2", Optional provider_ID_in_str As Boolean = True, Optional write_empties_only As Boolean = False, Optional export_2_ws As Boolean = True)
' FUNCTION
    ' Populate recipient wsc with data for each record by look up in provider wsc
' ARGUMENTS
    ' Optional String recipient_wsc
    ' Optional Boolean recipient_ID_in_str
    ' Optional String provider_wsc
    ' Optional Boolean provider_ID_in_str - ID instead of wsc name?
' VARIABLE DECLARATION
    Dim n_ws As Long
    Dim wsc_recipient As ws_access
    Dim wsc_provider As ws_access
    Dim h_i As Long
    Dim col_recipient As Long
    Dim col_provider As Long
    Dim row_recipient As Long
    Dim row_provider As Long
' ARGUMENT VALIDATION
    n_ws = CountWorksheets
    If (n_ws < 2 Or (recipient_wsc = provider_wsc And recipient_ID_in_str = provider_ID_in_str) Or Not ExistsWSC(recipient_wsc, recipient_ID_in_str) Or Not ExistsWSC(recipient_wsc, recipient_ID_in_str)) Then
        MsgBox "Error: Unable to populate Worksheet Container " & recipient_wsc & " with " & provider_wsc
        Exit Sub
    End If
' VARIABLE INSTANTIATION
    Set wsc_recipient = v_wss.Item(IndexWSC(recipient_wsc, recipient_ID_in_str))
    Set wsc_provider = v_wss.Item(IndexWSC(provider_wsc, provider_ID_in_str))
' METHODS
    If wsc_recipient.Orient = ColumnHeaders Then
        For row_recipient = 1 To wsc_recipient.RowMax - wsc_recipient.RowMin + 1
            row_provider = wsc_provider.Match(wsc_recipient.GetCellDataLocal(row_recipient, wsc_recipient.ColumnSearch))
            If row_provider > 0 Then
                For h_i = 1 To v_N_h
                    col_recipient = wsc_recipient.ColumnID(CStr(h_i), True)
                    col_provider = wsc_provider.ColumnID(CStr(h_i), True)
                    If col_provider > 0 And col_recipient > 0 Then
                        If col_recipient <> wsc_recipient.ColumnSearch Then
                            If wsc_provider.Orient = ColumnHeaders Then
                                wsc_recipient.SetCellDataLocal row_recipient, col_recipient, wsc_provider.GetCellDataLocal(row_provider, col_provider)
                            Else
                                wsc_recipient.SetCellDataLocal row_recipient, col_recipient, wsc_provider.GetCellDataLocal(col_provider, row_provider)
                            End If
                        End If
                    End If
                Next
            Else
                Debug.Print "Customer not found:" & vbCrLf & "Customer reference = " & wsc_recipient.GetCellDataLocal(row_recipient, wsc_recipient.ColumnSearch) & vbCrLf & "Row = " & CStr(row_recipient + wsc_recipient.RowMin - 1)
            End If
        Next
    Else
        For row_recipient = 1 To wsc_recipient.ColumnMax - wsc_recipient.ColumnMin + 1
            row_provider = wsc_provider.Match(wsc_recipient.GetCellDataLocal(wsc_recipient.RowSearch, row_recipient))
            If row_provider > 0 Then
                For h_i = 1 To v_N_h
                    col_recipient = wsc_recipient.ColumnID(CStr(h_i), True)
                    col_provider = wsc_provider.ColumnID(CStr(h_i), True)
                    If col_provider > 0 And col_recipient > 0 Then
                        If (wsc_recipient.Orient = ColumnHeaders And col_recipient <> wsc_recipient.ColumnSearch) Then
                            If wsc_provider.Orient = ColumnHeaders Then
                                wsc_recipient.SetCellDataLocal col_recipient, row_recipient, wsc_provider.GetCellDataLocal(row_provider, col_provider)
                            Else
                                wsc_recipient.SetCellDataLocal col_recipient, row_recipient, wsc_provider.GetCellDataLocal(col_provider, row_provider)
                            End If
                        End If
                    End If
                Next
            Else
                Debug.Print "Customer not found:" & vbCrLf & "Customer reference = " & wsc_recipient.GetCellDataLocal(wsc_recipient.RowSearch, row_recipient) & vbCrLf & "Column = " & CStr(row_recipient + wsc_recipient.ColumnMin - 1)
            End If
        Next
    End If
' RETURNS
    If export_2_ws Then wsc_recipient.ExportLocalData2WS
End Sub

Public Function ExistsWSC(ByVal wsc_name As String, Optional ID_in_str As Boolean = True) As Boolean
' FUNCTION
    ' Identify is Worksheet Container exists in relationship
' ARGUMENTS
    ' String wsc_name
    ' Optional Boolean ID_in_str - ID as String instead of name?
' VARIABLE DECLARATION
    Dim n_ws As Long
    Dim w_i As Long
    Dim ID As Long
' VARIABLE INSTANTIATION
    n_ws = CountWorksheets
    ExistsWSC = False
' ARGUMENT VALIDATION
    If wsc_name = "" Then Exit Function
' METHODS
    If ID_in_str Then
        If Not is_numeric(wsc_name) Then Exit Function
        ID = CLng(wsc_name)
        If Not (ID > 0 And ID <= n_ws) Then Exit Function
        ExistsWSC = True
        Exit Function
    End If
    For w_i = 1 To n_ws
        If v_wss.Item(w_i).name = wsc_name Then
            ExistsWSC = True
            Exit Function
        End If
    Next
' RETURNS
    MsgBox "Error: Worksheet Container " & wsc_name & " not found in relation."
End Function

Public Function IndexWSC(ByVal wsc_name As String, Optional ID_in_str As Boolean = True) As Long
' FUNCTION
    ' Get index for Worksheet Container within relationship
' ARGUMENTS
    ' String wsc_name
    ' Optional Boolean ID_in_str - ID as String instead of name?
' VARIABLE DECLARATION
    Dim n_ws As Long
    Dim w_i As Long
' VARIABLE INSTANTIATION
    n_ws = CountWorksheets
    IndexWSC = False
' ARGUMENT VALIDATION
    If Not ExistsWSC(wsc_name, ID_in_str) Then Exit Function
' METHODS
    If ID_in_str Then
        IndexWSC = CLng(wsc_name)
    Else
        For w_i = 1 To n_ws
            If v_wss.Item(w_i).name = wsc_name Then
                IndexWSC = w_i
                Exit For
            End If
        Next
    End If
End Function

Public Function SafeOpenWB(ByVal pathWB As String) As Workbook
' FUNCTION
    ' Open Workbook if not already open
' ARGUMENTS
    ' String pathWB
' VARIABLE DECLARATION
    Dim wb As Workbook
    Dim nameWB As String
    Dim nWB As Long
    Dim iWB As Long
' VARIABLE INSTANTIATION
    nameWB = Path2Name(pathWB)
    nWB = Workbooks.count
' ARGUMENT VALIDATION
    If nameWB = "" Then
        MsgBox "Error: Invalid workbook path"
        Exit Function
    End If
' METHODS
    For iWB = 1 To nWB
        Set wb = Workbooks.Item(iWB)
        If wb.name = nameWB Then
            Set SafeOpenWB = wb
            Exit Function
        End If
    Next
' RETURNS
    Set SafeOpenWB = Workbooks.Open(pathWB)
End Function

