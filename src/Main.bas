' Author: Edward Middleton-Smith
' Precision And Research Technology Systems Limited


' MODULE INITIALISATION
' Set array start index to 1 to match spreadsheet indices
Option Base 1
' Forced Variable Declaration
Option Explicit

' Sleep function
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub Main()
' FUNCTION
    ' Create database relationship bbetwee different tables of data in worksheets
' VARIABLE DECLARATION
    ' Excel
    Dim wb_me As Workbook
    Dim wb_cash As Workbook
    Dim wb_day As Workbook
    Dim wb_customers As Workbook
    Dim ws_out As Worksheet
    Dim ws_anal As Worksheet
    Dim ws_cash As Worksheet
    Dim ws_day As Worksheet
    Dim ws_customers As Worksheet
    ' Worksheet Container Relation
    Dim ws_rel As ws_relation
    ' Worksheet Containers
    Dim wsc_out As ws_access
    Dim wsc_cash As ws_access
    Dim wsc_day As ws_access
    Dim wsc_customers As ws_access
    Dim headings_S() As String
    ' PowerPoint
    Dim pptApp As PowerPoint.Application
    Dim ppt As PowerPoint.Presentation
    Dim pptSlide As PowerPoint.Slide
    Dim paste_shape As PowerPoint.Shape
    ' Temporary
    Dim headings_V() As Variant
    Dim i As Long
    Dim j As Long
    Dim tmp_total As Long
    Dim tmp_heading As String
    Dim i_col_out As Long
    Dim i_col_cash As Long
    Dim i_col_day As Long
    Dim tmp_S As String
    Dim t_sleep As Long
' VARIABLE INSTANTIATION
    t_sleep = 5000
    ' PowerPoint
    Set pptApp = New PowerPoint.Application
    Set ppt = pptApp.Presentations.Open("C:\Users\edwar\OneDrive\Documents\4 Shires\Alex Automation\4 Shires Books Public\Pivot Linked.pptx")
    Set pptSlide = ppt.Slides(2)
    pptSlide.Select
    Sleep t_sleep
    ' Worksheet Relation
    Set ws_rel = New ws_relation
    Set wb_me = ActiveWorkbook
    Set ws_out = create_sheet_out(wb_me)
    Set ws_anal = wb_me.Sheets("Analysis")
    get_downloaded_sheets wb_me, wb_cash, ws_cash, wb_day, ws_day, wb_customers, ws_customers, ws_rel, t_sleep
    ' Worksheet Containers
    ' Cashbook
    headings_V = Array("BATCH NO", "BATCH DATE", "TRAN DATE", "CUST REF", "CUSTOMER NAME", "", "TRAN REF", "FURTHER REF", "", "", "", "", "", "", "CASH", "DISCOUNT", "", "", "", "", "", "", "")
    convert_1D_Variant_2_String headings_V, headings_S
    Set wsc_cash = New ws_access
    wsc_cash.Init ws_cash, "Cashbook", headings_S, ColumnHeaders, 7, True, 3, False, 1, 0, 4
    ws_rel.AddWSC wsc_cash
    ' Daybook
    Erase headings_S
    Erase headings_V
    headings_V = Array("BATCH NO", "BATCH DATE", "TRAN DATE", "CUST REF", "CUST NAME", "CASH STATUS", "TRAN REF", "FURTHER REF", "", "GOODS", "VAT", "TOT INV", "TOT CRN", "ACC STAT", "", "", "", "", "", "", "", "", "")
    convert_1D_Variant_2_String headings_V, headings_S
    Set wsc_day = New ws_access
    wsc_day.Init ws_day, "Daybook", headings_S, ColumnHeaders, 7, True, 3, False, 1, 0, 4
    If Not ws_rel.AddWSC(wsc_day) Then GoTo errhand
    ' Customer List
    Erase headings_S
    Erase headings_V
    headings_V = Array("", "", "", "A/C", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "Name")
    convert_1D_Variant_2_String headings_V, headings_S
    Set wsc_customers = New ws_access
    wsc_customers.Init ws_customers, "Customer List", headings_S, ColumnHeaders, 4, True, 6, False, 2, 0, 7, 0, 10
    If Not ws_rel.AddWSC(wsc_customers) Then GoTo errhand
    ' Export
    ' Set ws_out = create_sheet_out()
    Erase headings_S
    Erase headings_V
    headings_V = Array("BATCH NO", "BATCH DATE", "Date", "Account Reference", "CUST NAME", "IMPORT ID NO.", "Reference", "Extra Reference", "User Name", "", "Tax Amount", "", "", "CASH STATUS", "", "", "Type", "Nominal A/C Ref", "Details", "Net Amount", "Tax Code", "TOT INV", "CUST NAME VLOOKUP")
    convert_1D_Variant_2_String headings_V, headings_S
    Set wsc_out = New ws_access
    wsc_out.Init ws_out, "Export", headings_S, ColumnHeaders, 7, True, 1, False, 1, 0, 2
    If Not ws_rel.AddWSC(wsc_out) Then GoTo errhand
    ' Worksheet size
    tmp_total = wsc_cash.RowMax - wsc_cash.RowMin + 1
' PROCESSING ACCELERATION
    ' Disable automatic spreadsheet calculation - prevents refreshing of whole ws on each cell entry
    Application.Calculation = xlCalculationManual
    ' Disable screen updating
    Application.ScreenUpdating = False
' METHODS
    ' Populate export wsc with references from cashbook, daybook
    wsc_out.ResizeLocal wsc_out.ColumnMin, wsc_out.ColumnMax, wsc_out.RowMin, wsc_cash.RowMax - wsc_cash.RowMin + wsc_day.RowMax - wsc_day.RowMin + wsc_out.RowMin + 1, True
    ' Cashbook
    For j = 1 To tmp_total
        ' ' Set key value for record-wise population
        ' wsc_out.SetCellDataLocal j, wsc_out.ColumnSearch, wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnSearch)
        
        ' Set all available values
        For i = 1 To ws_rel.nHeadings
            i_col_out = wsc_out.ColumnID(CStr(i), True)
            If i_col_out > 0 Then
                wsc_out.Cell(j + wsc_out.RowMin - 1, i_col_out + wsc_out.ColumnMin - 1).Interior.ColorIndex = 17 ' 6
                i_col_cash = wsc_cash.ColumnID(CStr(i), True)
                If i_col_cash > 0 Then
                    If tmp_heading = "Net Amount" Then
                        wsc_out.SetCellDataLocal j, i_col_out, CMoney_String(CDbl(wsc_cash.GetCellDataLocal(j, i_col_cash)))
                    Else
                        wsc_out.SetCellDataLocal j, i_col_out, wsc_cash.GetCellDataLocal(j, i_col_cash)
                    End If
                Else
                    tmp_heading = wsc_out.HeadingName(i)
                    If tmp_heading = "Type" Or tmp_heading = "Details" Then
                        ' pass
                    ElseIf tmp_heading = "Nominal A/C Ref" Then
                        tmp_S = Left(wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnID("FURTHER REF")), 2)
                        If tmp_S = "BA" Or tmp_S = "BX" Or tmp_S = "CQ" Then
                            wsc_out.SetCellDataLocal j, i_col_out, "1200"
                            wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Tax Code"), "T9"
                            If wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnID("FURTHER REF")) = "BX DUPLICATE ENTRY" Then
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "REFUND"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SP"
                            Else
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "PAYMENT RECEIVED"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SA"
                            End If
                        ElseIf tmp_S = "CA" Then
                            wsc_out.SetCellDataLocal j, i_col_out, "1230"
                            If CLng(wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnID("CASH"))) < 0 Then
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "REFUND"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SP"
                            Else
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "PAYMENT RECEIVED"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SA"
                            End If
                        ElseIf tmp_S = "CC" Or tmp_S = "WE" Then
                            wsc_out.SetCellDataLocal j, i_col_out, "1250"
                            If CLng(wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnID("CASH"))) < 0 Then
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "REFUND"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SP"
                            Else
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "PAYMENT RECEIVED"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SA"
                            End If
                        Else
                            wsc_out.SetCellDataLocal j, i_col_out, "1255"
                            If CLng(wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnID("CASH"))) < 0 Then
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "REFUND"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SP"
                            Else
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Details"), "PAYMENT RECEIVED"
                                wsc_out.SetCellDataLocal j, wsc_out.ColumnID("Type"), "SA"
                            End If
                        End If
                    ElseIf tmp_heading = "Net Amount" Then
                        wsc_out.SetCellDataLocal j, i_col_out, CMoney_String(Abs(CDbl(wsc_cash.GetCellDataLocal(j, wsc_cash.ColumnID("CASH")))))
                    ElseIf tmp_heading = "Tax Code" Then
                        wsc_out.SetCellDataLocal j, i_col_out, "T9"
                    ElseIf tmp_heading = "Tax Amount" Then
                        wsc_out.SetCellDataLocal j, i_col_out, "'0.00"
                    ElseIf tmp_heading = "User Name" Then
                        wsc_out.SetCellDataLocal j, i_col_out, "ALEX-IMPORTED"
                    End If
                End If
            End If
        Next
    Next
    ' Daybook
    For j = 1 To wsc_day.RowMax - wsc_day.RowMin + 1
        ' ' Set key value for record-wise population
        ' wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnSearch, wsc_day.GetCellDataLocal(j, wsc_day.ColumnSearch)
        
        ' Set all available values
        For i = 1 To ws_rel.nHeadings
            i_col_out = wsc_out.ColumnID(CStr(i), True)
            If i_col_out > 0 Then
                wsc_out.Cell(tmp_total + j + wsc_out.RowMin - 1, i_col_out + wsc_out.ColumnMin - 1).Interior.ColorIndex = 42 ' 50
                i_col_day = wsc_day.ColumnID(CStr(i), True)
                tmp_heading = wsc_out.HeadingName(i)
                If i_col_day > 0 Then
                    ' If tmp_heading = "Tax Amount" Or tmp_heading = "TOT INV" Then
                    '     wsc_out.SetCellDataLocal tmp_total + j, i_col_out, CMoney_String(CDbl(wsc_day.GetCellDataLocal(j, i_col_day)))
                    ' Else
                    wsc_out.SetCellDataLocal tmp_total + j, i_col_out, wsc_day.GetCellDataLocal(j, i_col_day)
                    ' End If
                Else
                    If tmp_heading = "Type" Then
                        ' pass
                        If CLng(wsc_day.GetCellDataLocal(j, wsc_day.ColumnID("GOODS"))) < 0 Then
                            wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnID("Type"), "SC"
                            wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnID("Details"), "CREDIT"
                        Else
                            wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnID("Type"), "SI"
                            wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnID("Details"), "INVOICE"
                        End If
                    ElseIf tmp_heading = "Nominal A/C Ref" Then
                        wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnSearch, "4000"
                    ElseIf tmp_heading = "Net Amount" Then
                        wsc_out.SetCellDataLocal tmp_total + j, i_col_out, CMoney_String(Abs(CDbl(wsc_day.GetCellDataLocal(j, wsc_day.ColumnID("GOODS")))))
                        wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnID("Tax Amount"), CMoney_String(Abs(CDbl(wsc_day.GetCellDataLocal(j, wsc_day.ColumnID("VAT")))))
                        wsc_out.SetCellDataLocal tmp_total + j, wsc_out.ColumnID("TOT INV"), CMoney_String(CDbl(wsc_day.GetCellDataLocal(j, wsc_day.ColumnID("GOODS"))) + CDbl(wsc_day.GetCellDataLocal(j, wsc_day.ColumnID("VAT"))))
                    ElseIf tmp_heading = "Tax Code" Then
                        wsc_out.SetCellDataLocal tmp_total + j, i_col_out, "T1"
                    ElseIf tmp_heading = "User Name" Then
                        wsc_out.SetCellDataLocal tmp_total + j, i_col_out, "ALEX-IMPORTED"
                    End If
                End If
            End If
        Next
    Next
    ' Customer name 'vlookup'
    wsc_out.ColumnSearch = wsc_out.ColumnID("Account Reference")
    ws_rel.Populate "Export", False, "Customer List", False
    wsc_out.ExportLocalData2WS
    
    ' ' Populate matching columns
    ' ws_rel.Populate "Export", False, "Cashbook", False, True
    ' ws_rel.Populate "Export", False, "Daybook", False, True
    
' RETURNS
    ' WS data
    wsc_out.ExportLocalData2WS
    ' table objects
    ws_out.ListObjects.Add(xlSrcRange, ws_out.Range("$A$1:$W$" & CStr(wsc_out.RowMax)), , xlYes).name = "tbl_audit" ' 'Audit Trail'!
    ws_anal.PivotTables("pivot_audit").ClearTable
    With ws_anal.PivotTables("pivot_audit")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ws_anal.PivotTables("pivot_audit").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ws_anal.PivotTables("pivot_audit").RepeatAllLabels xlRepeatLabels
    With ws_anal.PivotTables("pivot_audit").PivotFields("CUST NAME")
        .orientation = xlRowField
        .position = 1
    End With
    ws_anal.PivotTables("pivot_audit").AddDataField ws_anal.PivotTables( _
        "pivot_audit").PivotFields("Net Amount"), "Sum of Net Amount", xlSum
    ' Bring presentation to front
    ' AppActivate ppt.path
    ' Set ppt = pptApp.Presentations.Open("C:\Users\edwar\OneDrive\Documents\4 Shires\Alex Automation\4 Shires Books Public\Pivot Linked.pptx")
    ' PowerPoint.Presentations(ppt.path).Windows(1).Activate
    ws_anal.Range("A1:B27").Copy
    ' ppt.Slides(2).Shapes(1).TextFrame.TextRange.PasteSpecial
    pptSlide.Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile
    ' pptApp.ActiveWindow.View.PasteSpecial ppPasteEnhancedMetafile
    Set paste_shape = pptSlide.Shapes(pptSlide.Shapes.count)
    paste_shape.Left = paste_shape.Left - 150
    ws_anal.Range("A28:B54").Copy
    ' ppt.Slides(2).Shapes(1).TextFrame.TextRange.PasteSpecial
    pptSlide.Shapes.PasteSpecial DataType:=ppPasteEnhancedMetafile
    ' pptApp.ActiveWindow.View.PasteSpecial ppPasteEnhancedMetafile
    Set paste_shape = pptSlide.Shapes(pptSlide.Shapes.count)
    paste_shape.Left = paste_shape.Left + 150
    ' Add overdue statuses
    ws_anal.Cells(1, 3).value = "Overdue Status"
    ws_anal.Cells(1, 3).Interior.Color = RGB(209, 225, 239) ' "#D1E1EF"
    ws_anal.Cells(1, 3).Font.Bold = True
    ws_anal.Columns("C:C").ColumnWidth = 14.14
    For i = 2 To 53 Step 1
        If (i < 5) Then
            ws_anal.Cells(i, 3).value = OverdueStatusName(i - 2)
        Else
            ws_anal.Cells(i, 3).value = OverdueStatusName(CLng(Rnd() * 2))
        End If
    Next
    ' Emails
    ' MsgBox "Ready to send emails?"
    ' Sleep 25000
        ' Scroll through processed data, ppt
    mail_account_balances ws_anal, ws_customers
' ERROR HANDLING
errhand:
    On Error Resume Next
    wb_cash.Close
    wb_day.Close
    wb_customers.Close
' endgame:
' PROCESSING DECELARATION
    ' Enable automatic spreadsheet calculation - prevents refreshing of whole ws on each cell entry
    Application.Calculation = xlCalculationAutomatic
    ' Enable screen updating
    Application.ScreenUpdating = True
End Sub

Sub mail_account_balances(ByRef ws_anal As Worksheet, ByRef ws_customer As Worksheet)
' FUNCTION
    ' Contact all customers (as necessary) regarding their account balance
' CONSTANTS
    Const nCustomer = 52
' VARIABLE DECLARATION
    Dim olApp As Outlook.Application
    ' Dim olNS As Outlook.Namespace
    Dim ws_rel As ws_relation
    Dim wsc_anal As ws_access
    Dim wsc_customer As ws_access
    Dim headings_V() As Variant
    Dim headings_S() As String
    ' Iterables
    Dim iRow As Long
    Dim iRowCustomer As Long
    Dim colIDAnal_Company As Long
    Dim colIDAnal_Balance As Long
    Dim colIDAnal_status As Long
    Dim colIDCustomer_Contact As Long
    Dim colIDCustomer_Address As Long
' VARIABLE INSTANTIATION
    Set olApp = New Outlook.Application
    ' Set olNS = olApp.GetNamespace("MAPI")
    ' Worksheet Relation
    Set ws_rel = New ws_relation
    ' Worksheet Containers
    ' Customers
    headings_V = Array("A/C", "Name", "Contact Name", "Telephone", "Email", "", "")
    convert_1D_Variant_2_String headings_V, headings_S
    Set wsc_customer = New ws_access
    ' Init(ByRef ws As Worksheet, ByVal name As String, ByRef headings() As String, Optional Orient As orientation = orientation.ColumnHeaders, Optional search_col As Long = 1, Optional search_col_is_heading_index As Boolean = False, Optional search_row As Long = 1, Optional search_row_is_heading_index As Boolean = False, Optional col_min As Long = 1, Optional col_max As Long = 0, Optional row_min As Long = 2, Optional row_max As Long = 0, Optional gap_max As Long = 1, Optional mutable_headings As Boolean = False)
    ' wsc_customer.Init ws_customer, "Customers", headings_S, ColumnHeaders, 2, True, 6, False, 1, 0, 7
    wsc_customer.Init ws_customer, "Customer List", headings_S, ColumnHeaders, 2, True, 6, False, 2, 0, 7, 0, 10
    ws_rel.AddWSC wsc_customer
    ' Analysis
    Erase headings_S
    Erase headings_V
    headings_V = Array("", "Row Labels", "", "", "", "Sum of Net Amount", "Overdue Status")
    convert_1D_Variant_2_String headings_V, headings_S
    Set wsc_anal = New ws_access
    wsc_anal.Init ws_anal, "Analysis", headings_S, ColumnHeaders, 1, False, 1, False, 1, 0, 2
    If Not ws_rel.AddWSC(wsc_anal) Then GoTo errhand
    ' Indices
    colIDAnal_Company = wsc_anal.ColumnID("Row Labels")
    colIDAnal_Balance = wsc_anal.ColumnID("Sum of Net Amount")
    colIDAnal_status = wsc_anal.ColumnID("Overdue Status")
    colIDCustomer_Contact = wsc_customer.ColumnID("Contact Name")
    colIDCustomer_Address = wsc_customer.ColumnID("Email")
' METHODS
    For iRow = 2 To 5 Step 1 ' 1 + nCustomer
        iRowCustomer = wsc_customer.Match(wsc_anal.GetCellDataLocal(iRow - 1, colIDAnal_Company))
        mail_account_balance olApp, _
            wsc_customer.GetCellDataLocal(iRow - 1, colIDCustomer_Contact), _
            wsc_customer.GetCellDataLocal(iRow - 1, colIDCustomer_Address), _
            GetOverdueStatus(wsc_anal.GetCellDataLocal(iRow - 1, colIDAnal_status)), _
            CLng(wsc_anal.GetCellDataLocal(iRow - 1, colIDAnal_Balance))
    Next
' RETURNS
' ERROR HANDLING
errhand:
    Exit Sub
End Sub

Sub mail_account_balance(ByRef olApp As Outlook.Application, ByVal name As String, ByVal address As String, ByVal overdue_status As OverdueStatus, ByVal balance As Long)
' FUNCTION
    ' Send email to account holder regarding balance
' ARGUMENTS
    ' Outlook.Application olApp
    ' String name
    ' String address
    ' OverdueStatus overdue_status
    ' String balance
' VARIABLE DECLARATION
    Dim new_mail As Object
    Dim mail_body As String
    Dim subj As String
    
' VARIABLE INSTANTIATION
    Set new_mail = olApp.CreateItem(olMailItem)
    
    mail_body = "<html>"
    mail_body = mail_body & "<p>Dear " & name & ",</p>"
    
' METHODS
    Select Case overdue_status
        Case OverdueStatus.OVERDUE:
            mail_body = mail_body & "<p>Your account is overdue with a balance of £" & CStr(balance) & ".<br>"
            mail_body = mail_body & "Please resolve your balance as soon as possible.</p>"
            subj = "Overdue account balance"
        Case OverdueStatus.SUPEROVERDUE:
            mail_body = mail_body & "<p><span style=""color: red;""><strong>Your account is overdue with a balance of £" & CStr(balance) & ".</strong></span><br>"
            mail_body = mail_body & "Please resolve your balance as soon as possible.</p>"
            subj = "REMINDER: Overdue account balance"
        Case Else:
            mail_body = mail_body & "Your account has a negative balance of £" & CStr(balance) & ".<br>"
            mail_body = mail_body & "This balance is due to be settled by " & Format(DateSerial(2023, Month(Now()) + 1, -1), "Long Date") & ".</p>"
            subj = "Negative account balance"
    End Select
        
    mail_body = mail_body & "<p>Kind regards,<br>"
    mail_body = mail_body & "Lord Edward Middleton-Smith<br>"
    mail_body = mail_body & "Director<br>"
    mail_body = mail_body & "Precision And Research Technology Systems Limited"
    
' RETURNS
    With new_mail
        .To = address
        .Subject = subj
        .HTMLBody = mail_body
        .Display
    End With
End Sub

Sub get_downloaded_sheets(ByRef wb As Workbook, ByRef wb_cash As Workbook, ByRef ws_cash As Worksheet, ByRef wb_day As Workbook, ByRef ws_day As Worksheet, ByRef wb_customers As Workbook, ByRef ws_customers As Worksheet, ByRef ws_rel As ws_relation, ByVal t_sleep As Long)
' FUNCTION
    ' Get downloaded cashbook, daybook, and customer list
' ARGUMENTS
    ' Worksheet ws_cash
    ' Worksheet ws_day
    ' Worksheet ws_customers
    ' ws_relation ws_rel
' PROCESSING ACCELERATION
' CONSTANTS
' VARIABLE DECLARATION
    ' Dim wb_cash As Workbook
    ' Dim wb_day As Workbook
    ' Dim wb_customers As Workbook
    Dim fold As Scripting.Folder
    Dim f As Scripting.File
    Dim objFSO As Scripting.FileSystemObject
    Dim path_S As String
    Dim suffix As String
    Dim iTmp As Long
    Dim sTmp As String
' ARGUMENT VALIDATION
' VARIABLE INSTANTIATION
    Set objFSO = New Scripting.FileSystemObject
    path_S = GetLocalPath(wb.FullName)
    path_S = Left(path_S, Len(path_S) - Len(wb.name) - 1)
    If (Left(path_S, 41) = "https://d.docs.live.net/5728c5526437cee2/") Then
        path_S = "C:\Users\edwar\OneDrive\" & Mid(path_S, 42)
    End If
    Debug.Print path_S
    Set fold = objFSO.GetFolder(path_S)
' METHODS
    For Each f In fold.Files
        suffix = f.name
        iTmp = InStr(1, suffix, ".xl")
        If iTmp > 0 Then
            suffix = Mid(suffix, iTmp)
            suffix = Mid(suffix, 2, 2)
            If suffix = "xl" Then
            ' If f.Type = "Microsoft Macro-Enabled Workbook" Then
                sTmp = f.name
                If Not sTmp = wb.name Then
                    If Left(sTmp, 10) = "SL DAYBOOK" Then ' daybook
                        sTmp = f.path
                        Debug.Print sTmp
                        ' Set wb_day = Workbooks.Open(sTmp)
                        Set wb_day = ws_rel.SafeOpenWB(sTmp)
                        Set ws_day = wb_day.Sheets("Daybook")
                        ws_day.Activate
                        ws_day.Select
                        Sleep t_sleep
                    ElseIf Left(sTmp, 11) = "SL CASHBOOK" Then ' cashbook
                        sTmp = f.path
                        Debug.Print sTmp
                        ' Set wb_cash = Workbooks.Open(sTmp)
                        Set wb_cash = ws_rel.SafeOpenWB(sTmp)
                        Set ws_cash = wb_cash.Sheets("Cashbook")
                        ws_cash.Activate
                        ws_cash.Select
                        Sleep t_sleep
                    ElseIf sTmp = "Customers.xlsx" Then ' customer list
                        sTmp = f.path
                        Debug.Print sTmp
                        ' Set wb_customers = Workbooks.Open(sTmp)
                        Set wb_customers = ws_rel.SafeOpenWB(sTmp)
                        Set ws_customers = wb_customers.Sheets("Customer List")
                        ws_customers.Activate
                        ws_customers.Select
                        Sleep t_sleep
                    End If
                End If
            End If
        End If
    Next
' RETURNS
' ERROR HANDLING
' PROCESSING DECELARATION
End Sub

Function create_sheet_out(ByRef wb As Workbook) As Worksheet
' FUNCTION
    ' Create new output sheet
' ARGUMENTS
    ' Workbook wb
' PROCESSING ACCELERATION
' CONSTANTS
' VARIABLE DECLARATION
    Dim ws As Worksheet
    Dim headings_V() As Variant
    Dim colours_V() As Variant
    Dim i As Long
    Dim N As Long
    Dim found As Boolean
    Dim w As Long
' ARGUMENT VALIDATION
' VARIABLE INSTANTIATION
    headings_V = Array("Type", "Account Reference", "Nominal A/C Ref", "Department Code", "Date", "Reference", "Details", "Net Amount", "Tax Code", "Tax Amount", "Exchange Rate", "Extra Reference", "User Name", "Project Refn", "Cost Code Refn", "TOT INV", "IMPORT ID NO.", "BATCH NO", "BATCH DATE", "CASH STATUS", "NO", "CUST NAME", "CUST NAME VLOOKUP")
    ' colours_V = Array("50", "50", "50", "27", "50", "27", "27", "50", "50", "50", "27", "27", "27", "27", "27", "15", "15", "15", "15", "15", "15", "15", "15")
    colours_V = Array("42", "42", "42", "17", "42", "17", "17", "42", "42", "42", "17", "17", "17", "17", "17", "15", "15", "15", "15", "15", "15", "15", "15")
    N = SizeArrayDim_Variant(headings_V)
    Set ws = wb.Sheets.Add()
    ' Name ws
    found = True
    i = 0
    Do While found
        found = False
        For w = 1 To wb.Sheets.count
            If i = 0 Then
                If wb.Sheets.Item(w).name = "Audit Trail" Then
                    found = True
                    Exit For
                End If
            Else
                If wb.Sheets.Item(w).name = "Audit Trail " & CStr(i) Then
                    found = True
                    Exit For
                End If
            End If
        Next
        If found Then i = i + 1
    Loop
    If i = 0 Then
        ws.name = "Audit Trail"
    Else
        ws.name = "Audit Trail " & CStr(i)
    End If
' METHODS
    For i = 1 To N
        ws.Cells(1, i).value = headings_V(i)
        ws.Cells(1, i).Interior.ColorIndex = CLng(colours_V(i))
    Next
' RETURNS
    Set create_sheet_out = ws
' ERROR HANDLING
' PROCESSING DECELARATION
End Function

Function OverdueStatusName(ByVal overdue_status As OverdueStatus) As String
    Select Case overdue_status
        Case UNDUE:
            OverdueStatusName = "Not due"
        Case OVERDUE:
            OverdueStatusName = "Overdue"
        Case SUPEROVERDUE:
            OverdueStatusName = "Super Overdue"
    End Select
End Function

Function GetOverdueStatus(ByVal overdue_status As String) As OverdueStatus
    Select Case overdue_status
        Case "Not due":
            GetOverdueStatus = OverdueStatus.UNDUE
        Case "Overdue":
            GetOverdueStatus = OverdueStatus.OVERDUE
        Case "Super Overdue":
            GetOverdueStatus = OverdueStatus.SUPEROVERDUE
    End Select
End Function
