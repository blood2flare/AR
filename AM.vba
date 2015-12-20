Option Explicit
'Const file1251 = "C:\Users\Rudy Leo\AppData\Local\SAP\AGR_1251.XLSX"
'Const file1252 = "C:\Users\Rudy Leo\AppData\Local\SAP\AGR_1252.XLSX"
'Const fileTemp = "E:\AGR1251_TEMPLATE.XLSX"
'Const file1251 = "E:\MDNPUR 20151126\AGR1251_MDNPUR.XLSX"
'Const file1252 = "E:\MDNPUR 20151126\AGR1252_MDNPUR.XLSX"
'Const fileAgrs = "E:\MDNPUR 20151126\AGRAGRS_MDNPUR.XLSX"
'Const fileUsers = "E:\MDNPUR 20151126\AGRUSERS_MDNPUR.XLSX"
Dim fileTemp As String
Dim file1251 As String
Dim file1252 As String


Sub Create_TOC()
    Dim wbBook As Workbook
    Dim wsActive As Worksheet
    Dim wsSheet As Worksheet
    Dim lnRow As Long
    Dim lnPages As Long
    Dim lnCount As Long
    Set wbBook = ActiveWorkbook
    
    If Application.ScreenUpdating = True Then
        With Application
            .DisplayAlerts = False
            .ScreenUpdating = False
        End With
    End If
    'If the TOC sheet already exist delete it and add a new
    'worksheet.
    On Error Resume Next
    With wbBook
        .Worksheets("TOC").Delete
        .Worksheets.Add Before:=.Worksheets(1)
    End With
    On Error GoTo 0
    Set wsActive = wbBook.ActiveSheet
    With wsActive
        .Name = "TOC"
        With .Range("A1:B1")
            .Value = VBA.Array("Table of Contents", "Sheet # - # of Pages")
            .Font.Bold = True
        End With
    End With
    lnRow = 2
    lnCount = 1
    'Iterate through the worksheets in the workbook and create
    'sheetnames, add hyperlink and count & write the running number
    'of pages to be printed for each sheet on the TOC sheet.
    For Each wsSheet In wbBook.Worksheets
        If wsSheet.Name <> wsActive.Name Then
            wsSheet.Activate
            With wsActive
                .Hyperlinks.Add .Cells(lnRow, 1), "", _
                SubAddress:="'" & wsSheet.Name & "'!A1", _
                TextToDisplay:=wsSheet.Name
                'lnPages = wsSheet.PageSetup.Pages().Count
                '.Cells(lnRow, 2).Value = "'" & lnCount & "-" & lnPages
                wsSheet.Hyperlinks.Add wsSheet.Cells(1, 6), "", _
                    SubAddress:="'" & wsActive.Name & "'!A1", TextToDisplay:="Back"
            End With
            lnRow = lnRow + 1
            lnCount = lnCount + 1
        End If
    Next wsSheet
    wsActive.Activate
    wsActive.Columns("A:B").EntireColumn.AutoFit
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

Sub genFuncTemp()
    Dim wbFunc As Workbook
    Dim wsFunc As Worksheet
    Dim strPath As String
    Dim intChoice As Integer
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
    Application.FileDialog(msoFileDialogOpen).Title = "Open Dump AGR_1251 for Single Role"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        file1251 = strPath
    Else
        MsgBox "Program terminated"
        Exit Sub
    End If
    Application.FileDialog(msoFileDialogOpen).Title = "Open Dump AGR_1252 for Derived Role"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        file1252 = strPath
    Else
        MsgBox "Program terminated"
        Exit Sub
    End If
    
    Set wbFunc = Workbooks.Add
    With wbFunc
        .Title = "Functional Role List"
        .Subject = "Auth Revamp"
    End With
    
    parse1251Single wbFunc
    parse1252Derive wbFunc, "S"
    parse1252Derive wbFunc, "D"
    'ParseAgrs wbFunc
    
    If wbFunc.Worksheets.Count > 1 Then wbFunc.Worksheets("Sheet1").Delete
    wbFunc.Worksheets(1).Activate
    Create_TOC
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

Sub genTempList()
    Dim wbTemp As Workbook
    Dim wsTemp As Worksheet
    Dim strPath As String
    Dim intChoice As Integer
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    Application.FileDialog(msoFileDialogOpen).Title = "Open Dump AGR_1251 for Template Role"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
        fileTemp = strPath
    Else
        MsgBox "Program terminated"
        Exit Sub
    End If

    Set wbTemp = Workbooks.Add
    With wbTemp
        .Title = "Template Role List"
        .Subject = "Auth Revamp"
    End With
    
    parse1251Template wbTemp

    If wbTemp.Worksheets.Count > 1 Then wbTemp.Worksheets("Sheet1").Delete
    wbTemp.Worksheets(1).Activate
    Create_TOC

    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
End Sub

Sub parse1251Single(wbFunc As Workbook)
    Const c1251RoleCol = 1
    Const c1251ObjectCol = 3
    Const c1251AuthCol = 4
    Const c1251FieldCol = 6
    Const c1251LowCol = 7
    Const c1251HighCol = 8
    Const c1251DeleteCol = 10
    Const cFuncSingleTableHeaderRow = 5
    Const cFuncSingleHeaderTcodeStartRow = 3
    Const cFuncSingleHeaderTcodeCol = 1
    Const cFuncSingleHeaderRoleNameRow = 1
    Const cFuncSingleHeaderRoleNameCol = 1
    Const cSingleFieldClassCol = 1
    Const cSingleFieldObjectCol = 2
    Const cSingleFieldFieldCol = 3
    Const cSingleFieldLowCol = 4
    Const cSingleFieldHighCol = 5
    Const cSingleFieldAuthCol = 6
    Const cHeaderColorIndex = 25
    
    Dim wb1251 As Workbook
    Dim ws1251 As Worksheet
    Dim rg1251 As Range
    Dim wsFunc As Worksheet
        
    Dim ln1251Row As Long
    Dim lnFuncRow As Long
    Dim lnFuncTableHeaderRow As Long
    Dim strRoleName As String
    Dim tmpLastRole As String
    Dim intTcodeCount As Integer
    
    Set wsFunc = wbFunc.ActiveSheet
    Set wb1251 = Workbooks.Open(file1251)
    Set ws1251 = wb1251.Worksheets("Sheet1")
    ws1251.Activate
    
    Set rg1251 = ws1251.Range("A2").CurrentRegion
    'On Error Resume Next
    If ActiveWorkbook.Worksheets("Sheet1").AutoFilterMode Then
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    Else
        rg1251.AutoFilter
    End If
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .SortFields.Add Key:=Range("A2", Cells(rg1251.Rows.Count, 1)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("C2", Cells(rg1251.Rows.Count, 3)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("D2", Cells(rg1251.Rows.Count, 4)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("F2", Cells(rg1251.Rows.Count, 6)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("G2", Cells(rg1251.Rows.Count, 7)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("H2", Cells(rg1251.Rows.Count, 8)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'On Error GoTo 0
    
    ln1251Row = 2
    tmpLastRole = ""
    While ws1251.Cells(ln1251Row, c1251RoleCol).Value <> ""
        strRoleName = ws1251.Cells(ln1251Row, c1251RoleCol).Value
        If Mid(strRoleName, 1, 1) = "S" And ws1251.Cells(ln1251Row, c1251DeleteCol).Value = "" Then
            If tmpLastRole <> strRoleName Then
                'Create New Sheet
                On Error Resume Next
                wsFunc.Columns("A:Z").EntireColumn.AutoFit
                wsFunc = wbFunc.Worksheets(strRoleName)
                If Err.Number <> 0 Then
                    'Worksheet does not exist
                    wbFunc.Worksheets.Add after:=wbFunc.Worksheets(wbFunc.Worksheets.Count)
                    wbFunc.Worksheets(wbFunc.Worksheets.Count).Name = strRoleName
                    lnFuncRow = cFuncSingleTableHeaderRow + 1
                    Set wsFunc = wbFunc.ActiveSheet
                End If
                On Error GoTo 0
                'Filling Header
                wsFunc.Cells(cFuncSingleHeaderRoleNameRow, cFuncSingleHeaderRoleNameCol).Value = "Role :"
                wsFunc.Cells(cFuncSingleHeaderRoleNameRow, cFuncSingleHeaderRoleNameCol + 1).Value = strRoleName
                wsFunc.Cells(cFuncSingleHeaderTcodeStartRow, cFuncSingleHeaderTcodeCol).Value = "Tcode :"
                intTcodeCount = 0
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldClassCol)
                    .Value = "Class"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldObjectCol)
                    .Value = "Object"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldFieldCol)
                    .Value = "Field"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldLowCol)
                    .Value = "Low"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldHighCol)
                    .Value = "High"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldAuthCol)
                    .Value = "Auth"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                intTcodeCount = 0
            End If
            'Processing Single Item
            If ws1251.Cells(ln1251Row, c1251DeleteCol).Value = "" Then
                If ws1251.Cells(ln1251Row, c1251ObjectCol).Value = "S_TCODE" Then
                    If intTcodeCount = 0 Then
                        wsFunc.Cells(cFuncSingleHeaderTcodeStartRow, cFuncSingleHeaderTcodeCol + 1).Value = ws1251.Cells(ln1251Row, c1251LowCol).Value
                    ElseIf intTcodeCount > 0 Then
                        wsFunc.Cells(cFuncSingleHeaderTcodeStartRow + intTcodeCount, cFuncSingleHeaderTcodeCol + 1).EntireRow.Insert
                        wsFunc.Cells(cFuncSingleHeaderTcodeStartRow + intTcodeCount, cFuncSingleHeaderTcodeCol + 1).Value = ws1251.Cells(ln1251Row, c1251LowCol).Value
                        lnFuncRow = lnFuncRow + 1
                    End If
                    intTcodeCount = intTcodeCount + 1
                Else
                    wsFunc.Cells(lnFuncRow, cSingleFieldObjectCol).Value = ws1251.Cells(ln1251Row, c1251ObjectCol).Value
                    wsFunc.Cells(lnFuncRow, cSingleFieldFieldCol).Value = ws1251.Cells(ln1251Row, c1251FieldCol).Value
                    With wsFunc.Cells(lnFuncRow, cSingleFieldLowCol)
                        .NumberFormat = "@"
                        .Value = ws1251.Cells(ln1251Row, c1251LowCol).Value
                    End With
                    wsFunc.Cells(lnFuncRow, cSingleFieldHighCol).Value = ws1251.Cells(ln1251Row, c1251HighCol).Value
                    wsFunc.Cells(lnFuncRow, cSingleFieldAuthCol).Value = ws1251.Cells(ln1251Row, c1251AuthCol).Value
                    lnFuncRow = lnFuncRow + 1
                End If
            End If
            tmpLastRole = strRoleName
        End If
        ln1251Row = ln1251Row + 1
        Debug.Assert ln1251Row < 300000
    Wend
    wsFunc.Columns("A:Z").EntireColumn.AutoFit
    On Error GoTo 0
    wb1251.Close
End Sub


Sub parse1251Template(wbFunc As Workbook)
    Const c1251RoleCol = 1
    Const c1251ObjectCol = 3
    Const c1251AuthCol = 4
    Const c1251FieldCol = 6
    Const c1251LowCol = 7
    Const c1251HighCol = 8
    Const c1251DeleteCol = 10
    Const cFuncSingleTableHeaderRow = 5
    Const cFuncSingleHeaderTcodeStartRow = 3
    Const cFuncSingleHeaderTcodeCol = 1
    Const cFuncSingleHeaderRoleNameRow = 1
    Const cFuncSingleHeaderRoleNameCol = 1
    Const cSingleFieldClassCol = 1
    Const cSingleFieldObjectCol = 2
    Const cSingleFieldFieldCol = 3
    Const cSingleFieldLowCol = 4
    Const cSingleFieldHighCol = 5
    Const cSingleFieldAuthCol = 6
    Const cHeaderColorIndex = 25
    
    Dim wb1251 As Workbook
    Dim ws1251 As Worksheet
    Dim rg1251 As Range
    Dim wsFunc As Worksheet
        
    Dim ln1251Row As Long
    Dim lnFuncRow As Long
    Dim lnFuncTableHeaderRow As Long
    Dim strRoleName As String
    Dim tmpLastRole As String
    Dim intTcodeCount As Integer
    
    Set wsFunc = wbFunc.ActiveSheet
    Set wb1251 = Workbooks.Open(fileTemp)
    Set ws1251 = wb1251.Worksheets("Sheet1")
    ws1251.Activate
    
    Set rg1251 = ws1251.Range("A2").CurrentRegion
    'On Error Resume Next
    If ActiveWorkbook.Worksheets("Sheet1").AutoFilterMode Then
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    Else
        rg1251.AutoFilter
    End If
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .SortFields.Add Key:=Range("A2", Cells(rg1251.Rows.Count, 1)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("C2", Cells(rg1251.Rows.Count, 3)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("D2", Cells(rg1251.Rows.Count, 4)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("F2", Cells(rg1251.Rows.Count, 6)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("G2", Cells(rg1251.Rows.Count, 7)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("H2", Cells(rg1251.Rows.Count, 8)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'On Error GoTo 0
    
    ln1251Row = 2
    tmpLastRole = ""
    While ws1251.Cells(ln1251Row, c1251RoleCol).Value <> ""
        strRoleName = ws1251.Cells(ln1251Row, c1251RoleCol).Value
        If Mid(strRoleName, 1, 1) = "T" And ws1251.Cells(ln1251Row, c1251DeleteCol).Value = "" Then
            If tmpLastRole <> strRoleName Then
                'Create New Sheet
                On Error Resume Next
                wsFunc.Columns("A:Z").EntireColumn.AutoFit
                wsFunc = wbFunc.Worksheets(strRoleName)
                If Err.Number <> 0 Then
                    'Worksheet does not exist
                    wbFunc.Worksheets.Add after:=wbFunc.Worksheets(wbFunc.Worksheets.Count)
                    wbFunc.Worksheets(wbFunc.Worksheets.Count).Name = strRoleName
                    lnFuncRow = cFuncSingleTableHeaderRow + 1
                    Set wsFunc = wbFunc.ActiveSheet
                End If
                On Error GoTo 0
                'Filling Header
                wsFunc.Cells(cFuncSingleHeaderRoleNameRow, cFuncSingleHeaderRoleNameCol).Value = "Role :"
                wsFunc.Cells(cFuncSingleHeaderRoleNameRow, cFuncSingleHeaderRoleNameCol + 1).Value = strRoleName
                wsFunc.Cells(cFuncSingleHeaderTcodeStartRow, cFuncSingleHeaderTcodeCol).Value = "Tcode :"
                intTcodeCount = 0
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldClassCol)
                    .Value = "Class"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldObjectCol)
                    .Value = "Object"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldFieldCol)
                    .Value = "Field"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldLowCol)
                    .Value = "Low"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldHighCol)
                    .Value = "High"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                With wsFunc.Cells(cFuncSingleTableHeaderRow, cSingleFieldAuthCol)
                    .Value = "Auth"
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                intTcodeCount = 0
            End If
            'Processing Single Item
            If ws1251.Cells(ln1251Row, c1251DeleteCol).Value = "" Then
                If ws1251.Cells(ln1251Row, c1251ObjectCol).Value = "S_TCODE" Then
                    If intTcodeCount = 0 Then
                        wsFunc.Cells(cFuncSingleHeaderTcodeStartRow, cFuncSingleHeaderTcodeCol + 1).Value = ws1251.Cells(ln1251Row, c1251LowCol).Value
                    ElseIf intTcodeCount > 0 Then
                        wsFunc.Cells(cFuncSingleHeaderTcodeStartRow + intTcodeCount, cFuncSingleHeaderTcodeCol + 1).EntireRow.Insert
                        wsFunc.Cells(cFuncSingleHeaderTcodeStartRow + intTcodeCount, cFuncSingleHeaderTcodeCol + 1).Value = ws1251.Cells(ln1251Row, c1251LowCol).Value
                        lnFuncRow = lnFuncRow + 1
                    End If
                    intTcodeCount = intTcodeCount + 1
                Else
                    wsFunc.Cells(lnFuncRow, cSingleFieldObjectCol).Value = ws1251.Cells(ln1251Row, c1251ObjectCol).Value
                    wsFunc.Cells(lnFuncRow, cSingleFieldFieldCol).Value = ws1251.Cells(ln1251Row, c1251FieldCol).Value
                    With wsFunc.Cells(lnFuncRow, cSingleFieldLowCol)
                        .NumberFormat = "@"
                        .Value = ws1251.Cells(ln1251Row, c1251LowCol).Value
                    End With
                    wsFunc.Cells(lnFuncRow, cSingleFieldHighCol).Value = ws1251.Cells(ln1251Row, c1251HighCol).Value
                    wsFunc.Cells(lnFuncRow, cSingleFieldAuthCol).Value = ws1251.Cells(ln1251Row, c1251AuthCol).Value
                    lnFuncRow = lnFuncRow + 1
                End If
            End If
            tmpLastRole = strRoleName
        End If
        ln1251Row = ln1251Row + 1
        Debug.Assert ln1251Row < 300000
    Wend
    wsFunc.Columns("A:Z").EntireColumn.AutoFit
    On Error GoTo 0
    wb1251.Close
End Sub


Sub parse1252Derive(wbFunc As Workbook, strFlag As String)
    Const c1252RoleCol = 1
    Const c1252FieldCol = 3
    Const c1252LowCol = 4
    Const c1252HighCol = 5
    Const cFuncDeriveHeaderRoleNameRow = 1
    Const cFuncDeriveHeaderRoleNameCol = 1
    Const cFuncDeriveHeaderTemplateNameRow = 1
    Const cFuncDeriveHeaderTemplateNameCol = 3
    Const cFuncDeriveFieldRow = 5
    Const cFuncDeriveStartCol = 2
    Const cHeaderColorIndex = 25
    
    Dim wsFunc As Worksheet
    Dim wb1252 As Workbook
    Dim ws1252 As Worksheet
    Dim rg1252 As Range
    Dim intOffset As Integer
    
    Dim ln1252Row As Long
    Dim lnFuncRow As Long
    Dim lnFuncCol As Long
    Dim strRoleName As String
    Dim strLastRole As String
    Dim strFieldName As String
    Dim strLastField As String
    
    Set wsFunc = wbFunc.ActiveSheet
    Set wb1252 = Workbooks.Open(file1252)
    Set ws1252 = wb1252.Worksheets("Sheet1")
    ws1252.Activate
    
    Set rg1252 = ws1252.Range("A2").CurrentRegion
    'On Error Resume Next
    If ActiveWorkbook.Worksheets("Sheet1").AutoFilterMode Then
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    Else
        rg1252.AutoFilter
    End If
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .SortFields.Add Key:=Range("A2", Cells(rg1252.Rows.Count, 1)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("C2", Cells(rg1252.Rows.Count, 3)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("D2", Cells(rg1252.Rows.Count, 4)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=Range("E2", Cells(rg1252.Rows.Count, 5)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'On Error GoTo 0

    ln1252Row = 2
    strLastRole = ""
    While ws1252.Cells(ln1252Row, c1252RoleCol).Value <> ""
        strRoleName = ws1252.Cells(ln1252Row, c1252RoleCol).Value
        If Mid(strRoleName, 1, 1) = strFlag Then
        'If strRoleName = "D01MM_11_MDNPUR_PO_CHA" Then
            If strFlag = "S" Then
                intOffset = 7
            Else
                intOffset = 0
            End If
            'Err = Nothing
            If strLastRole <> strRoleName Then
                'Create New Sheet
                On Error Resume Next
                wsFunc.Columns("A:Z").EntireColumn.AutoFit
                Set wsFunc = wbFunc.Worksheets(strRoleName)
                If Err.Number <> 0 Then
                    'Worksheet does not exist
                    wbFunc.Worksheets.Add after:=wbFunc.Worksheets(wbFunc.Worksheets.Count)
                    wbFunc.Worksheets(wbFunc.Worksheets.Count).Name = strRoleName
                    'lnFuncRow = cFuncSingleTableHeaderRow + 1
                    Set wsFunc = wbFunc.ActiveSheet
                    lnFuncCol = 1
                    strLastField = ""
                    'Filling Header
                    wsFunc.Cells(cFuncDeriveHeaderRoleNameRow, cFuncDeriveHeaderRoleNameCol).Value = "Role :"
                    wsFunc.Cells(cFuncDeriveHeaderRoleNameRow, cFuncDeriveHeaderRoleNameCol + 1).Value = strRoleName
                    wsFunc.Cells(cFuncDeriveHeaderTemplateNameRow, cFuncDeriveHeaderTemplateNameCol).Value = "Template :"
                    wsFunc.Cells(cFuncDeriveHeaderTemplateNameRow, cFuncDeriveHeaderTemplateNameCol + 1).Value = "'=TODO="
                    wsFunc.Cells(cFuncDeriveFieldRow, cFuncDeriveStartCol - 1).Value = "VALUE"
                    wsFunc.Cells(cFuncDeriveFieldRow - 2, cFuncDeriveStartCol).Value = "FIELD"
                Else
                    lnFuncCol = 1
                    strLastField = ""
                    wsFunc.Cells(cFuncDeriveFieldRow, cFuncDeriveStartCol - 1 + intOffset).Value = "VALUE"
                    wsFunc.Cells(cFuncDeriveFieldRow - 2, cFuncDeriveStartCol + intOffset).Value = "FIELD"
                End If
                On Error GoTo 0
            End If
            'Processing Single Item
            strFieldName = ws1252.Cells(ln1252Row, c1252FieldCol).Value
            If strLastField <> strFieldName Then
                lnFuncCol = lnFuncCol + 1
                With wsFunc.Cells(cFuncDeriveFieldRow, lnFuncCol + intOffset)
                    .Value = Mid(strFieldName, 2, 100)
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                lnFuncRow = cFuncDeriveFieldRow + 1
                strLastField = strFieldName
            End If
            wsFunc.Cells(lnFuncRow, lnFuncCol + intOffset).NumberFormat = "@"
            wsFunc.Cells(lnFuncRow, lnFuncCol + intOffset).Value = ws1252.Cells(ln1252Row, c1252LowCol).Value
            lnFuncRow = lnFuncRow + 1
            strLastRole = strRoleName
        End If
        ln1252Row = ln1252Row + 1
        Debug.Assert ln1252Row < 300000
    Wend
    wsFunc.Columns("A:Z").EntireColumn.AutoFit
    On Error GoTo 0
    wb1252.Close
End Sub

Sub UserTemplate_BuildSummary()
    Dim wbBook As Workbook
    Dim wsSummary As Worksheet
    Dim wsJobAll As Worksheet
    Dim wsSheet As Worksheet
    Dim lnRow As Long
    Dim lnPages As Long
    Dim lnCount As Long
    
    Const masterJobCol = 4          'Shortname
    Const masterLoginCol = 5
    Const detJobRow = 4
    Const detIDRow = 5
    Const detAuthRow = 6
    Const detAuthColInit = 5
    Dim tmpRow As Long
    Dim detAuthCol As Long
    Dim detAuthName As String
    Dim detJob As Variant
    'Dim masterJob As Dictionary
    Dim masterJob As Variant
    Dim detAuthJob As Variant
    'Dim detAuthRec As Dictionary
    Dim detAuthRec As Variant
    'Dim detAuthGroup As Dictionary
    Dim detAuthGroup As Variant
    'Dim detAuthPos As Dictionary
    Dim detAuthPos As Variant
    'Dim detIDRec As Dictionary
    Dim detIDRec As Variant
    
    Dim tmpJob() As String
    Dim tmpJobCount As Integer
    Dim tmpI As Integer
    
    Set wbBook = ActiveWorkbook
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    On Error Resume Next
    With wbBook
        .Worksheets("Summary").Delete
        .Worksheets.Add Before:=.Worksheets(1)
    End With
    On Error GoTo 0
    
    Set wsSummary = wbBook.ActiveSheet
    With wsSummary
        .Name = "Summary"
        .Range("C1").Value = "Authorization Matrix"
    End With
    lnRow = 2
    lnCount = 1
    
    'Collect All Jobrole
    wbBook.Worksheets("Login List").Activate
    Set wsJobAll = wbBook.ActiveSheet
    tmpI = 1
    On Error Resume Next
    Set masterJob = CreateObject("Scripting.Dictionary")
    With wsSummary
        'Job Role
        tmpRow = 2
        While wsJobAll.Cells(tmpRow, masterJobCol).Value <> ""
            If masterJob.Exists(wsJobAll.Cells(tmpRow, masterJobCol).Value) = False Then
                masterJob.Add wsJobAll.Cells(tmpRow, masterJobCol).Value, tmpI
                tmpI = tmpI + 1
                .Cells(1, tmpI + 2).Value = wsJobAll.Cells(tmpRow, masterJobCol).Value
            End If
            tmpRow = tmpRow + 1
        Wend
        
        'Login
        tmpRow = 2
        While wsJobAll.Cells(tmpRow, masterLoginCol).Value <> ""
            If masterJob.Exists(wsJobAll.Cells(tmpRow, masterLoginCol).Value) = False Then
                masterJob.Add Trim(wsJobAll.Cells(tmpRow, masterLoginCol).Value), tmpI
                tmpI = tmpI + 1
                .Cells(1, tmpI + 2).Value = wsJobAll.Cells(tmpRow, masterLoginCol).Value
            End If
            tmpRow = tmpRow + 1
        Wend
    End With
    On Error GoTo 0
    
    'Get Detail Auth
    tmpRow = 2
    For Each wsSheet In wbBook.Worksheets
        If wsSheet.Name <> "TOC" And wsSheet.Name <> "Summary" And wsSheet.Name <> "How To Map The Template." _
                                 And wsSheet.Name <> "Login List" And wsSheet.Name <> "Summary0" _
                                 And wsSheet.Name <> "Legend" And wsSheet.Name <> "Unlisted Tcode" _
                                 And wsSheet.Name <> "User Asignment" _
                                 And wsSheet.Name <> "PR" And wsSheet.Name <> "RFQ" Then
        'If wsSheet.Name = "POSPOSTO" Then
            wsSheet.Activate
            'Get Authorization
            detAuthName = ""
            detAuthCol = detAuthColInit
            Set detAuthRec = CreateObject("Scripting.Dictionary")
            Set detAuthGroup = CreateObject("Scripting.Dictionary")
            Set detAuthPos = CreateObject("Scripting.Dictionary")
            Set detIDRec = CreateObject("Scripting.Dictionary")
            detAuthRec.RemoveAll
            Dim tmpJ As Integer
            Dim tmpStr As String
            tmpJ = 1
            Do
                With wsSheet
                    detAuthName = .Cells(detAuthRow, detAuthCol).Value
                    If detAuthName <> "" And .Cells(detJobRow, detAuthCol).Value <> "" Then
                        Set detAuthJob = New Collection
                        tmpStr = Replace(.Cells(detJobRow, detAuthCol).Value, "(", ",")
                        tmpStr = Replace(tmpStr, ")", ",")
                        tmpJob = Split(tmpStr, ",")
                        For tmpI = 0 To Application.CountA(tmpJob) - 1
                            detAuthJob.Add Trim(tmpJob(tmpI))
                        Next tmpI
                        detAuthRec.Add detAuthName, detAuthJob
                        detAuthGroup.Add detAuthName, wsSheet.Name
                        detAuthPos.Add detAuthName, .Cells(detAuthRow, detAuthCol).Address(ReferenceStyle:=xlA1)
                        detIDRec.Add detAuthName, .Cells(detIDRow, detAuthCol).Value
                    End If
                End With
                detAuthCol = detAuthCol + 1
                tmpJ = tmpJ + 1
            Loop Until detAuthName = ""
            'Loop Until tmpJ = 7
            
            'With wsSummary
            '    .Hyperlinks.Add .Cells(lnRow, 1), _
            '        "", SubAddress:="'" & wsSheet.Name & "'!A1", _
            '        TextToDisplay:=wsSheet.Name
            '    'lnPages = wsSheet.PageSetup.Pages().Count
            '    .Cells(lnRow, 2).Value = "'" & lnCount '& "-" & lnPages
            'End With
            'lnRow = lnRow + 1
            'lnCount = lnCount + 1
            Dim tmpVar As Variant
            With wsSummary
                For Each tmpVar In detAuthRec.Keys
                    Set wsSheet = wbBook.Worksheets(detAuthGroup(tmpVar))
                    .Cells(tmpRow, 1).Value = detAuthGroup(tmpVar)
                    .Cells(tmpRow, 2).Value = detIDRec(tmpVar)
                    '.Cells(tmpRow, 2).Value = tmpVar
                    .Hyperlinks.Add .Cells(tmpRow, 3), "", _
                                    SubAddress:="'" & detAuthGroup(tmpVar) & "'!" & detAuthPos(tmpVar), _
                                    TextToDisplay:=tmpVar
                    .Hyperlinks.Add wsSheet.Cells(1, 4), "", _
                                    SubAddress:="'Summary'!A1", _
                                    TextToDisplay:="<= Summary"
                    wsSheet.Range("D1").HorizontalAlignment = xlRight
                    wsSheet.Range("D1").Font.Bold = True
                    For Each detAuthJob In detAuthRec(tmpVar)
                        If masterJob(detAuthJob) = 0 Then
                            'With .Cells(tmpRow, 2)
                            '    .Font.Bold = True
                            '    .Font.ColorIndex = 3     'Red
                            '    .Interior.ColorIndex = 6 'Yellow
                            'End With
                        Else
                            'With .Cells(tmpRow, 2)
                            '    If .Font.ColorIndex = 3 Then
                            '        .Font.Bold = False
                            '        .Font.ColorIndex = xlColorIndexNone
                            '        .Interior.ColorIndex = xlColorIndexNone
                            '    End If
                            'End With
                            With .Cells(tmpRow, masterJob(detAuthJob) + 3)
                                .Value = "X"
                                .Font.Bold = True
                            End With
                        End If
                    Next
                    tmpRow = tmpRow + 1
                Next
            End With
        End If
    Next wsSheet
    wsSummary.Columns.EntireColumn.AutoFit
    wsSummary.Rows.EntireRow.AutoFit
    wsSummary.Columns(2).EntireColumn.Hidden = True
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    wsSummary.Activate
            
    Set masterJob = Nothing
    Set detAuthRec = Nothing
    Set detAuthGroup = Nothing
    Set detAuthPos = Nothing
    Set detIDRec = Nothing
End Sub


Sub UserTemplate_ResetFormatting()
    Const colorDarkPerc = 40
    Const colorMedPerc = 55
    Const colorLightPerc = 95
    Const colorLightAPerc = 90
    'Const colorAuthField = RGB(0, 254, 0)
    'Const colorTcode = RGB(254, 0, 0)
    'Const colorFontUserDefined = RGB(0, 0, 254)
    'Const colorBackUserDefined = RGB(254, 254, 254)
    Dim colorFontUserDefined As Long
    Dim colorBackUserDefined As Long
    
    Dim wbUser As Workbook
    Dim wsSheet As Worksheet
    
    Dim intColor As Integer
    Dim lnRow As Long
    Dim lnCol As Long
    Dim lnTmpRow As Long
    Dim lnLastRow As Long
    Dim lnLastCol As Long
    Dim lnTcodeRow As Long
    Dim flagAlt As Boolean
    
    colorFontUserDefined = RGB(0, 0, 165)
    colorBackUserDefined = RGB(254, 254, 254)
        
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    Set wbUser = ActiveWorkbook
    For Each wsSheet In wbUser.Worksheets
        If wsSheet.Name <> "How To Map The Template." And wsSheet.Name <> "User Asignment" And _
           wsSheet.Name <> "Login List" And wsSheet.Name <> "Summary0" And wsSheet.Name <> "Summary" And _
           wsSheet.Name <> "Legend" And wsSheet.Name <> "Unlisted Tcode" And wsSheet.Visible = xlSheetVisible Then
        'If wsSheet.Name = "SO" Then
            With wsSheet
                .Activate
                .Cells.Font.Name = "Calibri Light"
                .Cells.Font.Size = 11
                .Cells.SpecialCells(xlLastCell).Select
                lnLastRow = Range(Selection.Address).Row
                lnLastCol = Range(Selection.Address).Column
                
                .Cells(1, 1).Value = "Location :"
                .Cells(1, 1).Font.Bold = True
                .Cells(2, 1).Value = "Department :"
                .Cells(2, 1).Font.Bold = True
                .Cells(4, 1).Value = "Job Roles / Login :"
                .Cells(4, 1).HorizontalAlignment = xlRight
                .Cells(4, 1).Font.Bold = True
                .Range("A4:D4").Merge
                .Cells(5, 1).Value = "Authorization ID :"
                .Cells(5, 1).HorizontalAlignment = xlRight
                .Cells(5, 1).Font.Bold = True
                .Range("A5:D5").Merge
                .Cells(6, 1).Value = "Authorization Name :"
                .Cells(6, 1).HorizontalAlignment = xlRight
                .Cells(6, 1).Font.Bold = True
                .Range("A6:D6").Merge
                intColor = colorLightPerc / 100 * 255 - 1
                .Range("A1:E6").Interior.Color = RGB(intColor, intColor, intColor)
                .Range("C1:C2,E4:E6").Interior.Color = colorBackUserDefined
                .Range("C1:C2,E4:E6").Font.Color = colorFontUserDefined
                .Range("C1:C2,E4:E6").Font.Bold = True
                .Range("A4:E6").Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range("A4:E6").Borders(xlEdgeTop).Weight = xlMedium
                .Range("A4:E6").Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Range("A4:E6").Borders(xlInsideHorizontal).Weight = xlThin
                .Range("A4:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range("A4:E6").Borders(xlEdgeBottom).Weight = xlMedium
                .Range("E4:E6").HorizontalAlignment = xlCenter
                
                .Cells(7, 1).Value = "Authorization Object Value"
                .Cells(7, 1).HorizontalAlignment = xlLeft
                .Cells(7, 1).Font.Bold = True
                intColor = colorLightPerc / 100 * 255 - 1
                .Range(.Cells(7, 1), .Cells(7, 5)).Interior.Color = RGB(intColor, 254, intColor)
                .Range("A7:D7").Merge
                .Cells(8, 1).Value = "Organizational Level"
                .Cells(8, 1).HorizontalAlignment = xlLeft
                .Cells(8, 1).Font.Bold = True
                .Cells(8, 4).Value = "Format"
                .Cells(8, 4).Font.Bold = True
                intColor = colorDarkPerc / 100 * 255 - 1
                .Range(.Cells(8, 1), .Cells(8, 5)).Interior.Color = RGB(0, intColor, 0)
                .Range(.Cells(8, 1), .Cells(8, 5)).Font.Color = RGB(254, 254, 254)
                .Range(.Cells(8, 1), .Cells(8, 5)).Font.Bold = True
                lnRow = 9
                lnTmpRow = 9
                While Left(.Cells(lnRow, 1).Value, 10) <> "Non-Organi"
                    lnRow = lnRow + 1
                    'Stop if this value exceeded
                    Debug.Assert lnRow < 1000
                Wend
                'flagAlt = False
                If lnRow > lnTmpRow Then
                    intColor = colorDarkPerc / 100 * 255 - 1
                    .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Interior.Color = RGB(0, intColor, 0)
                    .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Font.Color = RGB(254, 254, 254)
                    .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Font.Bold = True
                    'If flagAlt = False Then
                        intColor = colorLightPerc / 100 * 255 - 1
                    '    flagAlt = True
                    'Else
                    '    intColor = colorLightAPerc / 100 * 255 - 1
                    '    flagAlt = False
                    'End If
                    With .Range(.Cells(lnTmpRow, 1), .Cells(lnRow - 1, 4))
                        .Interior.Color = RGB(intColor, 254, intColor)
                    '    .Borders(xlEdgeTop).Weight = xlThin
                    '    .Borders(xlEdgeLeft).Weight = xlThin
                    '    .Borders(xlEdgeRight).Weight = xlThin
                    '    .Borders(xlEdgeBottom).Weight = xlThin
                    '    .Borders(xlInsideHorizontal).Weight = xlThin
                    '    .Borders(xlInsideVertical).Weight = xlThin
                        .HorizontalAlignment = xlLeft
                    End With
                    With .Range(.Cells(lnTmpRow, 5), .Cells(lnRow - 1, 5))
                        .Interior.Color = colorBackUserDefined
                        .Font.Color = colorFontUserDefined
                    '    .Borders(xlEdgeTop).Weight = xlThin
                    '    .Borders(xlEdgeLeft).Weight = xlThin
                    '    .Borders(xlEdgeRight).Weight = xlThin
                    '    .Borders(xlEdgeBottom).Weight = xlThin
                    '    .Borders(xlInsideHorizontal).Weight = xlThin
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                'With .Range(.Cells(7, 1), .Cells(lnRow, 4))
                '    .Font.Color = RGB(0, 0, 0)
                '    intColor = colorLightPerc / 100 * 255 - 1
                '    .Interior.Color = RGB(intColor, 254, intColor)
                '    .HorizontalAlignment = xlLeft
                'End With
                'With .Range(.Cells(7, 5), .Cells(lnRow, 5))
                '    .Interior.Color = colorBackUserDefined
                '    .Font.Color = colorFontUserDefined
                '    .HorizontalAlignment = xlCenter
                'End With
                lnTmpRow = lnRow + 1
                While .Cells(lnRow, 1).Value <> "Transaction"
                    lnRow = lnRow + 1
                    'Stop if this value exceeded
                    Debug.Assert lnRow < 1000
                Wend
                With .Range(.Cells(7, 1), .Cells(lnRow, 5))
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Weight = xlMedium
                    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
                    .Borders(xlInsideHorizontal).Weight = xlThin
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = xlMedium
                End With
                lnTcodeRow = lnRow + 1
                If lnRow > lnTmpRow Then
                    .Cells(lnRow, 4).Value = "Description / Activity"
                    .Cells(lnRow + 1, 4).Value = ""
                    intColor = colorDarkPerc / 100 * 255 - 1
                    .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Interior.Color = RGB(intColor, 0, 0)
                    .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Font.Color = RGB(254, 254, 254)
                    .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Font.Bold = True
                    intColor = colorLightPerc / 100 * 255 - 1
                    .Range(.Cells(lnTmpRow, 1), .Cells(lnRow - 1, 4)).Interior.Color = RGB(intColor, 254, intColor)
                    With .Range(.Cells(lnTmpRow, 5), .Cells(lnRow - 1, 5))
                        .Interior.Color = colorBackUserDefined
                        .Font.Color = colorFontUserDefined
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Font.Color = RGB(0, 0, 0)
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Borders(xlEdgeTop).LineStyle = xlContinuous
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Borders(xlEdgeTop).Weight = xlMedium
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Borders(xlInsideHorizontal).Weight = xlThin
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).Borders(xlEdgeBottom).Weight = xlMedium
                .Range(.Cells(lnTcodeRow, 7), .Cells(5, lnRow)).HorizontalAlignment = xlCenter
                lnRow = lnRow + 1
                lnTmpRow = lnRow
                While lnRow <= lnLastRow
                    If .Cells(lnRow, 2).Value <> "" And .Cells(lnRow, 4).Value = "" Then
                        intColor = colorMedPerc / 100 * 255 - 1
                        .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Interior.Color = RGB(254, intColor, intColor)
                        .Range(.Cells(lnRow, 1), .Cells(lnRow, 5)).Font.Bold = True
                    Else
                        intColor = colorLightPerc / 100 * 255 - 1
                        .Range(.Cells(lnRow, 1), .Cells(lnRow, 4)).Interior.Color = RGB(254, intColor, intColor)
                        With .Range(.Cells(lnRow, 5), .Cells(lnRow, 5))
                            .Interior.Color = colorBackUserDefined
                            .Font.Color = colorFontUserDefined
                            .Font.Bold = False
                            .HorizontalAlignment = xlCenter
                        End With
                    End If
                    lnRow = lnRow + 1
                    'Stop if this value exceeded
                    Debug.Assert lnRow < 1000
                Wend
                lnCol = 5
                While lnCol <= lnLastCol
                    lnRow = lnTcodeRow
                    While lnRow <= lnLastRow
                        If .Cells(lnRow, lnCol).Value = "X" Then
                            .Rows(lnRow).Font.Bold = True
                        End If
                        lnRow = lnRow + 1
                    Wend
                    If lnCol > 5 Then
                        Columns("E:E").Select
                        Selection.Copy
                        Columns(lnCol).Select
                        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False
                    End If
                    lnCol = lnCol + 1
                    'Stop if this value exceeded
                    Debug.Assert lnCol < 1000
                Wend
                .Columns.EntireColumn.AutoFit
                .Rows.EntireRow.AutoFit
                .Rows.EntireRow.VerticalAlignment = xlTop
                .Columns(1).ColumnWidth = 5
            End With
        End If
    Next

    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    UserTemplate_BuildSummary
    ActiveSheet.Activate
End Sub

Private Sub ParseAgrs(wbFunc As Workbook)
    Dim wbAgrs As Workbook
    Dim wsFunc As Worksheet
    Dim wsAgrs As Worksheet
    Dim strLastRole As String

    Set wsFunc = wbFunc.ActiveSheet
    Set wbAgrs = Workbooks.Open(fileAgrs)
    Set wsAgrs = wbAgrs.Worksheets("Sheet1")
    wsAgrs.Activate
    
    Set rgAgrs = wsAgrs.Range("A2").CurrentRegion
    'On Error Resume Next
    If ActiveWorkbook.Worksheets("Sheet1").AutoFilterMode Then
        ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    Else
        rgAgrs.AutoFilter
    End If
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .SortFields.Add Key:=Range("A2", Cells(rgAgrs.Rows.Count, 1)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    'On Error GoTo 0

    lnAgrsRow = 2
    strLastRole = ""
    While wsAgrs.Cells(lnAgrsRow, cAgrsRoleCol).Value <> ""
        strRoleName = wsAgrs.Cells(lnAgrsRow, cAgrsRoleCol).Value
        If Mid(strRoleName, 1, 1) = strFlag Then
        'If strRoleName = "D01MM_11_MDNPUR_PO_CHA" Then
            If strFlag = "S" Then
                intOffset = 7
            Else
                intOffset = 0
            End If
            'Err = Nothing
            If strLastRole <> strRoleName Then
                'Create New Sheet
                On Error Resume Next
                wsFunc.Columns("A:Z").EntireColumn.AutoFit
                Set wsFunc = wbFunc.Worksheets(strRoleName)
                If Err.Number <> 0 Then
                    'Worksheet does not exist
                    wbFunc.Worksheets.Add after:=wbFunc.Worksheets(wbFunc.Worksheets.Count)
                    wbFunc.Worksheets(wbFunc.Worksheets.Count).Name = strRoleName
                    'lnFuncRow = cFuncSingleTableHeaderRow + 1
                    Set wsFunc = wbFunc.ActiveSheet
                    lnFuncCol = 1
                    strLastField = ""
                    'Filling Header
                    wsFunc.Cells(cFuncDeriveHeaderRoleNameRow, cFuncDeriveHeaderRoleNameCol).Value = "Role :"
                    wsFunc.Cells(cFuncDeriveHeaderRoleNameRow, cFuncDeriveHeaderRoleNameCol + 1).Value = strRoleName
                    wsFunc.Cells(cFuncDeriveHeaderTemplateNameRow, cFuncDeriveHeaderTemplateNameCol).Value = "Template :"
                    wsFunc.Cells(cFuncDeriveHeaderTemplateNameRow, cFuncDeriveHeaderTemplateNameCol + 1).Value = "'=TODO="
                    wsFunc.Cells(cFuncDeriveFieldRow, cFuncDeriveStartCol - 1).Value = "VALUE"
                    wsFunc.Cells(cFuncDeriveFieldRow - 2, cFuncDeriveStartCol).Value = "FIELD"
                Else
                    lnFuncCol = 1
                    strLastField = ""
                    wsFunc.Cells(cFuncDeriveFieldRow, cFuncDeriveStartCol - 1 + intOffset).Value = "VALUE"
                    wsFunc.Cells(cFuncDeriveFieldRow - 2, cFuncDeriveStartCol + intOffset).Value = "FIELD"
                End If
                On Error GoTo 0
            End If
            'Processing Single Item
            strFieldName = wsAgrs.Cells(lnAgrsRow, cAgrsFieldCol).Value
            If strLastField <> strFieldName Then
                lnFuncCol = lnFuncCol + 1
                With wsFunc.Cells(cFuncDeriveFieldRow, lnFuncCol + intOffset)
                    .Value = Mid(strFieldName, 2, 100)
                    .Interior.ColorIndex = cHeaderColorIndex
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                lnFuncRow = cFuncDeriveFieldRow + 1
                strLastField = strFieldName
            End If
            wsFunc.Cells(lnFuncRow, lnFuncCol + intOffset).NumberFormat = "@"
            wsFunc.Cells(lnFuncRow, lnFuncCol + intOffset).Value = wsAgrs.Cells(lnAgrsRow, cAgrsLowCol).Value
            lnFuncRow = lnFuncRow + 1
            strLastRole = strRoleName
        End If
        lnAgrsRow = lnAgrsRow + 1
        Debug.Assert lnAgrsRow < 300000
    Wend
    wsFunc.Columns("A:Z").EntireColumn.AutoFit
    On Error GoTo 0
    wbAgrs.Close
End Sub

Sub Collect_Plant()
    Dim dtPlant As Dictionary
    Dim dtPlantRec As Dictionary
    Dim wbBook As Workbook
    Dim wsSheet As Worksheet
    Dim lnFirstRec As Long
    Dim lnLastRec As Long
    Dim Content As Variant
    Dim i As Integer
    Dim j As Integer
    Dim Entry As Variant
    Dim Detail As Variant
    Dim Search As Variant
    Dim detailOK As Boolean
    Dim searchOK As Boolean
    Dim dtSearch As Dictionary
    
    Set wbBook = ActiveWorkbook
    Set wsSheet = wbBook.ActiveSheet
    
    Set dtPlant = New Dictionary
    lnFirstRec = 5
    lnLastRec = wsSheet.Cells(lnFirstRec, 1).End(xlDown).Row
    Content = wsSheet.Range(Cells(lnFirstRec, 1), Cells(lnLastRec, 9)).Value
    For i = 1 To UBound(Content, 1)
        If dtPlant.Exists(Content(i, 1)) Then
            'Skip duplicate
        Else
            Set dtPlantRec = New Dictionary
            'Set dtPlantRec = CreateObject("Scripting.Dictionary")
            For j = 4 To 9
                dtPlantRec.Add wsSheet.Cells(4, j).Value, Content(i, j)
            Next j
            dtPlant.Add Content(i, 1), dtPlantRec
        End If
    Next
    
    Set dtSearch = New Dictionary
    With dtSearch
        .Add "Acronym", "[ALL]"
        .Add "Group", "[ALL]"
        .Add "Location", "[ALL]"
        .Add "Name", "CUSTODY"
        .Add "PRJ", "[ALL]"
        .Add "Status", "[ALL]"
    End With
    i = 0
    For Each Entry In dtPlant.Keys
        detailOK = False
        For Each Detail In dtPlant(Entry).Keys
            searchOK = True
            For Each Search In dtSearch.Keys
                searchOK = searchOK And (dtSearch(Search) = "[ALL]" Or dtSearch(Search) = dtPlant(Entry)(Detail))
            Next
            detailOK = detailOK Or searchOK
        Next
        If detailOK Then
            i = i + 1
        End If
    Next
    MsgBox i
    
End Sub

Sub UserTemplate_CollectFieldList()
    Const cAuthRow = 6
    Const cAuthColStart = 5
    Const cFieldRowStart = 9
    Const cFieldCol = 1
    
    Dim wbUser As Workbook
    Dim wsUser As Worksheet
    Dim wbThis As Workbook
    Dim wsThis As Worksheet
    Dim rgThis As Range
    Dim dtExcludedSheet As Dictionary
    Dim dtAuth As Dictionary
    Dim dtAuthDet As Dictionary
    Dim lnAuthCount As Long
    Dim lnRow, lnCol As Long
    Dim intChoice As Integer
    Dim strPath As String
    Dim strFieldCache As String
    Dim strFieldCheck As String
    Dim strField As String
    Dim strContent As String
    
    With Application
        .DisplayAlerts = False
        .ScreenUpdating = False
    End With
    
    Set wbThis = ActiveWorkbook
    Set wsThis = ActiveWorkbook.ActiveSheet
    If Selection.Cells.Count > 1 Then
        MsgBox "Please select only 1 cell containing field to check", vbCritical, "Error"
        Exit Sub
    Else
        strFieldCheck = Selection.Cells.Value
        MsgBox "Checking " & strFieldCheck & ".."
    End If
    
    Application.FileDialog(msoFileDialogOpen).Title = "Open User Template"
    intChoice = Application.FileDialog(msoFileDialogOpen).Show
    If intChoice <> 0 Then
        strPath = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
    Else
        MsgBox "Program terminated"
        Exit Sub
    End If
    Set wbUser = Workbooks.Open(strPath)
    
    Set dtExcludedSheet = New Dictionary
    dtExcludedSheet.Add "Summary", "Summary"
    dtExcludedSheet.Add "How To Map The Template.", "How To Map The Template."
    dtExcludedSheet.Add "Login List", "Login List"
    dtExcludedSheet.Add "Summary0", "Summary0"
    
    Set dtAuth = New Dictionary
    
    For Each wsUser In wbUser.Worksheets
        If dtExcludedSheet.Exists(wsUser.Name) Or wsUser.Visible = False Then
            'Skip excluded sheets
        Else
            'Get Auth
            lnRow = cFieldRowStart
            strField = wsUser.Cells(lnRow, cFieldCol).Value
            While strField <> "Transaction"
                If strField <> strFieldCheck And strField <> "" Then
                    strFieldCache = strField
                End If
                If strFieldCache = strFieldCheck Then
                    lnCol = cAuthColStart
                    While wsUser.Cells(cAuthRow, lnCol).Value <> ""
                        strContent = wsUser.Cells(lnRow, lnCol).Value
                        If strContent <> "" Then
                            If Not dtAuth.Exists(wsUser.Cells(cAuthRow, lnCol).Value) Then
                                'Indiscriminate Collection
                                dtAuth.Add wsUser.Cells(cAuthRow, lnCol).Value, wsUser.Cells(lnRow, lnCol).Value
                                'TODO: Check collection
                            End If
                        End If
                        lnCol = lnCol + 1
                        Debug.Assert lnCol < 1000
                    Wend
                    
                End If
                lnRow = lnRow + 1
                strField = wsUser.Cells(lnRow, cFieldCol).Value
                '#TEST
                Debug.Assert lnRow < 1000
            Wend
            
        End If
    Next
    
    With Application
        .DisplayAlerts = True
        .ScreenUpdating = True
    End With
    

End Sub

Sub UserTemplate_CheckData()

End Sub
