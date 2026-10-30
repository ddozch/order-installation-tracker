Attribute VB_Name = "modSetup"
'==============================================================================
' Installation Tracker
' Builds the workbook: sheets, tables and the dashboard. Run once after
' importing the modules (tools\build_workbook.ps1 does this for you).
'==============================================================================
Option Explicit

Public Sub SetupWorkbook()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    EnsureSheet SH_DASH
    SetupIssueSheet EnsureSheet(SH_SCHED_ERR), TBL_SCHED_ERR
    SetupIssueSheet EnsureSheet(SH_SPEC_ERR), TBL_SPEC_ERR
    SetupIssueSheet EnsureSheet(SH_CONTRACT_ERR), TBL_CONTRACT_ERR
    EnsureTable EnsureSheet(SH_MANAGERS), TBL_MANAGERS, ManagerHeaders()
    EnsureTable EnsureSheet(SH_ORDERS), TBL_ORDERS, OrderHeaders()
    EnsureTable EnsureSheet(SH_SPECS), TBL_SPECS, SpecHeaders()
    EnsureTable EnsureSheet(SH_LOG), TBL_LOG, LogHeaders()
    EnsureTable EnsureSheet(SH_ARCHIVE), TBL_ARCHIVE, ArchiveHeaders()
    SetupLists EnsureSheet(SH_LISTS)
    ' The dashboard formulas refer to the tables, so it is built last
    BuildDashboard ThisWorkbook.Worksheets(SH_DASH)
    RemoveEmptyExtraSheets
    HideGridlines ThisWorkbook.Worksheets(SH_DASH)
    ThisWorkbook.Worksheets(SH_DASH).Activate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set EnsureSheet = ws
End Function

Private Sub EnsureTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal headers As Variant)
    Dim lo As ListObject, n As Long
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0
    If Not lo Is Nothing Then Exit Sub

    ws.Cells.Clear
    n = UBound(headers) + 1
    ws.Range("A1").Resize(1, n).Value = headers
    Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range("A1").Resize(1, n), XlListObjectHasHeaders:=xlYes)
    lo.Name = tableName
    lo.TableStyle = "TableStyleMedium1"
    lo.Range.Columns.AutoFit
    FreezeHeader ws
End Sub

' Issue sheets: hidden Key column (it links notes to issues) and a wide Note column
Private Sub SetupIssueSheet(ByVal ws As Worksheet, ByVal tableName As String)
    EnsureTable ws, tableName, IssueHeaders()
    With ws.ListObjects(tableName)
        .ListColumns("Key").Range.EntireColumn.Hidden = True
        .ListColumns("Issue").Range.ColumnWidth = 24
        .ListColumns("Details").Range.ColumnWidth = 48
        .ListColumns("Note").Range.ColumnWidth = 40
    End With
End Sub

' Product list for the drop-down on the Dashboard
Private Sub SetupLists(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "All"
    ws.Range("A2").Formula2 = "=SORT(UNIQUE(tblOrders[Product]))"
    ws.Visible = xlSheetHidden
End Sub

' Removes the empty default sheet of a new workbook; sheets with content are kept
Private Sub RemoveEmptyExtraSheets()
    Dim i As Long, ws As Worksheet
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        Select Case ws.Name
            Case SH_DASH, SH_SCHED_ERR, SH_SPEC_ERR, SH_CONTRACT_ERR, SH_MANAGERS, SH_ORDERS, SH_SPECS, _
                 SH_LOG, SH_ARCHIVE, SH_LISTS
            Case Else
                If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
        End Select
    Next i
End Sub

' Window settings need an active window; skipped quietly when Excel runs hidden
Private Sub HideGridlines(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

Private Sub FreezeHeader(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    On Error GoTo 0
End Sub
