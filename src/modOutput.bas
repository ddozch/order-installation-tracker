Attribute VB_Name = "modOutput"
'==============================================================================
' Installation Tracker
' Writes the tables in one block each, keeps the notes and formats the result.
'==============================================================================
Option Explicit

Public Sub WriteAll(ByVal orderRows As Collection)
    Dim loSched As ListObject, loSpec As ListObject, loContract As ListObject, loOrders As ListObject
    Dim notesSched As Object, notesSpec As Object, notesContract As Object

    Set loSched = TableOf(SH_SCHED_ERR, TBL_SCHED_ERR)
    Set loSpec = TableOf(SH_SPEC_ERR, TBL_SPEC_ERR)
    Set loContract = TableOf(SH_CONTRACT_ERR, TBL_CONTRACT_ERR)
    Set loOrders = TableOf(SH_ORDERS, TBL_ORDERS)

    ' Notes are read before the issue tables are rebuilt
    Set notesSched = SaveNotes(loSched)
    Set notesSpec = SaveNotes(loSpec)
    Set notesContract = SaveNotes(loContract)

    FillTable loOrders, orderRows
    FillTable TableOf(SH_SPECS, TBL_SPECS), SpecRows
    FillTable loSched, ScheduleIssues
    FillTable loSpec, SpecIssues
    FillTable loContract, ContractIssues
    FillTable TableOf(SH_MANAGERS, TBL_MANAGERS), BuildManagers(orderRows)

    RestoreNotes loSched, notesSched, SH_SCHED_ERR
    RestoreNotes loSpec, notesSpec, SH_SPEC_ERR
    RestoreNotes loContract, notesContract, SH_CONTRACT_ERR

    FormatOrders loOrders
    FormatColumns TableOf(SH_SPECS, TBL_SPECS), "Spec Date", "yyyy-mm-dd"
    FormatColumns TableOf(SH_SPECS, TBL_SPECS), "Spec Month", "yyyy-mm"
    FormatIssues loSched
    FormatIssues loSpec
    FormatIssues loContract
End Sub

Public Sub WriteRunLog()
    Dim lo As ListObject
    Set lo = TableOf(SH_LOG, TBL_LOG)
    FillTable lo, RunLog
    FormatColumns lo, "Seconds", "0.00"
End Sub

Private Sub FormatOrders(ByVal lo As ListObject)
    Dim fc As FormatCondition
    FormatColumns lo, "Order Date,Planned Start,Planned Finish,Actual Start,Actual Finish,Contract Sent,Contract Signed", "yyyy-mm-dd"
    FormatColumns lo, "Order Month,Planned Month,Finish Month", "yyyy-mm"
    FormatColumns lo, "Planned Units,Installed Units,Spec Units", "#,##0"
    If lo.DataBodyRange Is Nothing Then Exit Sub
    With lo.ListColumns("Status").DataBodyRange
        .FormatConditions.Delete
        ' Plain value rules only: rules with functions are read in the language of the Excel interface
        Set fc = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & ST_DELAYED & """")
        fc.Interior.Color = ClrBad()
        Set fc = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & ST_TO_SCHEDULE & """")
        fc.Interior.Color = ClrWarn()
        Set fc = .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & ST_INSTALLED & """")
        fc.Interior.Color = ClrGood()
    End With
    lo.Range.Columns.AutoFit
End Sub

Private Sub FormatIssues(ByVal lo As ListObject)
    Dim nm As Variant
    FormatColumns lo, "Days", "0"
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.ListColumns("Note").DataBodyRange.Interior.Color = ClrInput()
    For Each nm In Array("Order ID", "Client", "Manager", "Crew", "Days")
        lo.ListColumns(CStr(nm)).Range.Columns.AutoFit
    Next nm
End Sub

Private Sub FormatColumns(ByVal lo As ListObject, ByVal columnNames As String, ByVal fmt As String)
    Dim nm As Variant
    If lo.DataBodyRange Is Nothing Then Exit Sub
    For Each nm In Split(columnNames, ",")
        lo.ListColumns(CStr(nm)).DataBodyRange.NumberFormat = fmt
    Next nm
End Sub
