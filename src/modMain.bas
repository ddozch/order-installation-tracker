Attribute VB_Name = "modMain"
'==============================================================================
' Installation Tracker
' One button runs the whole chain. Each step is timed in the Run Log.
'==============================================================================
Option Explicit

' Assigned to the Update button on the Dashboard
Public Sub UpdateTracker()
    RunUpdate True
End Sub

' interactive = False skips message boxes (used by tools\test_tracker.ps1)
Public Sub RunUpdate(ByVal interactive As Boolean)
    Dim started As Double, t As Double, n As Long, errText As String
    Dim dataFolder As String, v As Variant, orderRows As Collection

    On Error GoTo Fail
    started = Timer
    Set RunLog = New Collection
    Set ScheduleIssues = New Collection
    Set SpecIssues = New Collection
    Set ContractIssues = New Collection

    v = SettingValue(NM_REPORT_DATE)
    If VarType(v) <> vbDate And Not (VarType(v) = vbDouble And Not IsEmpty(v)) Then
        Err.Raise vbObjectError + 10, "RunUpdate", "Report date must be a date."
    End If
    ReportDate = DateSerial(Year(v), Month(v), Day(v))
    dataFolder = ResolvePath(CStr(SettingValue(NM_DATA_FOLDER)))
    SpeedOn

    t = Timer
    SetStatus "reading orders"
    LoadOrders dataFolder & "\" & FILE_ORDERS
    LogStep "Read orders", t, OrderList.Count

    t = Timer
    SetStatus "collecting crew schedules"
    n = CollectSchedules(dataFolder & "\" & DIR_SCHEDULES)
    LogStep "Collect crew schedules", t, n

    t = Timer
    SetStatus "reading contracts"
    LoadContracts dataFolder & "\" & FILE_CONTRACTS
    LogStep "Read contracts", t, Contracts.Count

    t = Timer
    SetStatus "reading specifications"
    LoadSpecs dataFolder & "\" & FILE_SPECS
    LogStep "Read specifications", t, SpecRows.Count

    t = Timer
    SetStatus "checking orders"
    Set orderRows = BuildOrders()
    LogStep "Status and checks", t, ScheduleIssues.Count + SpecIssues.Count + ContractIssues.Count

    t = Timer
    SetStatus "writing sheets"
    WriteAll orderRows
    LogStep "Write sheets, keep notes", t, orderRows.Count

    t = Timer
    SetStatus "recalculating"
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    LogStep "Recalculate dashboard", t, Empty

    LogStep "Total", started, Empty
    WriteRunLog
    ThisWorkbook.Names(NM_LAST_UPDATE).RefersToRange.Value = "Updated " & Format$(Now, "yyyy-mm-dd hh:nn") & _
        "  |  report date " & Format$(ReportDate, "yyyy-mm-dd") & "  |  " & Format$(Timer - started, "0.0") & " s"
    SpeedOff
    ThisWorkbook.Worksheets(SH_DASH).Activate

    If interactive Then
        MsgBox "Updated in " & Format$(Timer - started, "0.0") & " s." & vbCrLf & vbCrLf & _
               OrderList.Count & " orders, " & ScheduleIssues.Count + SpecIssues.Count + ContractIssues.Count & _
               " open issues.", vbInformation, "Installation Tracker"
    End If
    Exit Sub

Fail:
    errText = Err.Description
    Resume FailCleanup
FailCleanup:
    On Error Resume Next
    SpeedOff
    ThisWorkbook.Names(NM_LAST_UPDATE).RefersToRange.Value = "Update failed: " & errText
    On Error GoTo 0
    If interactive Then MsgBox "The update was not finished." & vbCrLf & vbCrLf & errText, vbExclamation, "Installation Tracker"
End Sub
