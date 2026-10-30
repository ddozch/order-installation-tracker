Attribute VB_Name = "modConfig"
'==============================================================================
' Installation Tracker
' Names, business rules, record layouts and the state of the current update.
'==============================================================================
Option Explicit

' Sheets
Public Const SH_DASH As String = "Dashboard"
Public Const SH_SCHED_ERR As String = "Schedule Issues"
Public Const SH_SPEC_ERR As String = "Spec Issues"
Public Const SH_CONTRACT_ERR As String = "Contract Issues"
Public Const SH_MANAGERS As String = "Managers"
Public Const SH_ORDERS As String = "Orders"
Public Const SH_SPECS As String = "Specs"
Public Const SH_LOG As String = "Run Log"
Public Const SH_ARCHIVE As String = "Notes Archive"
Public Const SH_LISTS As String = "Lists"

' Tables
Public Const TBL_ORDERS As String = "tblOrders"
Public Const TBL_SPECS As String = "tblSpecs"
Public Const TBL_SCHED_ERR As String = "tblScheduleIssues"
Public Const TBL_SPEC_ERR As String = "tblSpecIssues"
Public Const TBL_CONTRACT_ERR As String = "tblContractIssues"
Public Const TBL_MANAGERS As String = "tblManagers"
Public Const TBL_LOG As String = "tblRunLog"
Public Const TBL_ARCHIVE As String = "tblNotesArchive"

' Named cells on the Dashboard
Public Const NM_REPORT_DATE As String = "cfgReportDate"
Public Const NM_YEAR As String = "selYear"
Public Const NM_PRODUCT As String = "selProduct"
Public Const NM_DATA_FOLDER As String = "cfgDataFolder"
Public Const NM_LAST_UPDATE As String = "outLastUpdate"

' Source files in the data folder
Public Const FILE_ORDERS As String = "orders.xlsx"
Public Const FILE_CONTRACTS As String = "contracts.xlsx"
Public Const FILE_SPECS As String = "specifications.xlsx"
Public Const DIR_SCHEDULES As String = "schedules"

' Business rules
Public Const CONTRACT_SIGN_DAYS As Long = 14   ' a contract should be signed within 14 days of sending
Public Const SPEC_DUE_DAYS As Long = 14        ' a specification is due 14 days after installation

' Order statuses
Public Const ST_INSTALLED As String = "Installed"
Public Const ST_IN_PROGRESS As String = "In progress"
Public Const ST_TO_SCHEDULE As String = "To schedule"
Public Const ST_DELAYED As String = "Delayed start"
Public Const ST_SCHEDULED As String = "Scheduled"

' Issue types
Public Const IS_NOT_SCHEDULED As String = "Not scheduled"
Public Const IS_NO_PLAN_DATES As String = "No planned dates"
Public Const IS_NO_ACTUAL_DATES As String = "No actual dates"
Public Const IS_FINISH_BEFORE_START As String = "Finish before start"
Public Const IS_START_DELAYED As String = "Start delayed"
Public Const IS_DUPLICATE As String = "Duplicate schedule"
Public Const IS_UNKNOWN_ORDER As String = "Unknown order"
Public Const IS_BAD_VALUE As String = "Unreadable value"
Public Const IS_NO_SPEC As String = "No specification"
Public Const IS_SPEC_INCOMPLETE As String = "Specification incomplete"
Public Const IS_SIGNED_LATE As String = "Signed late"
Public Const IS_NOT_SIGNED As String = "Not signed, overdue"
Public Const IS_NO_CONTRACT As String = "No contract"

' Order record from orders.xlsx (0-based)
Public Const O_ID As Long = 0
Public Const O_DATE As Long = 1
Public Const O_CLIENT As Long = 2
Public Const O_PRODUCT As Long = 3
Public Const O_UNITS As Long = 4
Public Const O_PMONTH As Long = 5
Public Const O_MANAGER As Long = 6

' Schedule record from a crew file
Public Const S_CREW As Long = 0
Public Const S_PSTART As Long = 1
Public Const S_PFINISH As Long = 2
Public Const S_ASTART As Long = 3
Public Const S_AFINISH As Long = 4
Public Const S_UNITS As Long = 5
Public Const S_FILE As Long = 6
Public Const S_ROW As Long = 7

' Contract record
Public Const C_TYPE As Long = 0
Public Const C_SENT As Long = 1
Public Const C_SIGNED As Long = 2

' Positions in a row of the Orders table, same order as OrderHeaders()
Public Const OR_MANAGER As Long = 7
Public Const OR_STATUS As Long = 15

' Positions in a row of an issue table, same order as IssueHeaders()
Public Const I_KEY As Long = 0
Public Const I_ORDER As Long = 1
Public Const I_CLIENT As Long = 2
Public Const I_MANAGER As Long = 3

' State of the current update
Public ReportDate As Date
Public OrderList As Collection      ' order records in file order
Public OrderIndex As Object         ' order id -> order record
Public Schedules As Object          ' order id -> schedule record (first crew file wins)
Public Contracts As Object          ' order id -> contract record
Public SpecRows As Collection       ' rows of the Specs table
Public SpecUnits As Object          ' order id -> specified units
Public ScheduleIssues As Collection
Public SpecIssues As Collection
Public ContractIssues As Collection
Public RunLog As Collection

Public Function OrderHeaders() As Variant
    OrderHeaders = Array("Order ID", "Order Date", "Order Month", "Client", "Product", "Planned Units", _
                         "Planned Month", "Manager", "Crew", "Planned Start", "Planned Finish", "Actual Start", _
                         "Actual Finish", "Installed Units", "Finish Month", "Status", "Contract Sent", _
                         "Contract Signed", "Spec Units")
End Function

Public Function SpecHeaders() As Variant
    SpecHeaders = Array("Spec No", "Order ID", "Product", "Spec Date", "Spec Month", "Units")
End Function

Public Function IssueHeaders() As Variant
    IssueHeaders = Array("Key", "Order ID", "Client", "Manager", "Crew", "Issue", "Details", "Days", "Note")
End Function

Public Function ManagerHeaders() As Variant
    ManagerHeaders = Array("Manager", "Orders", "Open Orders", "Delayed Start", "Schedule Issues", _
                           "Spec Issues", "Contract Issues")
End Function

Public Function LogHeaders() As Variant
    LogHeaders = Array("Step", "Seconds", "Rows")
End Function

Public Function ArchiveHeaders() As Variant
    ArchiveHeaders = Array("Archived At", "Sheet", "Key", "Note")
End Function

' Colours
Public Function ClrAccent() As Long
    ClrAccent = RGB(22, 82, 76)
End Function

Public Function ClrAccentLight() As Long
    ClrAccentLight = RGB(224, 239, 236)
End Function

Public Function ClrMuted() As Long
    ClrMuted = RGB(90, 96, 94)
End Function

Public Function ClrInput() As Long
    ClrInput = RGB(255, 250, 230)
End Function

Public Function ClrBad() As Long
    ClrBad = RGB(248, 215, 210)
End Function

Public Function ClrWarn() As Long
    ClrWarn = RGB(255, 235, 180)
End Function

Public Function ClrGood() As Long
    ClrGood = RGB(214, 234, 220)
End Function

Public Function ClrWarnText() As Long
    ClrWarnText = RGB(180, 60, 40)
End Function

Public Function ClrPlan() As Long
    ClrPlan = RGB(170, 190, 186)
End Function
