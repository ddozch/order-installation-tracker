Attribute VB_Name = "modChecks"
'==============================================================================
' Installation Tracker
' Order status on the report date, automatic checks and the manager summary.
' The rules are written out in README.md, "Metric definitions".
'==============================================================================
Option Explicit

Public Function OrderStatus(ByVal rec As Variant, ByVal hasSchedule As Boolean, ByVal s As Variant) As String
    Dim units As Double
    If hasSchedule Then
        If Not IsEmpty(s(S_UNITS)) Then units = s(S_UNITS)
        If Not IsEmpty(s(S_AFINISH)) Then
            If s(S_AFINISH) <= ReportDate And units >= rec(O_UNITS) Then
                OrderStatus = ST_INSTALLED
                Exit Function
            End If
        End If
        If Not IsEmpty(s(S_ASTART)) Then
            OrderStatus = ST_IN_PROGRESS
            Exit Function
        End If
    End If

    If Not hasSchedule Then
        OrderStatus = ST_TO_SCHEDULE
    ElseIf IsEmpty(s(S_PSTART)) Then
        OrderStatus = ST_TO_SCHEDULE
    ElseIf s(S_PSTART) < ReportDate Then
        OrderStatus = ST_DELAYED
    Else
        OrderStatus = ST_SCHEDULED
    End If
End Function

' Rows of the Orders table; the three issue collections are filled on the way
Public Function BuildOrders() As Collection
    Dim out As Collection, rec As Variant, s As Variant, c As Variant, id As String
    Dim hasS As Boolean, hasC As Boolean, status As String, specTotal As Double
    Dim reportMonth As Date, afterNext As Date

    Set out = New Collection
    reportMonth = MonthStart(ReportDate)
    afterNext = DateSerial(Year(ReportDate), Month(ReportDate) + 1, 1)

    For Each rec In OrderList
        id = rec(O_ID)
        hasS = Schedules.Exists(id)
        If hasS Then s = Schedules(id) Else s = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
        hasC = Contracts.Exists(id)
        If hasC Then c = Contracts(id) Else c = Array(Empty, Empty, Empty)
        specTotal = 0
        If SpecUnits.Exists(id) Then specTotal = SpecUnits(id)
        status = OrderStatus(rec, hasS, s)

        CheckSchedule rec, hasS, s, status, reportMonth
        If status = ST_INSTALLED Then CheckSpecification rec, s, specTotal
        CheckContract rec, hasC, c, afterNext

        ' Same order as OrderHeaders()
        out.Add Array(id, rec(O_DATE), MonthOrEmpty(rec(O_DATE)), rec(O_CLIENT), rec(O_PRODUCT), rec(O_UNITS), _
                      rec(O_PMONTH), TextOrEmpty(rec(O_MANAGER)), s(S_CREW), s(S_PSTART), s(S_PFINISH), _
                      s(S_ASTART), s(S_AFINISH), s(S_UNITS), MonthOrEmpty(s(S_AFINISH)), status, _
                      c(C_SENT), c(C_SIGNED), IIf(specTotal > 0, specTotal, Empty))
    Next rec
    Set BuildOrders = out
End Function

Private Sub CheckSchedule(ByVal rec As Variant, ByVal hasS As Boolean, ByVal s As Variant, _
                          ByVal status As String, ByVal reportMonth As Date)
    Dim id As String, crew As String, units As Double

    id = rec(O_ID)
    If Not hasS Then
        If Not IsEmpty(rec(O_PMONTH)) Then
            If rec(O_PMONTH) <= reportMonth Then
                AddIssue ScheduleIssues, id & "|" & IS_NOT_SCHEDULED, id, "", IS_NOT_SCHEDULED, _
                         "Planned for " & Format$(rec(O_PMONTH), "yyyy-mm") & ", not in any crew schedule", Empty
            End If
        End If
        Exit Sub
    End If

    crew = s(S_CREW)
    If Not IsEmpty(s(S_UNITS)) Then units = s(S_UNITS)
    If IsEmpty(s(S_PSTART)) Or IsEmpty(s(S_PFINISH)) Then
        AddIssue ScheduleIssues, id & "|" & IS_NO_PLAN_DATES, id, crew, IS_NO_PLAN_DATES, _
                 "Planned start or finish is empty", Empty
    End If
    If (units > 0 And IsEmpty(s(S_ASTART))) Or (units >= rec(O_UNITS) And IsEmpty(s(S_AFINISH))) Then
        AddIssue ScheduleIssues, id & "|" & IS_NO_ACTUAL_DATES, id, crew, IS_NO_ACTUAL_DATES, _
                 units & " units installed, but actual start or finish is empty", Empty
    End If
    If Not IsEmpty(s(S_ASTART)) And Not IsEmpty(s(S_AFINISH)) Then
        If s(S_AFINISH) < s(S_ASTART) Then
            AddIssue ScheduleIssues, id & "|" & IS_FINISH_BEFORE_START, id, crew, IS_FINISH_BEFORE_START, _
                     "Start " & DateText(s(S_ASTART)) & ", finish " & DateText(s(S_AFINISH)), Empty
        End If
    End If
    If status = ST_DELAYED Then
        AddIssue ScheduleIssues, id & "|" & IS_START_DELAYED, id, crew, IS_START_DELAYED, _
                 "Planned start " & DateText(s(S_PSTART)), CLng(ReportDate - s(S_PSTART))
    End If
End Sub

' Only for installed orders: the specification closes the job
Private Sub CheckSpecification(ByVal rec As Variant, ByVal s As Variant, ByVal specTotal As Double)
    Dim id As String
    id = rec(O_ID)
    If specTotal = 0 Then
        If s(S_AFINISH) <= ReportDate - SPEC_DUE_DAYS Then
            AddIssue SpecIssues, id & "|" & IS_NO_SPEC, id, s(S_CREW), IS_NO_SPEC, _
                     "Installed " & DateText(s(S_AFINISH)), CLng(ReportDate - s(S_AFINISH))
        End If
    ElseIf specTotal < s(S_UNITS) Then
        AddIssue SpecIssues, id & "|" & IS_SPEC_INCOMPLETE, id, s(S_CREW), IS_SPEC_INCOMPLETE, _
                 "Specified " & specTotal & " of " & s(S_UNITS) & " installed units", Empty
    End If
End Sub

Private Sub CheckContract(ByVal rec As Variant, ByVal hasC As Boolean, ByVal c As Variant, ByVal afterNext As Date)
    Dim id As String, overdue As Long

    id = rec(O_ID)
    If Not hasC Then
        If Not IsEmpty(rec(O_PMONTH)) Then
            If rec(O_PMONTH) <= afterNext Then
                AddIssue ContractIssues, id & "|" & IS_NO_CONTRACT, id, "", IS_NO_CONTRACT, _
                         "Installation planned for " & Format$(rec(O_PMONTH), "yyyy-mm"), Empty
            End If
        End If
        Exit Sub
    End If
    If IsEmpty(c(C_SENT)) Then Exit Sub

    If Not IsEmpty(c(C_SIGNED)) Then
        overdue = CLng(c(C_SIGNED) - c(C_SENT)) - CONTRACT_SIGN_DAYS
        If overdue > 0 Then
            AddIssue ContractIssues, id & "|" & IS_SIGNED_LATE, id, "", IS_SIGNED_LATE, _
                     "Sent " & DateText(c(C_SENT)) & ", signed " & DateText(c(C_SIGNED)), overdue
        End If
    Else
        overdue = CLng(ReportDate - c(C_SENT)) - CONTRACT_SIGN_DAYS
        If overdue > 0 Then
            AddIssue ContractIssues, id & "|" & IS_NOT_SIGNED, id, "", IS_NOT_SIGNED, _
                     "Sent " & DateText(c(C_SENT)) & ", not signed yet", overdue
        End If
    End If
End Sub

' Adds an issue row; client and manager come from the order list
Public Sub AddIssue(ByVal target As Collection, ByVal key As String, ByVal orderId As String, ByVal crew As String, _
                    ByVal issue As String, ByVal details As String, ByVal days As Variant)
    Dim client As String, manager As String, rec As Variant
    If Len(orderId) > 0 Then
        If OrderIndex.Exists(orderId) Then
            rec = OrderIndex(orderId)
            client = rec(O_CLIENT)
            manager = rec(O_MANAGER)
        End If
    End If
    ' Same order as IssueHeaders(); the Note column is filled by modNotes
    target.Add Array(key, orderId, client, TextOrEmpty(manager), TextOrEmpty(crew), issue, details, days)
End Sub

' One row per manager: orders, open orders, delayed starts and issues by sheet
Public Function BuildManagers(ByVal orderRows As Collection) As Collection
    Dim d As Object, row As Variant, key As String, v As Variant, names As Collection
    Dim sorted As Variant, out As Collection, i As Long

    Set d = NewDict()
    For Each row In orderRows
        key = ManagerLabel(row(OR_MANAGER))
        If Not d.Exists(key) Then d.Add key, Array(0, 0, 0, 0, 0, 0)
        v = d(key)
        v(0) = v(0) + 1
        If row(OR_STATUS) <> ST_INSTALLED Then v(1) = v(1) + 1
        If row(OR_STATUS) = ST_DELAYED Then v(2) = v(2) + 1
        d(key) = v
    Next row
    CountIssues d, ScheduleIssues, 3
    CountIssues d, SpecIssues, 4
    CountIssues d, ContractIssues, 5

    Set names = New Collection
    For Each v In d.keys
        If v <> ManagerLabel(Empty) Then names.Add CStr(v)
    Next v
    sorted = SortedText(names)
    Set out = New Collection
    For i = 0 To UBound(sorted)
        v = d(sorted(i))
        out.Add Array(sorted(i), v(0), v(1), v(2), v(3), v(4), v(5))
    Next i
    If d.Exists(ManagerLabel(Empty)) Then
        v = d(ManagerLabel(Empty))
        out.Add Array(ManagerLabel(Empty), v(0), v(1), v(2), v(3), v(4), v(5))
    End If
    Set BuildManagers = out
End Function

Private Sub CountIssues(ByVal d As Object, ByVal issues As Collection, ByVal position As Long)
    Dim it As Variant, key As String, v As Variant
    For Each it In issues
        ' Issues of rows that are not real orders (unknown order, unreadable file) have no manager
        If OrderIndex.Exists(it(I_ORDER)) Then
            key = ManagerLabel(it(I_MANAGER))
            If Not d.Exists(key) Then d.Add key, Array(0, 0, 0, 0, 0, 0)
            v = d(key)
            v(position) = v(position) + 1
            d(key) = v
        End If
    Next it
End Sub

Private Function ManagerLabel(ByVal v As Variant) As String
    If IsEmpty(v) Then
        ManagerLabel = "(no manager)"
    ElseIf Len(CStr(v)) = 0 Then
        ManagerLabel = "(no manager)"
    Else
        ManagerLabel = CStr(v)
    End If
End Function
