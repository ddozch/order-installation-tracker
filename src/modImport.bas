Attribute VB_Name = "modImport"
'==============================================================================
' Installation Tracker
' Reads orders, crew schedules, contracts and specifications from the data folder.
' Source files are opened read-only and never changed.
'==============================================================================
Option Explicit

Public Sub LoadOrders(ByVal path As String)
    Dim aliases As Object, data As Variant, colMap As Object, headerRow As Long, r As Long
    Dim id As String, ok As Boolean, pm As Variant, units As Variant, rec As Variant

    Set aliases = NewDict()
    AddAliases aliases, "OrderID", "orderid,orderno"
    AddAliases aliases, "OrderDate", "orderdate"
    AddAliases aliases, "Client", "client,customer"
    AddAliases aliases, "Product", "product"
    AddAliases aliases, "Units", "plannedunits,units"
    AddAliases aliases, "PlannedMonth", "plannedmonth"
    AddAliases aliases, "Manager", "accountmanager,manager"
    RequireSource path, aliases, 6, data, colMap, headerRow

    Set OrderList = New Collection
    Set OrderIndex = NewDict()
    For r = headerRow + 1 To UBound(data, 1)
        id = UCase$(CleanText(FieldValue(data, r, colMap, "OrderID")))
        If Len(id) > 0 And Not OrderIndex.Exists(id) Then
            pm = ReadDate(FieldValue(data, r, colMap, "PlannedMonth"), ok)
            If Not IsEmpty(pm) Then pm = MonthStart(pm)
            units = ReadNumber(FieldValue(data, r, colMap, "Units"), ok)
            If IsEmpty(units) Then units = 0
            rec = Array(id, ReadDate(FieldValue(data, r, colMap, "OrderDate"), ok), _
                        CleanText(FieldValue(data, r, colMap, "Client")), _
                        CleanText(FieldValue(data, r, colMap, "Product")), units, pm, _
                        CleanText(FieldValue(data, r, colMap, "Manager")))
            OrderList.Add rec
            OrderIndex.Add id, rec
        End If
    Next r
End Sub

' Reads every Crew_*.xlsx file. Files are taken in name order and the first file wins
' when an order appears twice; the extra row is reported, not imported.
Public Function CollectSchedules(ByVal folderPath As String) As Long
    Dim fso As Object, f As Object, names As Collection, fileNames As Variant, i As Long
    Dim aliases As Object, data As Variant, colMap As Object, headerRow As Long, r As Long
    Dim fileName As String, crew As String, id As String

    Set Schedules = NewDict()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Err.Raise vbObjectError + 3, "CollectSchedules", "Schedule folder not found: " & folderPath
    End If
    Set names = New Collection
    For Each f In fso.GetFolder(folderPath).Files
        If LCase$(f.Name) Like "crew_*.xls*" Then names.Add f.Name
    Next f
    fileNames = SortedText(names)

    ' Each crew names its columns a little differently
    Set aliases = NewDict()
    AddAliases aliases, "OrderID", "orderid,orderno,order,ordernumber"
    AddAliases aliases, "PlannedStart", "plannedstart,planstart,plannedfrom"
    AddAliases aliases, "PlannedFinish", "plannedfinish,planfinish,plannedend,planend,plannedto"
    AddAliases aliases, "ActualStart", "actualstart,start"
    AddAliases aliases, "ActualFinish", "actualfinish,finish,end"
    AddAliases aliases, "Units", "installedunits,installed,units"

    For i = 0 To UBound(fileNames)
        fileName = fileNames(i)
        crew = Mid$(fso.GetBaseName(fileName), Len("Crew_") + 1)
        Set colMap = Nothing
        If ReadSource(folderPath & "\" & fileName, aliases, 4, data, colMap, headerRow) Then
            For r = headerRow + 1 To UBound(data, 1)
                id = UCase$(CleanText(FieldValue(data, r, colMap, "OrderID")))
                If Len(id) > 0 Then ImportScheduleRow data, r, colMap, id, crew, fileName
            Next r
        Else
            AddIssue ScheduleIssues, fileName & "|" & IS_BAD_VALUE, "", crew, IS_BAD_VALUE, _
                     "No header row found in " & fileName, Empty
        End If
    Next i
    CollectSchedules = Schedules.Count
End Function

Private Sub ImportScheduleRow(data As Variant, ByVal r As Long, ByVal colMap As Object, ByVal id As String, _
                              ByVal crew As String, ByVal fileName As String)
    Dim fields As Variant, values(0 To 3) As Variant, units As Variant, ok As Boolean, i As Long, first As Variant

    If Not OrderIndex.Exists(id) Then
        AddIssue ScheduleIssues, id & "|" & IS_UNKNOWN_ORDER & "|" & crew, id, crew, IS_UNKNOWN_ORDER, _
                 "Row " & r & " in " & fileName & ": order is not in the order list", Empty
        Exit Sub
    End If
    If Schedules.Exists(id) Then
        first = Schedules(id)
        AddIssue ScheduleIssues, id & "|" & IS_DUPLICATE & "|" & crew, id, crew, IS_DUPLICATE, _
                 "Already in " & first(S_FILE) & "; row " & r & " in " & fileName & " ignored", Empty
        Exit Sub
    End If

    fields = Array("PlannedStart", "PlannedFinish", "ActualStart", "ActualFinish")
    For i = 0 To 3
        values(i) = ReadDate(FieldValue(data, r, colMap, fields(i)), ok)
        If Not ok Then
            AddIssue ScheduleIssues, id & "|" & IS_BAD_VALUE & "|" & fields(i), id, crew, IS_BAD_VALUE, _
                     fields(i) & " in row " & r & " of " & fileName & " is not a date", Empty
        End If
    Next i
    units = ReadNumber(FieldValue(data, r, colMap, "Units"), ok)
    If Not ok Then
        AddIssue ScheduleIssues, id & "|" & IS_BAD_VALUE & "|Units", id, crew, IS_BAD_VALUE, _
                 "Units in row " & r & " of " & fileName & " is not a number", Empty
    End If
    Schedules.Add id, Array(crew, values(0), values(1), values(2), values(3), units, fileName, r)
End Sub

Public Sub LoadContracts(ByVal path As String)
    Dim aliases As Object, data As Variant, colMap As Object, headerRow As Long, r As Long
    Dim id As String, ok As Boolean

    Set aliases = NewDict()
    AddAliases aliases, "OrderID", "orderid,orderno"
    AddAliases aliases, "Type", "contracttype,type"
    AddAliases aliases, "Sent", "senttoclient,sent,sentdate"
    AddAliases aliases, "Signed", "signed,signeddate,signedon"
    RequireSource path, aliases, 3, data, colMap, headerRow

    Set Contracts = NewDict()
    For r = headerRow + 1 To UBound(data, 1)
        id = UCase$(CleanText(FieldValue(data, r, colMap, "OrderID")))
        If Len(id) > 0 And Not Contracts.Exists(id) Then
            Contracts.Add id, Array(CleanText(FieldValue(data, r, colMap, "Type")), _
                                    ReadDate(FieldValue(data, r, colMap, "Sent"), ok), _
                                    ReadDate(FieldValue(data, r, colMap, "Signed"), ok))
        End If
    Next r
End Sub

Public Sub LoadSpecs(ByVal path As String)
    Dim aliases As Object, data As Variant, colMap As Object, headerRow As Long, r As Long
    Dim id As String, specNo As String, ok As Boolean, specDate As Variant, units As Variant, rec As Variant

    Set aliases = NewDict()
    AddAliases aliases, "SpecNo", "specno,specnumber,specificationno"
    AddAliases aliases, "OrderID", "orderid,orderno"
    AddAliases aliases, "SpecDate", "specdate,date"
    AddAliases aliases, "Units", "units,specunits"
    RequireSource path, aliases, 4, data, colMap, headerRow

    Set SpecRows = New Collection
    Set SpecUnits = NewDict()
    For r = headerRow + 1 To UBound(data, 1)
        id = UCase$(CleanText(FieldValue(data, r, colMap, "OrderID")))
        specNo = CleanText(FieldValue(data, r, colMap, "SpecNo"))
        If Len(id) > 0 Then
            If OrderIndex.Exists(id) Then
                rec = OrderIndex(id)
                specDate = ReadDate(FieldValue(data, r, colMap, "SpecDate"), ok)
                units = ReadNumber(FieldValue(data, r, colMap, "Units"), ok)
                If IsEmpty(units) Then units = 0
                SpecRows.Add Array(specNo, id, rec(O_PRODUCT), specDate, MonthOrEmpty(specDate), units)
                SpecUnits(id) = SpecUnits(id) + units
            Else
                AddIssue SpecIssues, id & "|" & IS_UNKNOWN_ORDER & "|" & specNo, id, "", IS_UNKNOWN_ORDER, _
                         "Specification " & specNo & " points to an order that is not in the order list", Empty
            End If
        End If
    Next r
End Sub

' Orders, contracts and specifications are required: without them the update stops
Private Sub RequireSource(ByVal path As String, ByVal aliases As Object, ByVal minMatches As Long, _
                          ByRef data As Variant, ByRef colMap As Object, ByRef headerRow As Long)
    If Not CreateObject("Scripting.FileSystemObject").FileExists(path) Then
        Err.Raise vbObjectError + 1, "RequireSource", "File not found: " & path
    End If
    If Not ReadSource(path, aliases, minMatches, data, colMap, headerRow) Then
        Err.Raise vbObjectError + 2, "RequireSource", "No header row found in " & FileNameOf(path)
    End If
End Sub

' Opens a source read-only, takes the first sheet with a recognisable header row
Private Function ReadSource(ByVal path As String, ByVal aliases As Object, ByVal minMatches As Long, _
                            ByRef data As Variant, ByRef colMap As Object, ByRef headerRow As Long) As Boolean
    Dim wb As Workbook, ws As Worksheet, errNo As Long, errText As String
    Set wb = Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    On Error GoTo Failed
    For Each ws In wb.Worksheets
        data = SheetData(ws)
        If FindHeader(data, aliases, minMatches, colMap, headerRow) Then Exit For
    Next ws
    wb.Close SaveChanges:=False
    ReadSource = Not colMap Is Nothing
    Exit Function

Failed:
    errNo = Err.Number
    errText = Err.Description
    wb.Close SaveChanges:=False
    Err.Raise errNo, "ReadSource", FileNameOf(path) & ": " & errText
End Function
