Attribute VB_Name = "modUtils"
'==============================================================================
' Installation Tracker
' Helpers: dictionaries, settings, paths, text, tables, timing, app state.
'==============================================================================
Option Explicit

' Case-insensitive Scripting.Dictionary (late bound, no reference needed)
Public Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set NewDict = d
End Function

Public Function SettingValue(ByVal rangeName As String) As Variant
    SettingValue = ThisWorkbook.Names(rangeName).RefersToRange.Value
End Function

' Relative paths are resolved against the folder of this workbook
Public Function ResolvePath(ByVal p As String) As String
    p = Replace(Trim$(p), "/", "\")
    Do While Len(p) > 1 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    If p Like "[A-Za-z]:\*" Or Left$(p, 2) = "\\" Then
        ResolvePath = p
    Else
        ResolvePath = ThisWorkbook.Path & "\" & p
    End If
End Function

Public Function FileNameOf(ByVal fullPath As String) As String
    FileNameOf = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
End Function

Public Function MonthStart(ByVal d As Date) As Date
    MonthStart = DateSerial(Year(d), Month(d), 1)
End Function

Public Function MonthOrEmpty(ByVal v As Variant) As Variant
    If IsEmpty(v) Then MonthOrEmpty = Empty Else MonthOrEmpty = MonthStart(v)
End Function

Public Function DateText(ByVal v As Variant) As String
    If Not IsEmpty(v) Then DateText = Format$(v, "yyyy-mm-dd")
End Function

' Trims, removes non-breaking spaces and tabs, collapses double spaces
Public Function CleanText(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Or IsError(v) Then Exit Function
    s = Trim$(Replace(Replace(CStr(v), ChrW(160), " "), vbTab, " "))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

Public Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Len(CleanText(v)) = 0)
    End If
End Function

' Empty instead of "" so COUNTIFS(...,"") finds the blank cell
Public Function TextOrEmpty(ByVal s As String) As Variant
    If Len(s) = 0 Then TextOrEmpty = Empty Else TextOrEmpty = s
End Function

' Collection of 0-based row arrays -> 2D array for one fast write.
' Rows shorter than colCount leave the remaining cells empty.
Public Function CollectionTo2D(ByVal items As Collection, ByVal colCount As Long) As Variant
    Dim out() As Variant, rowArr As Variant, r As Long, c As Long
    ReDim out(1 To items.Count, 1 To colCount)
    For Each rowArr In items
        r = r + 1
        For c = 1 To colCount
            If c - 1 <= UBound(rowArr) Then out(r, c) = rowArr(c - 1)
        Next c
    Next rowArr
    CollectionTo2D = out
End Function

Public Function TableOf(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Set TableOf = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
End Function

' Replaces the rows of a table with the items of a collection in one write
Public Sub FillTable(ByVal lo As ListObject, ByVal items As Collection)
    Dim colCount As Long
    colCount = lo.ListColumns.Count
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    If items.Count = 0 Then Exit Sub
    lo.HeaderRowRange.Offset(1, 0).Resize(items.Count, colCount).Value = CollectionTo2D(items, colCount)
    lo.Resize lo.HeaderRowRange.Resize(items.Count + 1, colCount)
End Sub

' Values of one table column as a 2D array, also when the table has one row.
' The caller checks that the table has rows.
Public Function ColumnValues(ByVal lo As ListObject, ByVal columnName As String) As Variant
    Dim rng As Range, one(1 To 1, 1 To 1) As Variant
    Set rng = lo.ListColumns(columnName).DataBodyRange
    If rng.Rows.Count = 1 Then
        one(1, 1) = rng.Value
        ColumnValues = one
    Else
        ColumnValues = rng.Value
    End If
End Function

' Case-insensitive alphabetical order of a collection of strings
Public Function SortedText(ByVal items As Collection) As Variant
    Dim arr() As String, i As Long, j As Long, tmp As String
    If items.Count = 0 Then
        SortedText = Array()
        Exit Function
    End If
    ReDim arr(0 To items.Count - 1)
    For i = 1 To items.Count
        arr(i - 1) = items(i)
    Next i
    For i = 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= 0
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortedText = arr
End Function

' Whole used area of a sheet from A1, always as a 2D array (sheet row = array row)
Public Function SheetData(ByVal ws As Worksheet) As Variant
    Dim lastRow As Long, lastCol As Long, one(1 To 1, 1 To 1) As Variant
    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With
    If lastRow = 1 And lastCol = 1 Then
        one(1, 1) = ws.Cells(1, 1).Value
        SheetData = one
    Else
        SheetData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Value
    End If
End Function

' One line in the Run Log: step name, seconds since started, rows handled
Public Sub LogStep(ByVal stepName As String, ByVal started As Double, ByVal rowCount As Variant)
    RunLog.Add Array(stepName, Round(Timer - started, 2), rowCount)
End Sub

Public Sub SpeedOn()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
    End With
End Sub

Public Sub SpeedOff()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .StatusBar = False
    End With
End Sub

Public Sub SetStatus(ByVal text As String)
    Application.StatusBar = "Installation Tracker: " & text & "..."
End Sub
