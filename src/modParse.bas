Attribute VB_Name = "modParse"
'==============================================================================
' Installation Tracker
' Header detection and tolerant parsing of dates and numbers in source files.
' Text is never converted with CDbl or CDate, so results do not depend on the
' Windows language settings (comma or dot decimals, day or month first).
'==============================================================================
Option Explicit

' Registers column names (normalised: lower case letters and digits) for one field
Public Sub AddAliases(ByVal d As Object, ByVal fieldName As String, ByVal aliases As String)
    Dim a As Variant
    For Each a In Split(aliases, ",")
        d(CStr(a)) = fieldName
    Next a
End Sub

' "Planned From" -> "plannedfrom"
Public Function NormalizeHeader(ByVal v As Variant) As String
    Dim s As String, ch As String, i As Long, out As String
    If IsEmpty(v) Or IsError(v) Then Exit Function
    s = LCase$(CStr(v))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch Like "[a-z0-9]" Then out = out & ch
    Next i
    NormalizeHeader = out
End Function

' Finds the header row in the first 15 rows. On success fills colMap (field -> column).
Public Function FindHeader(data As Variant, ByVal aliases As Object, ByVal minMatches As Long, _
                           ByRef colMap As Object, ByRef headerRow As Long) As Boolean
    Dim r As Long, c As Long, lastScan As Long, key As String, found As Object
    lastScan = UBound(data, 1)
    If lastScan > 15 Then lastScan = 15
    For r = 1 To lastScan
        Set found = NewDict()
        For c = 1 To UBound(data, 2)
            key = NormalizeHeader(data(r, c))
            If Len(key) > 0 Then
                If aliases.Exists(key) Then
                    If Not found.Exists(aliases(key)) Then found(aliases(key)) = c
                End If
            End If
        Next c
        If found.Count >= minMatches Then
            Set colMap = found
            headerRow = r
            FindHeader = True
            Exit Function
        End If
    Next r
End Function

' Value of a field in row r, or Empty when the file has no such column
Public Function FieldValue(data As Variant, ByVal r As Long, ByVal colMap As Object, ByVal fieldName As String) As Variant
    Dim v As Variant
    If colMap.Exists(fieldName) Then
        v = data(r, colMap(fieldName))
        If IsError(v) Then v = Empty
    End If
    FieldValue = v
End Function

' Blank -> Empty; a date -> Date; anything else -> Empty with ok = False
Public Function ReadDate(ByVal v As Variant, ByRef ok As Boolean) As Variant
    Dim d As Date
    ok = True
    If IsBlankValue(v) Then
        ReadDate = Empty
    ElseIf TryParseDate(v, d) Then
        ReadDate = d
    Else
        ok = False
        ReadDate = Empty
    End If
End Function

' Blank -> Empty; a number -> Double; anything else -> Empty with ok = False
Public Function ReadNumber(ByVal v As Variant, ByRef ok As Boolean) As Variant
    Dim n As Double
    ok = True
    If IsBlankValue(v) Then
        ReadNumber = Empty
    ElseIf TryParseNumber(v, n) Then
        ReadNumber = n
    Else
        ok = False
        ReadNumber = Empty
    End If
End Function

' Real numbers, or text like "1 149,00", "1,149.00", "12"
Public Function TryParseNumber(ByVal v As Variant, ByRef result As Double) As Boolean
    Dim s As String, lastComma As Long, lastDot As Long
    Select Case VarType(v)
        Case vbDouble, vbInteger, vbLong, vbSingle, vbCurrency, vbDecimal
            result = CDbl(v)
            TryParseNumber = True
            Exit Function
        Case vbString
        Case Else
            Exit Function
    End Select

    s = Replace(Replace(CStr(v), ChrW(160), ""), " ", "")
    If Len(s) = 0 Then Exit Function
    ' The separator that comes last is the decimal one
    lastComma = InStrRev(s, ",")
    lastDot = InStrRev(s, ".")
    If lastComma > 0 And lastDot > 0 Then
        If lastComma > lastDot Then
            s = Replace(Replace(s, ".", ""), ",", ".")
        Else
            s = Replace(s, ",", "")
        End If
    ElseIf lastComma > 0 Then
        s = Replace(s, ",", ".")
    End If
    If Not IsPlainNumber(s) Then Exit Function
    result = Val(s)
    TryParseNumber = True
End Function

' Real dates, Excel serial numbers, or text as dd.mm.yyyy, dd/mm/yyyy or yyyy-mm-dd.
' Impossible dates such as 31.09 fail.
Public Function TryParseDate(ByVal v As Variant, ByRef result As Date) As Boolean
    Dim s As String, parts() As String, y As Long, m As Long, d As Long
    Select Case VarType(v)
        Case vbDate
            result = DateSerial(Year(v), Month(v), Day(v))
            TryParseDate = True
            Exit Function
        Case vbDouble, vbLong, vbInteger, vbSingle
            If v >= 1 And v < 2958466 Then
                result = DateSerial(Year(CDate(v)), Month(CDate(v)), Day(CDate(v)))
                TryParseDate = True
            End If
            Exit Function
        Case vbString
        Case Else
            Exit Function
    End Select

    s = Replace(Replace(CleanText(v), "/", "."), "-", ".")
    parts = Split(s, ".")
    If UBound(parts) <> 2 Then Exit Function
    If Not (IsAllDigits(parts(0)) And IsAllDigits(parts(1)) And IsAllDigits(parts(2))) Then Exit Function
    If Len(parts(0)) = 4 Then
        y = Val(parts(0)): m = Val(parts(1)): d = Val(parts(2))
    ElseIf Len(parts(2)) = 4 Then
        d = Val(parts(0)): m = Val(parts(1)): y = Val(parts(2))
    Else
        Exit Function
    End If
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    ' DateSerial rolls 31.09 over to 01.10, so check that the day survived
    result = DateSerial(y, m, d)
    If Day(result) <> d Or Month(result) <> m Then Exit Function
    TryParseDate = True
End Function

Private Function IsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long, digits As Long
    If Left$(s, 1) = "-" Then s = Mid$(s, 2)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf ch Like "#" Then
            digits = digits + 1
        Else
            Exit Function
        End If
    Next i
    IsPlainNumber = (digits > 0)
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not Mid$(s, i, 1) Like "#" Then Exit Function
    Next i
    IsAllDigits = True
End Function
