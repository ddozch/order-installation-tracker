Attribute VB_Name = "modNotes"
'==============================================================================
' Installation Tracker
' Keeps what people type in the Note column of the issue sheets.
' A note follows its issue by Key across updates. When the issue is gone,
' the note moves to the Notes Archive sheet instead of disappearing.
'==============================================================================
Option Explicit

' Key -> Note of an issue table, read before the table is rebuilt
Public Function SaveNotes(ByVal lo As ListObject) As Object
    Dim notes As Object, keys As Variant, vals As Variant, r As Long
    Set notes = NewDict()
    If Not lo.DataBodyRange Is Nothing Then
        keys = ColumnValues(lo, "Key")
        vals = ColumnValues(lo, "Note")
        For r = 1 To UBound(keys, 1)
            If Not IsError(vals(r, 1)) And Not IsError(keys(r, 1)) Then
                If Len(Trim$(CStr(vals(r, 1)))) > 0 And Len(Trim$(CStr(keys(r, 1)))) > 0 Then
                    notes(CStr(keys(r, 1))) = vals(r, 1)
                End If
            End If
        Next r
    End If
    Set SaveNotes = notes
End Function

' Puts the notes back next to the same Key and archives the rest
Public Sub RestoreNotes(ByVal lo As ListObject, ByVal notes As Object, ByVal sheetName As String)
    Dim keys As Variant, out() As Variant, r As Long, k As String
    If Not lo.DataBodyRange Is Nothing And notes.Count > 0 Then
        keys = ColumnValues(lo, "Key")
        ReDim out(1 To UBound(keys, 1), 1 To 1)
        For r = 1 To UBound(keys, 1)
            k = CStr(keys(r, 1))
            If notes.Exists(k) Then
                out(r, 1) = notes(k)
                notes.Remove k
            End If
        Next r
        lo.ListColumns("Note").DataBodyRange.Value = out
    End If
    ArchiveNotes notes, sheetName
End Sub

Private Sub ArchiveNotes(ByVal notes As Object, ByVal sheetName As String)
    Dim lo As ListObject, k As Variant, newRow As ListRow
    If notes.Count = 0 Then Exit Sub
    Set lo = TableOf(SH_ARCHIVE, TBL_ARCHIVE)
    For Each k In notes.keys
        Set newRow = lo.ListRows.Add
        newRow.Range.Value = Array(Now, sheetName, CStr(k), notes(k))
    Next k
    lo.ListColumns(1).DataBodyRange.NumberFormat = "yyyy-mm-dd hh:mm"
End Sub
