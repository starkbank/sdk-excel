Public Sub HelpDigSign()
    With ViewHelpForm
        .MultiPage1.Value = 2
        .Show
    End With
End Sub

Public Function DebugModeOn()
    DebugModeOn = (Sheets("InputLog").Cells(4, 2) = True)
End Function

Public Function CurrentDebugRow()
    CurrentDebugRow = Sheets("InputLog").Range("A9").CurrentRegion.Rows.Count + 9
End Function

Public Sub DebugPrint(header As String, content As String)
    Dim row As Integer
    Dim entry As String
    entry = CStr(content)
    If InStr(entry, """password""") Then
        entry = Left(entry, InStr(1, entry, """password""") - 1) + """password"":""***"",""platform"":""web""}"
    End If
    If InStr(entry, """accessToken""") Then
        entry = Left(entry, InStr(1, entry, """accessToken""") - 1) + """accessToken"":""***""}"
    End If
    row = CurrentDebugRow()
    Sheets("InputLog").Cells(row, 1) = Now
    Sheets("InputLog").Cells(row, 2) = header
    Sheets("InputLog").Cells(row, 3) = entry
End Sub