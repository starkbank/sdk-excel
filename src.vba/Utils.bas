Function CollectionToString(C As Collection, Optional Delimiter As String) As String
    Dim elString As String: elString = ""
    If C.Count <> 0 Then
        For Each el In C
            elString = elString + el + Delimiter
        Next
        elString = Left(elString, Len(elString) - 1)
    End If
    
    CollectionToString = elString
End Function

Public Function ISODATEZ(iso As String) As Date
    Dim yearPart As Integer: yearPart = CInt(Mid(iso, 1, 4))
    Dim monPart As Integer: monPart = CInt(Mid(iso, 6, 2))
    Dim dayPart As Integer: dayPart = CInt(Mid(iso, 9, 2))
    Dim hourPart As Integer: hourPart = CInt(Mid(iso, 12, 2))
    Dim minPart As Integer: minPart = CInt(Mid(iso, 15, 2))
    Dim secPart As Integer: secPart = CInt(Mid(iso, 18, 2))
    Dim tz As String: tz = Mid(iso, 28)
    
    Dim dt As Date: dt = DateSerial(yearPart, monPart, dayPart) + TimeSerial(hourPart, minPart, secPart)
    
    ' Add the timezone
    If tz <> "" And Left(tz, 1) <> "Z" Then
        Dim colonPos As Integer: colonPos = InStr(tz, ":")
        If colonPos = 0 Then colonPos = Len(tz) + 1

        Dim minutes As Integer: minutes = CInt(Mid(tz, 2, colonPos - 2)) * 60 + CInt(Mid(tz, colonPos + 1))
        If Left(tz, 1) = "+" Then minutes = -minutes
        dt = DateAdd("n", minutes, dt)
    End If

    ' Return value is the ISO8601 date in the local time zone
    dt = TimeZoneConverter.UtcToBrt(dt)
    
    ISODATEZ = dt
End Function

Public Function applyStandardLayout(col As String)
    ActiveSheet.Range("A1:" + col + "8").Interior.Color = RGB(255, 255, 255)
    ActiveSheet.Range("A9:" + col + "9").Interior.Color = RGB(99, 114, 130)
    ActiveSheet.Range("A9:" + col + "9").Font.Color = RGB(255, 255, 255)
End Function

Public Function formatDateInUserForm(dateString As String)
    Dim chars() As Byte
    chars = StrConv(dateString, vbFromUnicode)

    Dim buffer As String
    Dim i As Integer
    For i = LBound(chars) To UBound(chars)
        If Len(buffer) = 2 Then buffer = buffer & "/"   'auto-insert the dash
        If Len(buffer) = 5 Then buffer = buffer & "/"   'auto-insert the dash
        If Len(buffer) = 10 Then Exit For               'limit to 10 chars
        If chars(i) >= 48 And chars(i) <= 57 Then       'ignore anything but numbers.
            buffer = buffer & Chr$(chars(i))
        End If
    Next i
    
    formatDateInUserForm = buffer
        
End Function

Function correctErrorLine(errorMessage As String, offset As Integer) As String
    Dim lineNumber As Integer
    Dim message As String
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "(\w+) +(\d+):(.+)"
        .Global = True
        For Each m In .Execute(errorMessage)
            lineNumber = CInt(m.SubMatches(1))
            message = m.SubMatches(2)
        Next
    End With
    
    correctErrorLine = "Linha " & CStr(lineNumber + offset) & ": " & message
    
End Function

Public Function IntegerFrom(value As String) As Long
    Dim temp As String
    temp = value
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^\d]+"
        .Global = True
        temp = .Replace(temp, "")
    End With
    IntegerFrom = CLng(temp)
End Function

Public Function DateToSendingFormat(dateInput As String) As String
    Dim dateToSend As String: dateToSend = Mid(dateInput, 7, 4) + "-" + Mid(dateInput, 4, 2) + "-" + Mid(dateInput, 1, 2)
    DateToSendingFormat = dateToSend
End Function

Public Function SingleFrom(value As String) As Single
    Dim temp As String
    temp = value
    With CreateObject("VBScript.RegExp")
        .Pattern = "%"
        .Global = True
        temp = .Replace(temp, "")
    End With
    SingleFrom = CSng(temp)
End Function