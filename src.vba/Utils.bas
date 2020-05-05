Public Function Min(a As Variant, b As Variant) As Variant
    If a >= b Then
        Min = b
    Else
        Min = a
    End If
End Function

Public Function Max(a As Variant, b As Variant) As Variant
    If a < b Then
        Max = b
    Else
        Max = a
    End If
End Function

Public Function CollectionToString(c As Collection, Optional Delimiter As String) As String
    Dim elString As String: elString = ""
    If c.count <> 0 Then
        For Each el In c
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

Public Function formatCurrencyInUserForm(buffer As String)
    Dim clearBuffer As String
    If buffer = "" Then
        buffer = "0"
    End If
    clearBuffer = clearNonNumeric(buffer)
    If Len(clearBuffer) > 9 Then
        clearBuffer = Left(clearBuffer, 9)
    End If
    If buffer = "" Then
        buffer = "0"
    End If
    If Len(clearBuffer) > 9 Then
        clearBuffer = Left(clearBuffer, 9)
    End If
    buffer = Format(IntegerFrom(clearBuffer) / 100, "R$ #,##0.00")
    formatCurrencyInUserForm = buffer
End Function

Public Function clearNonNumeric(text)
    Dim Reg As RegExp
    Set Reg = New RegExp
    Reg.Global = True
    Reg.Pattern = "\D"
    clearNonNumeric = Reg.Replace(text, "")
    Set Reg = Nothing
End Function

Function correctErrorLine(errorMessage As String, offset As Integer) As String
    Dim lineNumber As Integer
    Dim message As String
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "(\w+) +(\d+):(.+)"
        .Global = True
        For Each M In .Execute(errorMessage)
            lineNumber = CInt(M.submatches(1))
            message = M.submatches(2)
        Next
    End With
    
    correctErrorLine = "Linha " & CStr(lineNumber + offset) & ": " & message
    
End Function

Public Function IntegerFrom(Value As String) As Long
    Dim temp As String
    temp = Value
    On Error GoTo eh:
    With CreateObject("VBScript.RegExp")
        .Pattern = "[^\d]+"
        .Global = True
        temp = .Replace(temp, "")
    End With
    IntegerFrom = CLng(temp)
    Exit Function
eh:
    IntegerFrom = 0
End Function

Public Function DateToSendingFormat(dateInput As String) As String
    Dim dateToSend As String: dateToSend = Mid(dateInput, 7, 4) + "-" + Mid(dateInput, 4, 2) + "-" + Mid(dateInput, 1, 2)
    DateToSendingFormat = dateToSend
End Function

Public Function SingleFrom(Value As String) As Single
    Dim temp As String
    temp = Value
    With CreateObject("VBScript.RegExp")
        .Pattern = "%"
        .Global = True
        temp = .Replace(temp, "")
    End With
    SingleFrom = CSng(temp)
End Function

Public Function MoneyStringFrom(Value As Double) As String
    MoneyStringFrom = Format(Value / 100, "Currency")
End Function

Public Function SHA256function(sMessage As String)

    Dim clsX As CSHA256
    Set clsX = New CSHA256

    SHA256function = clsX.SHA256(sMessage)

    Set clsX = Nothing

End Function

Public Function encodeBase64(ByRef arrData() As Byte) As String
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument60
    
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    encodeBase64 = Replace(objNode.text, vbLf, "")
 
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Public Function decodeBase64(ByVal strData As String) As Byte()
    Dim objXML As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    
    Set objXML = New MSXML2.DOMDocument60
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = strData
    decodeBase64 = objNode.nodeTypedValue
    
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Public Function getSubByteArray(strHexa As String, indexI As Integer, indexF As Integer) As String
    ' Index starts with zero
    If indexF = -1 Or indexF > (Len(strHexa) + 1) / 3 - 1 Then
        If indexI > (Len(strHexa) + 1) / 3 - 1 Then
            getSubByteArray = vbNullString
        Else
            getSubByteArray = Right(strHexa, Len(strHexa) - (3 * indexI))
        End If
    Else
        getSubByteArray = Mid(strHexa, 3 * indexI + 1, 3 * (indexF - indexI) + 2)
    End If
End Function

Public Function bitwiseLeftShift(ByVal Value As String, ByVal Shift As Integer) As String
    bitwiseLeftShift = Value
    If Shift > 0 Then
        Dim exp As String: exp = "1"
        For i = 1 To Shift
            exp = BigIntMath.multiply(exp, "2")
        Next i
        bitwiseLeftShift = BigIntMath.multiply(exp, Value)
    End If
End Function

Public Function bitwiseRightShift(ByVal Value As String, ByVal Shift As Integer) As String
    bitwiseRightShift = Value
    If Shift > 0 Then
        Dim exp As String: exp = "1"
        For i = 1 To Shift
            exp = BigIntMath.multiply(exp, "2")
        Next i
        bitwiseRightShift = BigIntMath.Divide(Value, exp)
    End If
End Function

Public Function randrange(ByVal lowerbound As String, ByVal upperbound As String) As String
    Dim randNumb As String, part1 As String, part2 As String, part3 As String
    part1 = BigIntMath.Add(BigIntMath.Subtract(upperbound, lowerbound), "1")
    part2 = BigIntMath.multiply(part1, CStr(Int(Rnd * 10000000)))
    part3 = BigIntMath.Divide(part2, "10000000")
    randrange = BigIntMath.Add(part3, lowerbound)
End Function

Public Function correctTransferTags(tags As Collection)
    Dim pathBool As Boolean
    Dim allMatches As Object
    Dim index As Integer
    Dim tag As Variant
    
    index = 0
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "team/\d+/list/\d+"
        .Global = True
        .IgnoreCase = True
        
        For Each tag In tags
            index = index + 1
            pathBool = False
            Set allMatches = .Execute(tag)
            If allMatches.count <> 0 Then
                pathBool = True
            End If
            If pathBool = True Then
                tags.Remove index
                index = index - 1
            End If
        Next
    End With
    
    Set correctTransferTags = tags
End Function

Public Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
Dim element As Variant
On Error GoTo IsInArrayError:
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

Function GetFolder() As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Selecione uma pasta"
        .AllowMultiSelect = False
        .InitialFileName = Application.DefaultFilePath
        If .Show <> -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolder = sItem
    Set fldr = Nothing
End Function

Public Function ShellRun(sCmd As String) As String

    'Run a shell command, returning the output as a string

    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")

    'run command
    Dim oExec As Object
    Dim oOutput As Object
    Set oExec = oShell.Exec(sCmd)
    Set oOutput = oExec.StdOut

    'handle the results as they are written to and read from the StdOut object
    Dim s As String
    Dim sLine As String
    While Not oOutput.AtEndOfStream
        sLine = oOutput.ReadLine
        If sLine <> "" Then s = s & sLine & vbCrLf
    Wend

    ShellRun = s

End Function