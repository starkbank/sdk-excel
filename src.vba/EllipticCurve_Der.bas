Public Function fromPem(pem As String) As String
    Dim pieces() As String
    Dim d As String: d = ""
    pieces() = Split(pem, vbLf)
    For Each piece In pieces
        If piece <> vbNullString And Left(piece, 5) <> "-----" Then
            d = d & piece
        End If
    Next
    
    fromPem = Utils.decodeBase64(d)
End Function

Public Function removeObject(strHexa As String) As String
    If Left(strHexa, 2) <> "06" Then
         Err.Raise number:=vbObjectError + 513, description:="wanted object (0x06), got 0x" & Left(strHexa, 2)
    End If
    
    Dim str1 As String: str1 = Utils.getSubByteArray(strHexa, 1, -1)
    Dim lenghts() As String: lenghts = Split(readLength(str1), ",")
    Dim endseq As String: endseq = BigIntMath.Add("1", BigIntMath.Add(lenghts(0), lenghts(1)))
    
    Dim body As String: body = Utils.getSubByteArray(strHexa, 1 + CInt(lenghts(1)), CInt(endseq) - 1)
    Dim rest As String: rest = Utils.getSubByteArray(strHexa, CInt(endseq), -1)
    
    Dim n As String, ll As Integer, numbers() As String, numbersRead() As String, i As Integer: i = 0
    While body <> vbNullString
        numbersRead = Split(readNumber(body), ",")
        
        ReDim Preserve numbers(i)
        numbers(i) = numbersRead(0)
        i = i + 1
        body = Utils.getSubByteArray(body, CInt(numbersRead(1)), -1)
    Wend
    
    Dim first As String, second As String
    first = BigIntMath.Divide(numbers(0), "40")
    second = BigIntMath.Subtract(numbers(0), BigIntMath.multiply("40", first))
    
    ReDim Preserve numbers(i)
    For i = 0 To UBound(numbers) - 1
        numbers(UBound(numbers) - i) = numbers(UBound(numbers) - i - 1)
    Next
    numbers(0) = first
    numbers(1) = second
    
    Dim respNumbers As String: respNumbers = vbNullString
    For Each num In numbers
        respNumbers = respNumbers & num & ";"
    Next
    respNumbers = Left(respNumbers, Len(respNumbers) - 1)
    
    removeObject = respNumbers & "," & rest
End Function

Public Function removeConstructed(strHexa As String) As String
    Dim num As Integer: num = BigIntMath.BigIntFromString(Left(strHexa, 2), 16)
    
    If (num And 224) <> 160 Then
        Err.Raise number:=vbObjectError + 513, description:="wanted constructed tag (0xa0-0xbf), got 0x" & Left(strHexa, 2)
    End If
    
    Dim tag As String: tag = CStr(num And 31)
    Dim str1 As String: str1 = Utils.getSubByteArray(strHexa, 1, -1)
    Dim lenghts() As String: lenghts = Split(readLength(str1), ",")
    Dim endseq As String: endseq = BigIntMath.Add("1", BigIntMath.Add(lenghts(0), lenghts(1)))
    
    Dim body As String: body = Utils.getSubByteArray(strHexa, 1 + CInt(lenghts(1)), CInt(endseq) - 1)
    Dim rest As String: rest = Utils.getSubByteArray(strHexa, CInt(endseq), -1)
    
    removeConstructed = tag & "," & body & "," & rest
End Function

Public Function removeOctetString(strHexa As String) As String
    If Left(strHexa, 2) <> "04" Then
         Err.Raise number:=vbObjectError + 513, description:="wanted octetstring (0x04), got 0x" & Left(strHexa, 2)
    End If
    
    Dim str1 As String: str1 = Utils.getSubByteArray(strHexa, 1, -1)
    Dim lenghts() As String: lenghts = Split(readLength(str1), ",")
    Dim endseq As String: endseq = BigIntMath.Add("1", BigIntMath.Add(lenghts(0), lenghts(1)))
    
    Dim body As String: body = Utils.getSubByteArray(strHexa, 1 + CInt(lenghts(1)), CInt(endseq) - 1)
    Dim rest As String: rest = Utils.getSubByteArray(strHexa, CInt(endseq), -1)
    
    removeOctetString = body & "," & rest
End Function

Public Function removeSequence(strByte() As Byte) As String
    Dim strHexa As String: strHexa = BytesToHex(strByte)
    
    If Left(strHexa, 2) <> "30" Then
         Err.Raise number:=vbObjectError + 513, description:="wanted sequence (0x30), got 0x" & Left(strHexa, 2)
    End If
    
    Dim str1 As String: str1 = Utils.getSubByteArray(strHexa, 1, -1)
    Dim lenghts() As String: lenghts = Split(readLength(str1), ",")
    Dim endseq As String: endseq = BigIntMath.Add("1", BigIntMath.Add(lenghts(0), lenghts(1)))
    
    Dim resp1 As String: resp1 = Utils.getSubByteArray(strHexa, 1 + CInt(lenghts(1)), CInt(endseq) - 1)
    Dim resp2 As String: resp2 = Utils.getSubByteArray(strHexa, CInt(endseq), -1)
    
    removeSequence = resp1 & "," & resp2
End Function

Public Function removeInteger(strHexa As String) As String
    If Left(strHexa, 2) <> "02" Then
         Err.Raise number:=vbObjectError + 513, description:="wanted integer (0x02), got 0x" & Left(strHexa, 2)
    End If
    
    Dim str1 As String: str1 = Utils.getSubByteArray(strHexa, 1, -1)
    Dim lenghts() As String: lenghts = Split(readLength(str1), ",")
    Dim endseq As String: endseq = BigIntMath.Add("1", BigIntMath.Add(lenghts(0), lenghts(1)))
    
    Dim numberbytes As String: numberbytes = Utils.getSubByteArray(strHexa, 1 + CInt(lenghts(1)), CInt(endseq) - 1)
    Dim rest As String: rest = Utils.getSubByteArray(strHexa, CInt(endseq), -1)
    
    Dim nbytes As String: nbytes = BigIntMath.BigIntFromString(Left(strHexa, 2), 16)
    Debug.Assert nbytes < 128
    
    removeInteger = BigIntMath.BigIntFromString(numberbytes, 16) & "," & rest
End Function

Public Function readNumber(strHexa As String) As String
    Dim number As String: num = "0"
    Dim llen As Integer: llen = 0
    Dim d As Integer

    Do While True
        If llen > Len(strHexa) Then
            Err.Raise number:=vbObjectError + 513, description:="ran out of length bytes"
        End If
        
        number = Utils.bitwiseLeftShift(number, 7)
        d = BigIntMath.BigIntFromString(Utils.getSubByteArray(strHexa, llen, llen), 16)
        number = BigIntMath.Add(number, CStr(d And 127))
        llen = llen + 1
        
        If (d And 128) = 0 Then
            Exit Do
        End If
    Loop
    
    readNumber = number & "," & CStr(llen)
End Function

Public Function readLength(strHexa As String) As String
    Dim respArray As String
    Dim num As Integer: num = BigIntMath.BigIntFromString(Left(strHexa, 2), 16)

    If (num And 128) = 0 Then
        respArray = CStr(num And 127) & "," & CStr(1)
    Else
        Dim llen As Integer: llen = (num And 127)
        
        If llen > Len(strHexa) - 1 Then
            Err.Raise number:=vbObjectError + 513, description:="ran out of length bytes"
        End If
        
        respArray = BigIntMath.BigIntFromString(Utils.getSubByteArray(strHexa, 1, llen), 16) & "," & CStr(1 + llen)
    End If
    
    readLength = respArray
End Function

Public Function encodeSequence(r() As Byte, s() As Byte) As Byte()
    rLen = UBound(r, 1) - LBound(r, 1) + 1
    sLen = UBound(s, 1) - LBound(s, 1) + 1
    Dim totalLen As Integer: totalLen = rLen + sLen
    
    Dim encodedLength() As Byte: encodedLength = encodeLength(totalLen)
    encodedLengthInHexa = BytesToHex(encodedLength)
    
    encodeSequence = HexToBytes("30" & encodedLengthInHexa & BytesToHex(r) & BytesToHex(s))

End Function

Public Function encodeLength(length As Integer) As Byte()
    If length < 128 Then
        encodeLength = HexToBytes(Hex(length))
    Else
        H = Hex(length)
        If Len(H) Mod 2 = 1 Then
            H = "0" & H
        End If
        
        Dim hByte() As Byte: hByte = HexToBytes(H)
        llen = UBound(hByte, 1) - LBound(hByte, 1) + 1
        
        Dim union As Integer: union = 128 Or llen
        
        encodeLength = HexToBytes(Hex(union) & H)
        
    End If
End Function

Public Function encodeInteger(r As String) As Byte()
    Debug.Assert BigIntMath.Compare(r, "0") <> -1
    
    H = BigIntMath.BigIntToString(r, 16)
    If Len(H) Mod 2 = 1 Then
        H = "0" & H
    End If
    
    Dim hByte() As Byte: hByte = HexToBytes(H)
    llen = UBound(hByte, 1) - LBound(hByte, 1) + 1
    
    firstByteInHexa = hByte(0)
    
    If firstByteInHexa <= 127 Then
        llenInHexa = Hex(llen)
        If Len(llenInHexa) Mod 2 = 1 Then
            llenInHexa = "0" & llenInHexa
        End If
        
        encodeInteger = HexToBytes("02" & llenInHexa & H)
    Else
        llenInHexa = Hex(llen + 1)
        If Len(llenInHexa) Mod 2 = 1 Then
            llenInHexa = "0" & llenInHexa
        End If
        
        encodeInteger = HexToBytes("02" & llenInHexa & "00" & H)
    End If
End Function

Public Function HexToBytes(ByVal HexString As String) As Byte()
    'Quick and dirty hex String to Byte array.  Accepts:
    '
    '   "HH HH HH"
    '   "HHHHHH"
    '   "H HH H"
    '   "HH,HH,     HH" and so on.

    Dim bytes() As Byte
    Dim HexPos As Integer
    Dim HexDigit As Integer
    Dim BytePos As Integer
    Dim Digits As Integer

    ReDim bytes(Len(HexString) \ 2)  'Initial estimate.
    For HexPos = 1 To Len(HexString)
        HexDigit = InStr("0123456789ABCDEF", _
                         UCase$(Mid$(HexString, HexPos, 1))) - 1
        If HexDigit >= 0 Then
            If BytePos > UBound(bytes) Then
                'Add some room, we'll add room for 4 more to decrease
                'how often we end up doing this expensive step:
                ReDim Preserve bytes(UBound(bytes) + 4)
            End If
            bytes(BytePos) = bytes(BytePos) * &H10 + HexDigit
            Digits = Digits + 1
        End If
        If Digits = 2 Or HexDigit < 0 Then
            If Digits > 0 Then BytePos = BytePos + 1
            Digits = 0
        End If
    Next
    If Digits = 0 Then BytePos = BytePos - 1
    If BytePos < 0 Then
        bytes = "" 'Empty.
    Else
        ReDim Preserve bytes(BytePos)
    End If
    HexToBytes = bytes
End Function

Public Function BytesToHex(ByRef bytes() As Byte) As String
    'Quick and dirty Byte array to hex String, format:
    '
    '   "HH HH HH"

    Dim LB As Long
    Dim ByteCount As Long
    Dim BytePos As Integer

    LB = LBound(bytes)
    ByteCount = UBound(bytes) - LB + 1
    If ByteCount < 1 Then Exit Function
    BytesToHex = Space$(3 * (ByteCount - 1) + 2)
    For BytePos = LB To UBound(bytes)
        Mid$(BytesToHex, 3 * (BytePos - LB) + 1, 2) = _
            Right$("0" & Hex$(bytes(BytePos)), 2)
    Next
End Function