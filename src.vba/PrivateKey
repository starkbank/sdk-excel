Public curve As CurveFp
Public secret As String

Private Sub Class_Initialize()
    Set curve = New CurveFp
    secret = vbNullString
End Sub

Public Function fromPem(str As String)
    Dim index As Integer: index = InStr(str, "-----BEGIN EC PRIVATE KEY-----")
    Dim privkeyPem As String: privkeyPem = Mid(str, index, Len(str) - index + 1)
    
    Dim strByte() As Byte: strByte = EllipticCurve_Der.fromPem(privkeyPem)
    
    Call fromDer(strByte)
    
End Function

Public Function fromDer(strByte() As Byte)
    Dim strs() As String: strs = Split(EllipticCurve_Der.removeSequence(strByte), ",")
    Dim S As String: S = strs(0)
    Dim emptyStr As String: emptyStr = strs(1)
    
    If Len(emptyStr) <> 0 Then
        Err.Raise number:=vbObjectError + 513, description:="trailing junk after DER privkey: " & emptyStr
    End If
    
    strs = Split(EllipticCurve_Der.removeInteger(S), ",")
    Dim one As String: one = strs(0)
    S = strs(1)

    If one <> "1" Then
        Err.Raise number:=vbObjectError + 513, description:="expected '1' at start of DER privkey, got " & one
    End If
    
    strs = Split(EllipticCurve_Der.removeOctetString(S), ",")
    Dim privkeyStr As String: privkeyStr = strs(0)
    S = strs(1)
    
    strs = Split(EllipticCurve_Der.removeConstructed(S), ",")
    Dim tag As String: tag = strs(0)
    Dim curveOidStr As String: curveOidStr = strs(1)
    S = strs(2)
    
    If tag <> "0" Then
        Err.Raise number:=vbObjectError + 513, description:="expected tag 0 in DER privkey, got " & tag
    End If
    
    strs = Split(EllipticCurve_Der.removeObject(curveOidStr), ",")
    Dim oidCurve As String: oidCurve = strs(0)
    emptyStr = strs(1)
    
    If Len(emptyStr) <> 0 Then
        Err.Raise number:=vbObjectError + 513, description:="trailing junk after DER privkey curve_oid: " & emptyStr
    End If
    
    Dim c As CurveFp: Set c = New CurveFp
    Dim knownCurve As Boolean: knownCurve = c.initializeByOid(oidCurve)
    If knownCurve = False Then
        Err.Raise number:=vbObjectError + 513, description:="Unknown curve with oid " & oidCurve
    End If
    
    Dim lenPrivKey As Integer: lenPrivKey = (Len(privkeyStr) + 1) / 3
    If lenPrivKey < c.length() Then
        For i = 1 To c.length() - lenPrivKey
            privkeyStr = "00 " & privkeyStr
        Next
    End If
    
    Call fromString(privkeyStr, c)
End Function

Public Function fromString(strHexa As String, c As CurveFp)
    secret = BigIntMath.BigIntFromString(strHexa, 16)
    Set curve = c
End Function