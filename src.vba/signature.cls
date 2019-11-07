Public r As String
Public S As String

Private Sub Class_Initialize()
    r = vbNullString
    S = vbNullString
End Sub

Public Function setProperties(p1 As String, p2 As String)
    r = p1
    S = p2
End Function

Public Function toDer() As Byte()
    Dim rEncoded() As Byte: rEncoded = EllipticCurve_Der.encodeInteger(r)
    Dim sEncoded() As Byte: sEncoded = EllipticCurve_Der.encodeInteger(S)
    toDer = EllipticCurve_Der.encodeSequence(rEncoded, sEncoded)
End Function

Public Function toBase64()
    Dim byteArr() As Byte: byteArr = toDer()
    toBase64 = Utils.encodeBase64(byteArr)
End Function
