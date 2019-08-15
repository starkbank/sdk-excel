Public r As String
Public s As String

Private Sub Class_Initialize()
    r = vbNullString
    s = vbNullString
End Sub

Public Function setProperties(p1 As String, p2 As String)
    r = p1
    s = p2
End Function

Public Function toDer() As Byte()
    Dim rEncoded() As Byte: rEncoded = EllipticCurve_Der.encodeInteger(r)
    Dim sEncoded() As Byte: sEncoded = EllipticCurve_Der.encodeInteger(s)
    toDer = EllipticCurve_Der.encodeSequence(rEncoded, sEncoded)
End Function

Public Function toBase64()
    Dim byteArr() As Byte: byteArr = toDer()
    toBase64 = Utils.encodeBase64(byteArr)
End Function
