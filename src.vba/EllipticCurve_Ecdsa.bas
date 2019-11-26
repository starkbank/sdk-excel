Public Function sign(message As String, PrivateKey As PrivateKey) As signature
    Dim hashMessage As String: hashMessage = Utils.SHA256function(message)
    Dim numberMessage As String: numberMessage = BigIntMath.BigIntFromString(hashMessage, 16)
    
    Dim respJson As Dictionary: Set respJson = EllipticCurve_Math.randomPointOnCurve(PrivateKey.curve.Gx, PrivateKey.curve.Gy, PrivateKey.curve.n, PrivateKey.curve.A, PrivateKey.curve.P)
    If respJson("error").Count <> 0 Then
        Err.Raise number:=vbObjectError + 513, description:=respJson("error")("message")
    End If
    
    Dim randNum As String: randNum = respJson("success")("randNum")
    Dim xRandPoint As String: xRandPoint = respJson("success")("xRandSignPoint")
    Dim r As String: r = BigIntMath.Modulus(xRandPoint, PrivateKey.curve.n)
    
    Dim s1 As String: s1 = BigIntMath.Add(numberMessage, BigIntMath.multiply(r, PrivateKey.secret))
    Dim s2 As String: s2 = EllipticCurve_Math.inv(randNum, PrivateKey.curve.n)
    Dim S As String: S = BigIntMath.Modulus(BigIntMath.multiply(s1, s2), PrivateKey.curve.n)
    
    Dim sig As signature: Set sig = New signature
    Call sig.setProperties(r, S)
    Set sign = sig
End Function