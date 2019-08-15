Public Function sign(message As String, privateKey As privateKey) As signature
    Dim hashMessage As String: hashMessage = Utils.SHA256function(message)
    Dim numberMessage As String: numberMessage = BigIntMath.BigIntFromString(hashMessage, 16)
    
    Dim respJson As Dictionary: Set respJson = EllipticCurve_Math.randomPointOnCurve(privateKey.curve.Gx, privateKey.curve.Gy, privateKey.curve.N, privateKey.curve.A, privateKey.curve.P)
    If respJson("error").Count <> 0 Then
        Err.Raise number:=vbObjectError + 513, description:=respJson("error")("message")
    End If
    
    Dim randNum As String: randNum = respJson("success")("randNum")
    Dim xRandPoint As String: xRandPoint = respJson("success")("xRandSignPoint")
    Dim r As String: r = BigIntMath.Modulus(xRandPoint, privateKey.curve.N)
    
    Dim s1 As String: s1 = BigIntMath.Add(numberMessage, BigIntMath.multiply(r, privateKey.secret))
    Dim s2 As String: s2 = EllipticCurve_Math.inv(randNum, privateKey.curve.N)
    Dim s As String: s = BigIntMath.Modulus(BigIntMath.multiply(s1, s2), privateKey.curve.N)
    
    Dim sig As signature: Set sig = New signature
    Call sig.setProperties(r, s)
    Set sign = sig
End Function