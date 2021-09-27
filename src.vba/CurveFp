Public a As String
Public b As String
Public P As String
Public n As String
Public Gx As String
Public Gy As String
Public name As String
Public oid As String

Public Function contains(x As String, y As String) As Boolean
    Dim y2 As String: y2 = BigIntMath.multiply(y, y)
    Dim x2 As String: x2 = BigIntMath.multiply(x, x)
    Dim x3 As String: x3 = BigIntMath.multiply(x2, x)
    
    Dim cx As String: cx = BigIntMath.Add(x3, BigIntMath.Add(BigIntMath.multiply(a, x), b))
    Dim cxy As String: cxy = BigIntMath.Subtract(y2, cx)
    
    contains = (BigIntMath.Modulus(cxy, P) = "0")
    
End Function

Public Function length()
    Dim Ncopy As String: Ncopy = n
    Dim hexaN As String: hexaN = BigIntMath.BigIntToString(Ncopy, 16)
    length = BigIntMath.Divide(BigIntMath.Add("1", Len(hexaN)), "2")
End Function

Public Function initializeByOid(oidInput As String)
    Dim Result As Boolean: Result = False
    Select Case oidInput
        Case "1;3;132;0;10":
            a = BigIntMath.BigIntFromString("0000000000000000000000000000000000000000000000000000000000000000", 16)
            b = BigIntMath.BigIntFromString("0000000000000000000000000000000000000000000000000000000000000007", 16)
            P = BigIntMath.BigIntFromString("fffffffffffffffffffffffffffffffffffffffffffffffffffffffefffffc2f", 16)
            n = BigIntMath.BigIntFromString("fffffffffffffffffffffffffffffffebaaedce6af48a03bbfd25e8cd0364141", 16)
            Gx = BigIntMath.BigIntFromString("79be667ef9dcbbac55a06295ce870b07029bfcdb2dce28d959f2815b16f81798", 16)
            Gy = BigIntMath.BigIntFromString("483ada7726a3c4655da4fbfc0e1108a8fd17b448a68554199c47d08ffb10d4b8", 16)
            oid = "1;3;132;0;10"
            Result = True
    End Select
    initializeByOid = Result
End Function