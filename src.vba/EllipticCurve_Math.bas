Public Function randomPointOnCurve(px As String, py As String, n As String, A As String, P As String)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    Dim result As New Dictionary
    
    dict.Add "Gx", px
    dict.Add "Gy", py
    dict.Add "A", A
    dict.Add "P", P
    dict.Add "N", n
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set resp = StarkBankApi.externalPostRequest("https://us-central1-api-ms-auth-sbx.cloudfunctions.net/ellipticCurveMath", payload)
    
    If resp.Status = 200 Then
        result.Add "success", resp.json()
        result.Add "error", New Dictionary
        Set randomPointOnCurve = result
    Else
        result.Add "success", New Dictionary
        result.Add "error", resp.error()
        Set randomPointOnCurve = result
    End If

End Function




Public Function multiply(px As String, py As String, nn As String, n As String, A As String, P As String) As String
    multiply = fromJacobian(jacobianMultiply(toJacobian(px, py), nn, n, A, P), P)
End Function

Public Function toJacobian(x As String, y As String) As Point
    Dim pp As Point: Set pp = New Point
    Call pp.setCoordinates(x, y, "1")
    Set toJacobian = pp
End Function

Public Function fromJacobian(Point As Point, P As String) As String
    Dim z As String: z = inv(Point.z, P)
    Dim z2 As String: z2 = BigIntMath.multiply(z, z)
    Dim z3 As String: z3 = BigIntMath.multiply(z2, z)
    Dim x As String: x = BigIntMath.Modulus(BigIntMath.multiply(Point.x, z2), P)
    Dim y As String: y = BigIntMath.Modulus(BigIntMath.multiply(Point.y, z3), P)

    fromJacobian = x & ";" & y
End Function

Public Function inv(x As String, n As String) As String
    If BigIntMath.Compare(x, "0") = 0 Then
        inv = "0"
    Else
        Dim lm As String, hm As String, high As String, low As String
        lm = "1"
        hm = "0"
        high = n
        low = BigIntMath.Modulus(x, n)
        Dim r As String, nm As String, nw As String
        While BigIntMath.Compare(low, "1") = 1
            r = BigIntMath.Divide(high, low)
            nm = BigIntMath.Subtract(hm, BigIntMath.multiply(lm, r))
            nw = BigIntMath.Subtract(high, BigIntMath.multiply(low, r))
            high = low
            hm = lm
            low = nw
            lm = nm
        Wend
        inv = BigIntMath.Modulus(lm, n)
    End If
End Function

Public Function jacobianAdd(pointP As Point, pointQ As Point, A As String, P As String) As Point
    Dim pp As Point: Set pp = New Point
    If pointP.y = vbNullString Or BigIntMath.Compare(pointP.y, "0") = 0 Then
        Set jacobianAdd = pointQ
        
    ElseIf pointQ.y = vbNullString Or BigIntMath.Compare(pointQ.y, "0") = 0 Then
        Set jacobianAdd = pointP
        
    Else
        Dim U1 As String: U1 = BigIntMath.Modulus(BigIntMath.multiply(pointP.x, BigIntMath.multiply(pointQ.z, pointQ.z)), P)
        Dim U2 As String: U2 = BigIntMath.Modulus(BigIntMath.multiply(pointQ.x, BigIntMath.multiply(pointP.z, pointP.z)), P)
        Dim s1 As String: s1 = BigIntMath.Modulus(BigIntMath.multiply(pointP.y, BigIntMath.multiply(pointQ.z, BigIntMath.multiply(pointQ.z, pointQ.z))), P)
        Dim s2 As String: s2 = BigIntMath.Modulus(BigIntMath.multiply(pointQ.y, BigIntMath.multiply(pointP.z, BigIntMath.multiply(pointP.z, pointP.z))), P)
        
        If BigIntMath.Compare(U1, U2) = 0 Then
            If BigIntMath.Compare(s1, s2) <> 0 Then
                Call pp.setCoordinates("0", "0", "1")
                Set jacobianAdd = pp
            Else
                Set jacobianAdd = jacobianDouble(pointP, A, P)
            End If
        Else
            Dim H As String: H = BigIntMath.Subtract(U2, U1)
            Dim r As String: r = BigIntMath.Subtract(s2, s1)
            Dim H2 As String: H2 = BigIntMath.Modulus(BigIntMath.multiply(H, H), P)
            Dim H3 As String: H3 = BigIntMath.Modulus(BigIntMath.multiply(H, H2), P)
            Dim U1H2 As String: U1H2 = BigIntMath.Modulus(BigIntMath.multiply(U1, H2), P)
            Dim nx As String: nx = BigIntMath.Modulus(BigIntMath.Subtract(BigIntMath.multiply(r, r), BigIntMath.Add(H3, BigIntMath.multiply("2", U1H2))), P)
            Dim ny As String: ny = BigIntMath.Modulus(BigIntMath.Subtract(BigIntMath.multiply(r, BigIntMath.Subtract(U1H2, nx)), BigIntMath.multiply(s1, H3)), P)
            Dim nz As String: nz = BigIntMath.Modulus(BigIntMath.multiply(H, BigIntMath.multiply(pointP.z, pointQ.z)), P)
            
            Call pp.setCoordinates(nx, ny, nz)
            Set jacobianAdd = pp
        End If
    End If
End Function

Public Function jacobianMultiply(Point As Point, nn As String, n As String, A As String, P As String) As Point
    Dim pp As Point: Set pp = New Point
    If BigIntMath.Compare("0", Point.y) = 0 Or BigIntMath.Compare("0", nn) = 0 Then
        Call pp.setCoordinates("0", "0", "1")
        Set jacobianMultiply = pp
        
    ElseIf BigIntMath.Compare("1", nn) = 0 Then
        Set jacobianMultiply = Point
        
    ElseIf BigIntMath.Compare(nn, "0") = -1 Or BigIntMath.Compare(nn, n) <> -1 Then
        Set jacobianMultiply = jacobianMultiply(Point, BigIntMath.Modulus(nn, n), n, A, P)
        
    ElseIf BigIntMath.Compare(BigIntMath.Modulus(nn, 2), "0") = 0 Then
        Set jacobianMultiply = jacobianDouble(jacobianMultiply(Point, BigIntMath.Divide(nn, "2"), n, A, P), A, P)
    
    ElseIf BigIntMath.Compare(BigIntMath.Modulus(nn, 2), "1") = 0 Then
        Set jacobianMultiply = jacobianAdd(jacobianDouble(jacobianMultiply(Point, BigIntMath.Divide(nn, "2"), n, A, P), A, P), Point, A, P)
    
    End If
    
End Function

Public Function jacobianDouble(Point As Point, A As String, P As String) As Point
    Dim pp As Point: Set pp = New Point
    If Point.y = vbNullString Or BigIntMath.Compare(Point.y, "0") = 0 Then
        Call pp.setCoordinates("0", "0", "0")
        Set jacobianDouble = pp
    Else
        Dim ysq As String: ysq = BigIntMath.Modulus(BigIntMath.multiply(Point.y, Point.y), P)
        Dim s As String: s = BigIntMath.Modulus(BigIntMath.multiply("4", BigIntMath.multiply(Point.x, ysq)), P)
        Dim z4 As String: z4 = BigIntMath.multiply(Point.z, BigIntMath.multiply(Point.z, BigIntMath.multiply(Point.z, Point.z)))
        Dim M As String: M = BigIntMath.Modulus(BigIntMath.Add(BigIntMath.multiply("3", BigIntMath.multiply(Point.x, Point.x)), BigIntMath.multiply(A, z4)), P)
        Dim nx As String: nx = BigIntMath.Modulus(BigIntMath.Subtract(BigIntMath.multiply(M, M), BigIntMath.multiply("2", s)), P)
        
        Dim part1 As String: part1 = BigIntMath.multiply(M, BigIntMath.Subtract(s, nx))
        Dim part2 As String: part2 = BigIntMath.multiply("8", BigIntMath.multiply(ysq, ysq))
        Dim part3 As String: part3 = BigIntMath.Subtract(part1, part2)
        Dim part4 As String: part4 = BigIntMath.Modulus(part3, P)
        Dim ny As String: ny = BigIntMath.Modulus(BigIntMath.Subtract(BigIntMath.multiply(M, BigIntMath.Subtract(s, nx)), BigIntMath.multiply("8", BigIntMath.multiply(ysq, ysq))), P)
        Dim nz As String: nz = BigIntMath.Modulus(BigIntMath.multiply("2", BigIntMath.multiply(Point.y, Point.z)), P)
        
        Call pp.setCoordinates(nx, ny, nz)
        Set jacobianDouble = pp
    End If
End Function