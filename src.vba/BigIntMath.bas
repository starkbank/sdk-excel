Private Type PartialDivideInfo
    Quotient As Integer
    Subtrahend As String
    Remainder As String
End Type

Private LastRemainder As String

Private Const Alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Function Compare(ByVal A As String, ByVal b As String) As Integer
    Dim an, bn, rn As Boolean
    Dim i, av, bv As Integer
    an = (Left(A, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then A = Mid(A, 2)
    If bn Then b = Mid(b, 2)
    If an And bn Then
        rn = True
    ElseIf bn Then
        Compare = 1
        Exit Function
    ElseIf an Then
        Compare = -1
        Exit Function
    Else
        rn = False
    End If
    Do While Len(A) > 1 And Left(A, 1) = "0"
        A = Mid(A, 2)
    Loop
    Do While Len(b) > 1 And Left(b, 1) = "0"
        b = Mid(b, 2)
    Loop
    If Len(A) < Len(b) Then
        Compare = -1
    ElseIf Len(A) > Len(b) Then
        Compare = 1
    Else
        Compare = 0
        For i = 1 To Len(A)
            av = CInt(Mid(A, i, 1))
            bv = CInt(Mid(b, i, 1))
            If av < bv Then
                Compare = -1
                Exit For
            ElseIf av > bv Then
                Compare = 1
                Exit For
            End If
        Next i
    End If
    If rn Then
        Compare = -Compare
    End If
End Function

Public Function Add(ByVal A As String, ByVal b As String) As String
    Dim an, bn, rn As Boolean
    Dim ai, bi, carry As Integer
    an = (Left(A, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then A = Mid(A, 2)
    If bn Then b = Mid(b, 2)
    If an And bn Then
        rn = True
    ElseIf bn Then
        Add = Subtract(A, b)
        Exit Function
    ElseIf an Then
        Add = Subtract(b, A)
        Exit Function
    Else
        rn = False
    End If
    ai = Len(A)
    bi = Len(b)
    carry = 0
    Add = ""
    Do While ai > 0 And bi > 0
        carry = carry + CInt(Mid(A, ai, 1)) + CInt(Mid(b, bi, 1))
        Add = CStr(carry Mod 10) + Add
        carry = carry \ 10
        ai = ai - 1
        bi = bi - 1
    Loop
    Do While ai > 0
        carry = carry + CInt(Mid(A, ai, 1))
        Add = CStr(carry Mod 10) + Add
        carry = carry \ 10
        ai = ai - 1
    Loop
    Do While bi > 0
        carry = carry + CInt(Mid(b, bi, 1))
        Add = CStr(carry Mod 10) + Add
        carry = carry \ 10
        bi = bi - 1
    Loop
    Add = CStr(carry) + Add
    Do While Len(Add) > 1 And Left(Add, 1) = "0"
        Add = Mid(Add, 2)
    Loop
    If Add <> "0" And rn Then
        Add = "-" + Add
    End If
End Function

Private Function RealMod(ByVal A As Integer, ByVal b As Integer) As Integer
    If A Mod b = 0 Then
        RealMod = 0
    ElseIf A < 0 Then
        RealMod = b + A Mod b
    Else
        RealMod = A Mod b
    End If
End Function

Private Function RealDiv(ByVal A As Integer, ByVal b As Integer) As Integer
    If A Mod b = 0 Then
        RealDiv = A \ b
    ElseIf A < 0 Then
        RealDiv = A \ b - 1
    Else
        RealDiv = A \ b
    End If
End Function

Public Function Subtract(ByVal A As String, ByVal b As String) As String
    Dim an, bn, rn As Boolean
    Dim ai, bi, barrow As Integer
    an = (Left(A, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then A = Mid(A, 2)
    If bn Then b = Mid(b, 2)
    If an And bn Then
        rn = True
    ElseIf bn Then
        Subtract = Add(A, b)
        Exit Function
    ElseIf an Then
        Subtract = "-" + Add(A, b)
        Exit Function
    Else
        rn = False
    End If
    barrow = Compare(A, b)
    If barrow = 0 Then
        Subtract = "0"
        Exit Function
    ElseIf barrow < 0 Then
        Subtract = A
        A = b
        b = Subtract
        rn = Not rn
    End If
    ai = Len(A)
    bi = Len(b)
    barrow = 0
    Subtract = ""
    Do While ai > 0 And bi > 0
        barrow = barrow + CInt(Mid(A, ai, 1)) - CInt(Mid(b, bi, 1))
        Subtract = CStr(RealMod(barrow, 10)) + Subtract
        barrow = RealDiv(barrow, 10)
        ai = ai - 1
        bi = bi - 1
    Loop
    Do While ai > 0
        barrow = barrow + CInt(Mid(A, ai, 1))
        Subtract = CStr(RealMod(barrow, 10)) + Subtract
        barrow = RealDiv(barrow, 10)
        ai = ai - 1
    Loop
    Do While Len(Subtract) > 1 And Left(Subtract, 1) = "0"
        Subtract = Mid(Subtract, 2)
    Loop
    If Subtract <> "0" And rn Then
        Subtract = "-" + Subtract
    End If
End Function

Public Function multiply(ByVal A As String, ByVal b As String) As String
    Dim an, bn, rn As Boolean
    Dim M() As Long
    Dim al, bl, ai, bi As Integer
    Dim carry As Long
    an = (Left(A, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then A = Mid(A, 2)
    If bn Then b = Mid(b, 2)
    rn = (an <> bn)
    al = Len(A)
    bl = Len(b)
    ReDim M(1 To (al + bl - 1))
    For ai = 1 To al
        For bi = 1 To bl
            M(ai + bi - 1) = M(ai + bi - 1) + CLng(Mid(A, al - ai + 1, 1)) * CLng(Mid(b, bl - bi + 1, 1))
        Next bi
    Next ai
    carry = 0
    multiply = ""
    For ai = 1 To al + bl - 1
        carry = carry + M(ai)
        multiply = CStr(carry Mod 10) + multiply
        carry = carry \ 10
    Next ai
    multiply = CStr(carry) + multiply
    Do While Len(multiply) > 1 And Left(multiply, 1) = "0"
        multiply = Mid(multiply, 2)
    Loop
    If multiply <> "0" And rn Then
        multiply = "-" + multiply
    End If
End Function

Private Function PartialDivide(A As String, b As String) As PartialDivideInfo
    For PartialDivide.Quotient = 9 To 1 Step -1
        PartialDivide.Subtrahend = multiply(b, CStr(PartialDivide.Quotient))
        If Compare(PartialDivide.Subtrahend, A) <= 0 Then
            PartialDivide.Remainder = Subtract(A, PartialDivide.Subtrahend)
            Exit Function
        End If
    Next PartialDivide.Quotient
    PartialDivide.Quotient = 0
    PartialDivide.Subtrahend = "0"
    PartialDivide.Remainder = A
End Function

Public Function Divide(ByVal A As String, ByVal b As String) As String
    Dim an, bn, rn As Boolean
    Dim c As Integer
    Dim S As String
    Dim d As PartialDivideInfo
    an = (Left(A, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then A = Mid(A, 2)
    If bn Then b = Mid(b, 2)
    rn = (an <> bn)
    If Compare(b, "0") = 0 Then
        Err.Raise 11
        Exit Function
    ElseIf Compare(A, "0") = 0 Then
        Divide = "0"
        LastRemainder = "0"
        Exit Function
    End If
    c = Compare(A, b)
    If c < 0 Then
        Divide = "0"
        LastRemainder = A
        Exit Function
    ElseIf c = 0 Then
        If rn Then
            Divide = "-1"
        Else
            Divide = "1"
        End If
        LastRemainder = "0"
        Exit Function
    End If
    Divide = ""
    S = ""
    For c = 1 To Len(A)
        S = S + Mid(A, c, 1)
        d = PartialDivide(S, b)
        Divide = Divide + CStr(d.Quotient)
        S = d.Remainder
    Next c
    Do While Len(Divide) > 1 And Left(Divide, 1) = "0"
        Divide = Mid(Divide, 2)
    Loop
    If Divide <> "0" And rn Then
        Divide = "-" + Divide
    End If
    LastRemainder = S
End Function

Public Function LastModulus() As String
    LastModulus = LastRemainder
End Function

Public Function Modulus(ByVal A As String, ByVal b As String) As String
    Dim an As Boolean
    an = (Left(A, 1) = "-")
    
    Divide A, b
    If an Then
        A = Mid(A, 2)
        Modulus = Subtract(b, LastRemainder)
    Else
        Modulus = LastRemainder
    End If
End Function

Public Function BigIntFromString(ByVal S As String, ByVal base As Integer) As String
    Dim rn As Boolean
    Dim bs As String
    Dim i, v As Integer
    If Left(S, 1) = "-" Then
        rn = True
        S = Mid(S, 2)
    Else
        rn = False
    End If
    bs = CStr(base)
    BigIntFromString = "0"
    For i = 1 To Len(S)
        v = InStr(Alphabet, UCase(Mid(S, i, 1)))
        If v > 0 Then
            BigIntFromString = multiply(BigIntFromString, bs)
            BigIntFromString = Add(BigIntFromString, CStr(v - 1))
        End If
    Next i
    If rn Then
        BigIntFromString = "-" + BigIntFromString
    End If
End Function

Public Function BigIntToString(ByVal S As String, ByVal base As Integer) As String
    Dim rn As Boolean
    Dim bs As String
    Dim v As Integer
    If Left(S, 1) = "-" Then
        rn = True
        S = Mid(S, 2)
    Else
        rn = False
    End If
    bs = CStr(base)
    BigIntToString = ""
    Do While Compare(S, "0") > 0
        S = Divide(S, bs)
        v = CInt(LastModulus())
        BigIntToString = Mid(Alphabet, v + 1, 1) + BigIntToString
    Loop
    If BigIntToString = "" Then
        BigIntToString = "0"
    ElseIf BigIntToString <> "0" And rn Then
        BigIntToString = "-" + BigIntToString
    End If
End Function