Public Sub ListKeys()
    
End Sub

Public Function getDictKey(cursor As String)
    Dim resp As response
    
    Set resp = V2Rest.getRequest("/v2/dict-key", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.errors()("errors")(1)("message"), , "Erro"
    End If
    Set getDictKey = resp.json()
End Function