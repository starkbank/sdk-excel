Public Function getTransaction(cursor As String, optionalParam As Dictionary)
    Dim query As String
    Dim resp As response
    
    query = ""
    If cursor <> "" Then
        query = "?cursor=" + cursor
    End If
    
    If optionalParam.Count > 0 Then
        For Each Key In optionalParam
            If query = "" Then
                query = "?" + Key + "=" + CStr(optionalParam(Key))
            Else
                query = query + "&" + Key + "=" + CStr(optionalParam(Key))
            End If
        Next
    End If
    
    Debug.Print "query: " + query
    
    Set resp = StarkBankApi.getRequest("/v1/bank/transaction", query, New Dictionary)
    
    If resp.Status = 200 Then
        Set getTransaction = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
        Set getTransaction = New Dictionary
    End If

End Function