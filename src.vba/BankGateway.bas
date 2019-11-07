Public Function getTransaction(cursor As String, optionalParam As Dictionary)
    Dim query As String
    Dim resp As response
    
    query = ""
    If cursor <> "" Then
        query = "?cursor=" + cursor
    End If
    
    If optionalParam.Count > 0 Then
        For Each key In optionalParam
            If query = "" Then
                query = "?" + key + "=" + CStr(optionalParam(key))
            Else
                query = query + "&" + key + "=" + CStr(optionalParam(key))
            End If
        Next
    End If
    
    Set resp = StarkBankApi.getRequest("/v1/bank/transaction", query, New Dictionary)
    
    If resp.Status = 200 Then
        Set getTransaction = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
        Set getTransaction = New Dictionary
    End If

End Function

Public Function getAccount()
    Dim resp As response
    Dim baseUrl As String
    Dim workspaceId As String
    
    workspaceId = SessionGateway.getWorkspaceId()
    
    baseUrl = "/v1/bank/account/" + workspaceId
    
    Set resp = StarkBankApi.getRequest(baseUrl, "", New Dictionary)
    
    If resp.Status = 200 Then
        Set getAccount = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
        Set getAccount = New Dictionary
    End If

End Function