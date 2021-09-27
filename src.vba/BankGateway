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
    
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getTransaction = resp.json()

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



Public Function postTransaction(payload As String, signature As String)
    Dim resp As response
    Dim headers As New Dictionary
    
    '--------------- Include signature in headers -----------------
    headers.Add "Digital-Signature", signature
    
    '--------------- Send request ---------------------------------
    Set resp = StarkBankApi.postRequest("/v1/bank/transaction", payload, headers)
    
    If resp.Status = 200 Then
        MsgBox "Transferência interna executada com sucesso!", , "Sucesso"
    
    ElseIf resp.error().Exists("errors") Then
        Dim errors As Collection: Set errors = resp.error()("errors")
        Dim error As Dictionary
        Dim errorList As String
        Dim errorDescription As String
        
        For Each error In errors
            errorDescription = Utils.correctErrorLine(error("message"), TableFormat.HeaderRow() + 1)
            errorList = errorList & errorDescription & vbNewLine
        Next
        
        Dim messageBox As String
        messageBox = resp.error()("message") & vbNewLine & vbNewLine & errorList
        MsgBox messageBox, , "Erro"
    Else
        MsgBox resp.error()("message"), vbExclamation, "Erro"
    End If
    
    Set postTransaction = resp.json()
    
End Function