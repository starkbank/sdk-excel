Public Function getStatus(Status As String)
    Select Case Status
        Case "Todos": getStatus = "all"
        Case "Sucesso":  getStatus = "success"
        Case "Processando":  getStatus = "processing"
        Case "Falha":  getStatus = "failed"
    End Select
End Function

Public Function getStatusInPt(Status As String)
    Select Case Status
        Case "success":  getStatusInPt = "sucesso"
        Case "processing":  getStatusInPt = "processando"
        Case "failed":  getStatusInPt = "falha"
        Case "unknown":  getStatusInPt = "desconhecido"
    End Select
End Function

Public Function getTransfers(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = StarkBankApi.getRequest("/v1/transfer", query, New Dictionary)
    
    If resp.Status = 200 Then
        Set getTransfers = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
    End If

End Function