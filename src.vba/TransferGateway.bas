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
        For Each key In optionalParam
            If query = "" Then
                query = "?" + key + "=" + CStr(optionalParam(key))
            Else
                query = query + "&" + key + "=" + CStr(optionalParam(key))
            End If
        Next
    End If
    
    Set resp = StarkBankApi.getRequest("/v1/transfer", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getTransfers = resp.json()
    
End Function

Public Function getTransferLogs(cursor As String, optionalParam As Dictionary)

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
    
    Set resp = StarkBankApi.getRequest("/v1/transfer/log", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getTransferLogs = resp.json()
    
End Function

Public Function getOrders(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = StarkBankApi.getRequest("/v1/team/order", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getOrders = resp.json()

End Function

Public Function getTransfersFromSheet() As Collection
    Dim transfers As New Collection
    
    For Each obj In SheetParser.dict
        Dim amount As Long
        Dim taxId As String
        Dim name As String
        Dim bankCode As String
        Dim branchCode As String
        Dim accountNumber As String
        Dim tags() As String
        Dim transfer As Dictionary
        
        If obj("Valor") = "" Then
            MsgBox "Por favor, não deixe linhas em branco entre as ordens de transferência", , "Erro"
            Unload SendTransferForm
            End
        End If
        amount = getAmountLong((obj("Valor")))
        taxId = Trim(obj("CPF/CNPJ"))
        name = Trim(obj("Nome"))
        bankCode = Trim(obj("Código do Banco"))
        branchCode = Trim(obj("Agência"))
        accountNumber = Trim(obj("Conta"))
        tags = Split(obj("Tags"), ",")
        
        Set transfer = TransferGateway.transfer(amount, taxId, name, bankCode, branchCode, accountNumber, tags)
        transfers.Add transfer
        
    Next
    Set getTransfersFromSheet = transfers
End Function

Public Function transfer(amount As Long, taxId As String, name As String, bankCode As String, branchCode As String, accountNumber As String, tags() As String) As Dictionary
    Dim dict As New Dictionary
    
    dict.Add "amount", amount
    dict.Add "taxId", taxId
    dict.Add "name", name
    dict.Add "bankCode", bankCode
    dict.Add "branchCode", branchCode
    dict.Add "accountNumber", accountNumber
    dict.Add "tags", tags
    
    Set transfer = dict
    
End Function

Public Function createTransfers(payload As String, signature As String)
    Dim resp As response
    Dim headers As New Dictionary
    
    '--------------- Include signature in headers -----------------
    headers.Add "Digital-Signature", signature
    
    '--------------- Send request ---------------------------------
    Set resp = StarkBankApi.postRequest("/v1/transfer", payload, headers)
    
    If resp.Status = 200 Then
        MsgBox "Transferências executadas com sucesso!", , "Sucesso"
    
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
    
    Set createTransfers = resp.json()
    
End Function