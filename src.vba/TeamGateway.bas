Public Function getTeams(cursor As String, optionalParam As Dictionary)
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

    Set resp = StarkBankApi.getRequest("/v1/team", query, New Dictionary)
    
    If resp.Status = 200 Then
        Set getTeams = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
    End If

End Function

Public Function createOrders(teamId As String, orders As Collection)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    
    dict.Add "teamId", teamId
    dict.Add "orders", orders
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set resp = StarkBankApi.postRequest("/v1/team/order", payload, New Dictionary)
    
    If resp.Status = 200 Then
        createOrders = resp.json()("message")
        MsgBox resp.json()("message"), , "Sucesso"
        
    ElseIf resp.error().Exists("errors") Then
        Dim errors As Collection: Set errors = resp.error()("errors")
        Dim error As Dictionary
        Dim errorList As String
        Dim errorDescription As String
        
        For Each error In errors
            errorDescription = Utils.correctErrorLine(error("message"), TableFormat.HeaderRow() + 1)
            errorList = errorList & errorDescription & Chr(10)
        Next
        
        Dim messageBox As String
        messageBox = resp.error()("message") & Chr(10) & Chr(10) & errorList
        MsgBox messageBox, , "Erro"
        
    Else
        MsgBox resp.error()("message"), , "Erro"
        
    End If
    
End Function

Public Function getOrders() As Collection
    Dim orders As New Collection
    
    For Each obj In SheetParser.dict
        Dim amount As Long
        Dim taxId As String
        Dim name As String
        Dim bankCode As String
        Dim branchCode As String
        Dim accountNumber As String
        Dim tags() As String
        Dim description As String
        Dim order As Dictionary
        
        If obj("Valor") = "" Then
            MsgBox "Por favor, não deixe linhas em branco entre as ordens de transferência", , "Erro"
            Unload SendOrderForm
            End
        End If
        amount = Utils.IntegerFrom((obj("Valor")))
        taxId = Trim(obj("CPF/CNPJ"))
        name = Trim(obj("Nome"))
        bankCode = Trim(obj("Código do Banco"))
        branchCode = Trim(obj("Agência"))
        accountNumber = Trim(obj("Conta"))
        tags = Split(obj("Tags"), ",")
        description = obj("Descrição")
        
        Set order = TeamGateway.order(amount, taxId, name, bankCode, branchCode, accountNumber, tags, description)
        orders.Add order
        
    Next
    Set getOrders = orders
End Function

Public Function order(amount As Long, taxId As String, name As String, bankCode As String, branchCode As String, accountNumber As String, tags() As String, description As String) As Dictionary
    Dim dict As New Dictionary
    
    dict.Add "amount", amount
    dict.Add "taxId", taxId
    dict.Add "name", name
    dict.Add "bankCode", bankCode
    dict.Add "branchCode", branchCode
    dict.Add "accountNumber", accountNumber
    dict.Add "tags", tags
    dict.Add "description", description
    
    Set order = dict
    
End Function

Public Function getOrdersByTransfer(cursor As String, optionalParam As Dictionary)
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
    
    If resp.Status = 200 Then
        Set getOrdersByTransfer = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
    End If

End Function