Public Function getStatus(Status As String)
    Select Case Status
        Case "Todos": getStatus = "all"
        Case "Pagos":  getStatus = "paid"
        Case "Pendentes de Registro":  getStatus = "created"
        Case "Registrados":  getStatus = "registered"
        Case "Vencidos":  getStatus = "overdue"
        Case "Cancelados":  getStatus = "canceled"
    End Select
End Function

Public Function getStatusInPt(Status As String)
    Select Case Status
        Case "paid":  getStatusInPt = "pago"
        Case "created":  getStatusInPt = "pendente de registro"
        Case "registered":  getStatusInPt = "registrado"
        Case "overdue":  getStatusInPt = "vencido"
        Case "canceled":  getStatusInPt = "cancelado"
        Case "failed":  getStatusInPt = "falha"
        Case "unknown":  getStatusInPt = "desconhecido"
    End Select
End Function

Public Function getEventInPt(Status As String)
    Select Case Status
        Case "paid":  getEventInPt = "pago"
        Case "bank":  getEventInPt = "creditado"
        Case "register":  getEventInPt = "criado (pendente de registro)"
        Case "registered":  getEventInPt = "registrado"
        Case "overdue":  getEventInPt = "vencido"
        Case "cancel":  getEventInPt = "em cancelamento"
        Case "canceled":  getEventInPt = "cancelado"
        Case "failed":  getEventInPt = "falha"
        Case "unknown":  getEventInPt = "desconhecido"
        Case Else:  getEventInPt = Status
    End Select
End Function

Public Function getStatusFromId(id As String)
    Select Case id
        Case "00":  getStatusFromId = "register"
        Case "02":  getStatusFromId = "registered"
        Case "03":  getStatusFromId = "failed"
        Case "06":  getStatusFromId = "paid"
        Case "09":  getStatusFromId = "canceled"
        Case Else:  getStatusFromId = "unknown"
    End Select
End Function

Public Function getOccurrenceId(statusCode As String)
    Select Case statusCode
        Case "pendente de registro"
            getOccurrenceId = "00"
        Case "registrado"
            getOccurrenceId = "02"
        Case "vencido"
            getOccurrenceId = "02"
        Case "falha"
            getOccurrenceId = "03"
        Case "pago"
            getOccurrenceId = "06"
        Case "cancelado"
            getOccurrenceId = "09"
        Case Else
            getOccurrenceId = "99"
    End Select
End Function

Public Function getCharges(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = StarkBankApi.getRequest("/v1/charge", query, New Dictionary)
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getCharges = resp.json()

End Function

Public Function getChargeLogs(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = StarkBankApi.getRequest("/v1/charge/log", query, New Dictionary)
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getChargeLogs = resp.json()

End Function

Public Function getCustomers(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = StarkBankApi.getRequest("/v1/charge/customer", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getCustomers = resp.json()

End Function

Public Function getOrders(initRow As Long, midRow As Long) As Collection
    Dim orders As New Collection
    
    For Each obj In SheetParser.longDict(initRow, midRow)
        Dim amount As Long
        Dim customerId As String
        Dim dueDate As String
        Dim fine As Single: fine = 2
        Dim interest As Single: interest = 1
        Dim overdueLimit As Long: overdueLimit = 59
        Dim description1 As Dictionary: Set description1 = New Dictionary
        Dim description2 As Dictionary: Set description2 = New Dictionary
        Dim description3 As Dictionary: Set description3 = New Dictionary
        Dim order As Dictionary
        
        If obj("Valor") = "" Then
            MsgBox "Por favor, não deixe linhas em branco entre as ordens de cobrança", , "Erro"
            End
        End If
        amount = Utils.IntegerFrom((obj("Valor")))
        customerId = Trim(obj("Id do Cliente"))
        dueDate = Utils.DateToSendingFormat(Format(obj("Data de Vencimento"), "dd/mm/yyyy"))
        
        If obj("Multa") <> "" Then
            fine = Utils.SingleFrom((obj("Multa")))
        End If
        
        If obj("Juros ao Mês") <> "" Then
            interest = Utils.SingleFrom((obj("Juros ao Mês")))
        End If
        
        If obj("Dias para Baixa Automática") <> "" Then
            overdueLimit = Utils.IntegerFrom((obj("Dias para Baixa Automática")))
        End If
        
        If obj("Descrição 1") <> "" Then
            description1.Add "text", obj("Descrição 1")
        End If
        If obj("Valor 1") <> "" Then
            description1.Add "amount", Utils.IntegerFrom((obj("Valor 1")))
        End If
        
        If obj("Descrição 2") <> "" Then
            description2.Add "text", obj("Descrição 2")
        End If
        If obj("Valor 2") <> "" Then
            description2.Add "amount", Utils.IntegerFrom((obj("Valor 2")))
        End If
        
        If obj("Descrição 3") <> "" Then
            description3.Add "text", obj("Descrição 3")
        End If
        If obj("Valor 3") <> "" Then
            description3.Add "amount", Utils.IntegerFrom((obj("Valor 3")))
        End If
        
        Set order = ChargeGateway.order(amount, customerId, dueDate, fine, interest, overdueLimit, description1, description2, description3)
    
        orders.Add order
    Next
    Set getOrders = orders
End Function

Public Function getCustomerOrders(initRow As Long, midRow As Long) As Collection
    Dim orders As New Collection
    
    For Each obj In SheetParser.longDict(initRow, midRow)
        Dim name As String
        Dim taxId As String
        Dim email As String
        Dim phone As String
        Dim streetLine1 As String
        Dim streetLine2 As String
        Dim district As String
        Dim city As String
        Dim stateCode As String
        Dim zipCode As String
        Dim tags As String
        Dim order As Dictionary
        
        If obj("Nome") = "" Then
            MsgBox "Por favor, não deixe linhas em branco entre os clientes para cadastro", vbExclamation, "Erro"
            End
        End If
        name = obj("Nome")
        taxId = obj("CPF/CNPJ")
        email = obj("E-mail")
        phone = obj("Telefone")
        streetLine1 = obj("Logradouro")
        streetLine2 = obj("Complemento")
        district = obj("Bairro")
        city = obj("Cidade")
        stateCode = obj("Estado")
        zipCode = obj("CEP")
        tags = obj("Tags")
        
        Set order = ChargeGateway.customerOrder(name, taxId, email, phone, streetLine1, streetLine2, district, city, stateCode, zipCode, tags)
    
        orders.Add order
        
    Next
    Set getCustomerOrders = orders
End Function

Public Function order(amount As Long, customerId As String, dueDate As String, fine As Single, interest As Single, overdueLimit As Long, description1 As Dictionary, description2 As Dictionary, description3 As Dictionary) As Dictionary
    Dim dict As New Dictionary
    Dim descriptions As New Collection
    
    If description1.Count > 0 Then
        descriptions.Add description1
    End If
    If description2.Count > 0 Then
        descriptions.Add description2
    End If
    If description3.Count > 0 Then
        descriptions.Add description3
    End If
    
    dict.Add "amount", amount
    dict.Add "customerId", customerId
    dict.Add "dueDate", dueDate
    dict.Add "fine", fine
    dict.Add "interest", interest
    dict.Add "overdueLimit", overdueLimit
    dict.Add "descriptions", descriptions
    
    Set order = dict
    
End Function

Public Function customerOrder(name As String, taxId As String, email As String, phone As String, streetLine1 As String, streetLine2 As String, district As String, city As String, stateCode As String, zipCode As String, tags As String) As Dictionary
    Dim dict As New Dictionary
    Dim address As New Dictionary
    
    address.Add "streetLine1", streetLine1
    address.Add "streetLine2", streetLine2
    address.Add "district", district
    address.Add "city", city
    address.Add "stateCode", stateCode
    address.Add "zipCode", zipCode
    
    dict.Add "name", name
    dict.Add "taxId", taxId
    dict.Add "email", email
    dict.Add "phone", phone
    dict.Add "address", address
    dict.Add "tags", Split(tags, ",")
    
    Set customerOrder = dict
    
End Function

Public Function createCharges(charges As Collection)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    
    dict.Add "charges", charges
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set createCharges = StarkBankApi.postRequest("/v1/charge", payload, New Dictionary)
    
End Function

Public Function createCustomers(customers As Collection)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    
    dict.Add "customers", customers
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set createCustomers = StarkBankApi.postRequest("/v1/charge/customer", payload, New Dictionary)
    
End Function

Public Function getEventLog(chargeId As String, logEvent As String, optionalParam As Dictionary)
    Dim query As String
    Dim resp As response
    Dim elem As Variant
    
    query = ""
    If chargeId <> "" Then
        query = "?events=" + logEvent + "&chargeIds=" + chargeId
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
    Set resp = StarkBankApi.getRequest("/v1/charge/log", query, New Dictionary)
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getEventLog = resp.json()

End Function

