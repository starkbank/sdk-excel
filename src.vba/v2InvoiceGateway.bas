Public Function getStatus(Status As String)
    Select Case Status
        Case "Todos": getStatus = "all"
        Case "Pagos":  getStatus = "paid"
        Case "Criados":  getStatus = "created"
        Case "Vencidos":  getStatus = "overdue"
        Case "Cancelados":  getStatus = "canceled"
        Case "Expirados":  getStatus = "expired"
    End Select
End Function

Public Function getStatusInPt(Status As String)
    Select Case Status
        Case "paid":  getStatusInPt = "pago"
        Case "voided":  getStatusInPt = "anulado"
        Case "created":  getStatusInPt = "criado"
        Case "overdue":  getStatusInPt = "vencido"
        Case "canceled":  getStatusInPt = "cancelado"
        Case "expired":  getStatusInPt = "expirado"
        Case "unknown":  getStatusInPt = "desconhecido"
    End Select
End Function

Public Function getEventInPt(Status As String)
    Select Case Status
        Case "created":  getEventInPt = "criado"
        Case "updated":  getEventInPt = "atualizado"
        Case "canceled":  getEventInPt = "cancelado"
        Case "overdue":  getEventInPt = "vencido"
        Case "expired":  getEventInPt = "expirado"
        Case "paid":  getEventInPt = "pago"
        Case "credited":  getEventInPt = "creditado"
        Case "reversing":  getEventInPt = "em reversão"
        Case "sending":  getEventInPt = "em envio"
        Case "sent":  getEventInPt = "enviado"
        Case "failed":  getEventInPt = "falha"
        Case "refunded":  getEventInPt = "estornado"
        Case "reversed":  getEventInPt = "revertido"
        Case "voided":  getEventInPt = "anulado"
        Case "unknown":  getEventInPt = "desconhecido"
        Case Else:  getEventInPt = Status
    End Select
End Function

Public Function getInvoices(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = V2Rest.getRequest("/v2/invoice", query, New Dictionary)
    Debug.Print resp.Status
    
    If resp.Status >= 300 Then
        MsgBox resp.errors()("errors")(1)("message"), , "Erro"
    End If
    Set getInvoices = resp.json()

End Function

Public Function getInvoiceLogs(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = V2Rest.getRequest("/v2/invoice/log", query, New Dictionary)
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getInvoiceLogs = resp.json()

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
    
    Set resp = V2Rest.getRequest("/v2/invoice/customer", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getCustomers = resp.json()

End Function

Public Function getOrders(initRow As Long, midRow As Long) As Collection
    Dim orders As New Collection
    
    For Each obj In SheetParser.longDict(initRow, midRow)
        Dim amount As Long
        Dim name As String
        Dim taxId As String
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
        amount = getAmountLong((obj("Valor")))
        name = Trim(obj("Nome do Cliente"))
        taxId = Trim(obj("CPF/CNPJ do Cliente"))
        dueDate = Utils.DateToSendingFormat(Format(obj("Data de Vencimento"), "dd/mm/yyyy")) + "T23:59:50.000+00:00"
        
        If obj("Multa") <> "" Then
            fine = Utils.SingleFrom((obj("Multa")))
        End If
        
        If obj("Juros ao Mês") <> "" Then
            interest = Utils.SingleFrom((obj("Juros ao Mês")))
        End If
        
        If obj("Dias para Baixa Automática") <> "" Then
            overdueLimit = CLng(3600) * 24 * Utils.IntegerFrom((obj("Dias para Baixa Automática")))
        End If
        
        If obj("Descrição 1") <> "" Then
            description1.Add "key", obj("Descrição 1")
        End If
        If obj("Valor 1") <> "" Then
            description1.Add "value", CStr(obj("Valor 1"))
        End If
        
        If obj("Descrição 2") <> "" Then
            description2.Add "key", obj("Descrição 2")
        End If
        If obj("Valor 2") <> "" Then
            description2.Add "value", CStr(obj("Valor 2"))
        End If
        
        If obj("Descrição 3") <> "" Then
            description3.Add "key", obj("Descrição 3")
        End If
        If obj("Valor 3") <> "" Then
            description3.Add "value", CStr(obj("Valor 3"))
        End If
        
        Set order = v2InvoiceGateway.order(amount, name, taxId, dueDate, fine, interest, overdueLimit, description1, description2, description3)
    
        orders.Add order
    Next
    Set getOrders = orders
End Function

Public Function order(amount As Long, name As String, taxId As String, dueDate As String, fine As Single, interest As Single, overdueLimit As Long, description1 As Dictionary, description2 As Dictionary, description3 As Dictionary) As Dictionary
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
    dict.Add "name", name
    dict.Add "taxId", taxId
    dict.Add "due", dueDate
    dict.Add "fine", fine
    dict.Add "interest", interest
    dict.Add "expiration", overdueLimit
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

Public Function createInvoices(invoices As Collection)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    
    dict.Add "invoices", invoices
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set createInvoices = V2Rest.postRequest("/v2/invoice", payload, New Dictionary)
    
End Function

Public Function createCustomers(customers As Collection)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    
    dict.Add "customers", customers
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set createCustomers = StarkBankApi.postRequest("/v2/invoice/customer", payload, New Dictionary)
    
End Function

Public Function getEventLog(invoiceId As String, logEvent As String, optionalParam As Dictionary)
    Dim query As String
    Dim resp As response
    Dim elem As Variant
    
    query = ""
    If invoiceId <> "" Then
        query = "?events=" + logEvent + "&invoiceIds=" + invoiceId
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
    Set resp = StarkBankApi.getRequest("/v2/invoice/log", query, New Dictionary)
    If resp.Status >= 300 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    Set getEventLog = resp.json()

End Function


