Public Function getStatus(Status As String)
    Select Case Status
        Case "Todos": getStatus = "all"
        Case "Criados":  getStatus = "created"
        Case "Cancelados":  getStatus = "canceled"
        Case "Processando":  getStatus = "processing"
        Case "Pagos":  getStatus = "success"
        Case "Falha":  getStatus = "failed"
    End Select
End Function

Public Function getStatusInPt(Status As String)
    Select Case Status
        Case "created":  getStatusInPt = "criado"
        Case "canceled":  getStatusInPt = "cancelado"
        Case "processing":  getStatusInPt = "processando"
        Case "success":  getStatusInPt = "pago"
        Case "failed":  getStatusInPt = "falha"
        Case "unknown":  getStatusInPt = "desconhecido"
    End Select
End Function

Public Function getChargePayments(cursor As String, optionalParam As Dictionary)
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
    
    Set resp = StarkBankApi.getRequest("/v1/charge-payment", query, New Dictionary)
    
    If resp.Status = 200 Then
        Set getChargePayments = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
    End If

End Function


Public Function getChargePaymentsFromSheet() As Collection
    Dim payments As New Collection
    
    For Each obj In SheetParser.dict
        Dim lineOrBarCode As String
        Dim taxId As String
        Dim scheduled As String
        Dim description As String
        Dim tags() As String
        Dim payment As Dictionary
        
        If obj("Linha Digitável ou Código de Barras") = "" Then
            MsgBox "Por favor, não deixe linhas em branco entre as ordens de pagamento de boleto", , "Erro"
            Unload PayChargesForm
            End
        End If
        lineOrBarCode = Trim(obj("Linha Digitável ou Código de Barras"))
        taxId = Trim(obj("CPF/CNPJ do Beneficiário"))
        scheduled = Utils.DateToSendingFormat((obj("Data de Agendamento")))
        description = obj("Descrição")
        tags = Split(obj("Tags"), ",")
        
        Set payment = ChargePaymentGateway.payment(lineOrBarCode, taxId, scheduled, description, tags)
        payments.Add payment
        
    Next
    Set getChargePaymentsFromSheet = payments
End Function

Public Function payment(lineOrBarCode As String, taxId As String, scheduled As String, description As String, tags() As String) As Dictionary
    Dim dict As New Dictionary
    
    If Len(lineOrBarCode) = 44 Then
        dict.Add "barCode", lineOrBarCode
    Else
        dict.Add "line", lineOrBarCode
    End If
    dict.Add "taxId", taxId
    dict.Add "scheduled", scheduled
    dict.Add "description", description
    dict.Add "tags", tags
    
    Set payment = dict
    
End Function

Public Function createPayments(payload As String, signature As String)
    Dim resp As response
    Dim headers As New Dictionary
    
    '--------------- Include signature in headers -----------------
    headers.Add "Digital-Signature", signature
    
    '--------------- Send request ---------------------------------
    Set resp = StarkBankApi.postRequest("/v1/charge-payment", payload, headers)
    
    If resp.Status = 200 Then
        createPayments = resp.json()("message")
        MsgBox resp.json()("message"), , "Sucesso"
    
    ElseIf resp.error().Exists("errors") Then
        Dim errors As Collection: Set errors = resp.error()("errors")
        Dim error As Dictionary
        Dim errorList As String
        Dim errorDescription As String
        
        For Each error In errors
            errorDescription = Utils.correctErrorLine(error("message"), 10)
            errorList = errorList & errorDescription & Chr(10)
        Next
        
        Dim messageBox As String
        messageBox = resp.error()("message") & Chr(10) & Chr(10) & errorList
        MsgBox messageBox, , "Erro"
        
    Else
        MsgBox resp.error()("message"), , "Erro"
        
    End If
    
End Function
