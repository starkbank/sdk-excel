
Public Function getRequestStatusInPt(Status As String)
    Select Case Status
        Case "pending":  getRequestStatusInPt = "Pendente de aprovação"
        Case "approved":  getRequestStatusInPt = "Pendente de pagamento"
        Case "scheduled":  getRequestStatusInPt = "Agendado"
        Case "processing":  getRequestStatusInPt = "Em processamento"
        Case "success":  getRequestStatusInPt = "Sucesso"
        Case "failed":  getRequestStatusInPt = "Falha"
        Case "denied":  getRequestStatusInPt = "Negado"
        Case "canceled":  getRequestStatusInPt = "Cancelado"
        Case Else: getRequestStatusInPt = Status
    End Select
End Function

Public Function getRequestTypeInPt(paymentType As String)
    Select Case paymentType
        Case "transfer":  getRequestTypeInPt = "Transferência"
        Case "boleto-payment":  getRequestTypeInPt = "Pagamento de boleto"
        Case "utility-payment":  getRequestTypeInPt = "Pagamento de concessionária"
        Case "tax-payment":  getRequestTypeInPt = "Pagamento de imposto"
        Case "brcode-payment":  getRequestTypeInPt = "Pagamento de QR Code"
        Case Else: getRequestTypeInPt = paymentType
    End Select
End Function


Public Function getCostCenters(cursor As String, optionalParam As Dictionary)
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

    Set resp = V2Rest.getRequest("/v2/cost-center", query, New Dictionary)
    
    If resp.Status <> 200 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    
    Set getCostCenters = resp.json()

End Function

Public Function getPaymentRequests(cursor As String, optionalParam As Dictionary)
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

    Set resp = V2Rest.getRequest("/v2/payment-request", query, New Dictionary)
    
    If resp.Status <> 200 Then
        Dim errors As Collection: Set errors = resp.errors()("errors")
        Dim error As Dictionary
        Dim errorList As String
        Dim errorDescription As String
        
        For Each error In errors
            errorDescription = error("message")
            errorList = errorList & errorDescription & Chr(10)
        Next
        
        Dim messageBox As String
        messageBox = errorList
        MsgBox messageBox, , "Erro"
        End
    End If
    
    Set getPaymentRequests = resp.json()

End Function
