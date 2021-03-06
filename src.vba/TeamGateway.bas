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
    
    If resp.Status <> 200 Then
        MsgBox resp.error()("message"), , "Erro"
    End If
    
    Set getTeams = resp.json()

End Function

Public Sub clearOrders()
    Sheets("Transferências Com Aprovação").Range("A10:Z" & Rows.Count).ClearContents
End Sub

Public Function getAccountType(accountType As Variant) As String
    Select Case accountType
        Case "": getAccountType = "checking"
        Case "Corrente":  getAccountType = "checking"
        Case "Poupança":  getAccountType = "savings"
        Case "Salário":  getAccountType = "salary"
        Case Else:  getAccountType = "checking"
    End Select
End Function

Public Function createOrders(teamId As String)
    Dim orders As New Collection
    Dim orderNumbers As New Collection
    Dim externalIds As New Collection
    Dim iteration As Integer
    Dim startRow As Integer
    Dim currentRow As Integer
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    Dim returnMessage As String
    Dim warningMessage As String
    Dim errorMessage As String
    Dim anySent As Boolean
    
    anySent = False
    returnMessage = ""
    warningMessage = ""
    errorMessage = ""
    
    dict.Add "teamId", teamId
    iteration = 0
    startRow = TableFormat.HeaderRow() + 1
    
    For Each obj In SheetParser.dict
        Dim amount As Long
        Dim taxId As String
        Dim name As String
        Dim bankCode As String
        Dim branchCode As String
        Dim accountNumber As String
        Dim accountType As String
        Dim tags() As String
        Dim description As String
        Dim order As Dictionary
        
        iteration = iteration + 1
        currentRow = TableFormat.HeaderRow() + iteration
        
        If obj("Valor") = "" Then
            MsgBox "Por favor, não deixe linhas em branco entre as ordens de transferência", , "Erro"
            Unload SendOrderForm
            End
        End If
        amount = getAmountLong(obj("Valor"))
        taxId = Trim(obj("CPF/CNPJ"))
        name = Trim(obj("Nome"))
        bankCode = Trim(obj("Código do Banco/ISPB"))
        branchCode = Trim(obj("Agência"))
        accountNumber = Trim(obj("Conta"))
        accountType = getAccountType(obj("Tipo de Conta"))
        tags = Split(obj("Tags"), ",")
        description = obj("Descrição")
        externalId = obj("externalId")
        
        calculatedExternalId = calculateExternalId(amount, name, taxId, bankCode, branchCode, accountNumber)
        
        If calculatedExternalId = externalId Then
            warningMessage = "Aviso: Pedidos já enviados hoje não foram reenviados!" + Chr(10) + Chr(10)
        Else
            Set order = TeamGateway.order(amount, taxId, name, bankCode, branchCode, accountNumber, accountType, tags, description)
            orders.Add order
            orderNumbers.Add iteration
            externalIds.Add calculatedExternalId
        End If
        
        If (iteration Mod 100) = 0 Or (currentRow >= ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row) Then
            If orderNumbers.Count = 0 Then
                Set orders = Nothing
                Set orderNumbers = Nothing
                Set externalIds = Nothing
                GoTo nextIteration
            End If
            dict.Add "orders", orders
            
            payload = JsonConverter.ConvertToJson(dict)
            Set resp = StarkBankApi.postRequest("/v1/team/order", payload, New Dictionary)
            anySent = True
            
            dict.Remove "orders"
            
            If resp.Status = 200 Then
                createOrders = resp.json()("message")
                returnMessage = returnMessage + rowsMessage(startRow, currentRow) + resp.json()("message") + Chr(10)
                Dim j As Integer
                For j = 1 To externalIds.Count
                    ActiveSheet.Cells(TableFormat.HeaderRow() + orderNumbers.Item(j), 10).Value = externalIds.Item(j)
                Next j
                
            ElseIf resp.error().Exists("errors") Then
                Dim errors As Collection: Set errors = resp.error()("errors")
                Dim error As Dictionary
                Dim errorDescription As String
                
                For Each error In errors
                    errorDescription = Utils.correctErrorLine(error("message"), startRow)
                    errorMessage = errorMessage + errorDescription + Chr(10)
                Next
                
            Else
                errorDescription = Utils.correctErrorLine(resp.error()("message"), startRow)
                errorMessage = errorMessage + errorDescription + Chr(10)
        
            End If
nextIteration:
            startRow = currentRow + 1
            Set orders = Nothing
            Set orderNumbers = Nothing
            Set externalIds = Nothing
        End If
        
    Next
    If anySent Then
        MsgBox warningMessage + Chr(10) + returnMessage + Chr(10) + errorMessage
    Else
        MsgBox "Todos os pedidos listados já foram enviados"
    End If
End Function

Public Function order(amount As Long, taxId As String, name As String, bankCode As String, branchCode As String, accountNumber As String, accountType As String, tags() As String, description As String) As Dictionary
    Dim dict As New Dictionary
    
    dict.Add "amount", amount
    dict.Add "taxId", taxId
    dict.Add "name", name
    dict.Add "bankCode", bankCode
    dict.Add "branchCode", branchCode
    dict.Add "accountNumber", accountNumber
    dict.Add "accountType", accountType
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

Public Function calculateExternalId(amount As Long, name As String, taxId As String, bankCode As String, branchCode As String, accountNumber As String)
    calculateExternalId = bankCode + branchCode + accountNumber + name + taxId + CStr(amount)
End Function

Public Function rowsMessage(startRow As Integer, currentRow As Integer)
    rowsMessage = "Linhas " + CStr(startRow) + " a " + CStr(currentRow) + ": "
End Function