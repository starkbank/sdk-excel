Public Sub signIn()
    If OpensslRoutine Then
        SignInForm.Show
    End If
End Sub

Public Sub signOut()
    On Error Resume Next
    Dim response As Dictionary
    
    message1 = "Você quer mesmo encerrar a sessão? "
    message2 = "Dados que não foram salvos serão apagados."
    confirmationMessage = message1 + message2
    signOutAnswer = MsgBox(confirmationMessage, vbQuestion + vbYesNo, "Confirmação de encerramento")
    
    If signOutAnswer = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Set response = AuthGateway.deleteSession(SessionGateway.getAccessToken())
    
    If response("error").Count <> 0 And response("error")("code") <> "invalidAccessToken" Then
        MsgBox response("error")("message"), , "Erro"
        Exit Sub
    End If
    
    Call SessionGateway.saveSession("", "", "", "", "", "")
    For Each WS In ThisWorkbook.Worksheets
        If WS.name <> "Credentials" And WS.name <> "InputLog" And WS.name <> "Aux" Then
        
            WS.Cells(2, 1).Value = ""
            WS.Cells(3, 1).Value = ""
            WS.Cells(4, 1).Value = ""
            WS.Cells(5, 1).Value = ""
            WS.Cells(6, 1).Value = ""
            WS.Cells(7, 1).Value = ""
        End If
    Next
    clearDates
    clearAll
    Application.ScreenUpdating = True
    MsgBox response("success")("message"), , "Sucesso"
End Sub

Public Sub clearAll()
    For Each WS In ThisWorkbook.Worksheets
        If WS.name <> "Principal" And WS.name <> "Aux" Then
            WS.Cells.UnMerge
            WS.Range("A10:Z" & Rows.Count).ClearContents
            If WS.name = "InputLog" Then
                WS.Range("B:B").ClearContents
            End If
        End If
    Next
End Sub

Public Sub openHelp()
    With ViewHelpForm
        .MultiPage1.Value = 0
        .Show
    End With
End Sub

Public Sub openHelpDigSign()
    With ViewHelpForm
        .MultiPage1.Value = 2
        .Show
    End With
End Sub

Public Sub searchStatement()
    ViewStatementForm.Show
End Sub

Public Sub sendOrders()
    On Error Resume Next
    SendOrderForm.Show
End Sub

Public Sub searchCharges()
    ViewChargeForm.Show
End Sub

Public Sub searchInvoices()
    ViewInvoiceForm.Show
End Sub

Public Sub searchChargeEvents()
    ViewChargeEventsForm.Show
End Sub

Public Sub searchTransfers()
    ViewTransferForm.Show
End Sub

Public Sub keyGeneration()
    GenerateKeyForm.Show
End Sub

Public Sub keyUpload()
    SendKeyForm.Show
End Sub

Public Sub searchCustomers()
    On Error Resume Next
    
    Dim cursor As String
    Dim customers As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    'Table layout
    Utils.applyStandardLayout ("L")
    Range("A10:L" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Id do Cliente"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "E-mail"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Telefone"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Logradouro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Complemento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Bairro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Cidade"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Estado"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "CEP"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 12
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    row = 10

    Do
        Set respJson = ChargeGateway.getCustomers(cursor, optionalParam)

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set customers = respJson("customers")

        For Each customer In customers

            ActiveSheet.Cells(row, 1).Value = customer("id")
            ActiveSheet.Cells(row, 2).Value = customer("name")
            ActiveSheet.Cells(row, 3).Value = customer("taxId")
            ActiveSheet.Cells(row, 4).Value = customer("email")
            ActiveSheet.Cells(row, 5).Value = customer("phone")
            
            Dim address As Dictionary: Set address = customer("address")
            
            ActiveSheet.Cells(row, 6).Value = address("streetLine1")
            ActiveSheet.Cells(row, 7).Value = address("streetLine2")
            ActiveSheet.Cells(row, 8).Value = address("district")
            ActiveSheet.Cells(row, 9).Value = address("city")
            ActiveSheet.Cells(row, 10).Value = address("stateCode")
            ActiveSheet.Cells(row, 11).Value = address("zipCode")

            Dim tags As Collection: Set tags = customer("tags")
            ActiveSheet.Cells(row, 12).Value = CollectionToString(tags, ",")

            row = row + 1
        Next

    Loop While cursor <> ""
     
End Sub

Public Sub createCharges()
    Dim charges As Collection
    Dim resp As response
    Dim initRow As Long
    Dim midRow As Long
    Dim lastRow As Long
    Dim respMessage As String
    
    Call Utils.applyStandardLayout("L")
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Id do Cliente"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Data de Vencimento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Multa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Juros ao Mês"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Dias para Baixa Automática"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Descrição 1"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Valor 1"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Descrição 2"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Valor 2"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Descrição 3"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Valor 3"
    
    With ActiveWindow
        .SplitColumn = 12
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    initRow = 10
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    midRow = initRow - 1
    Do
        midRow = IIf(midRow + 100 >= lastRow, lastRow, midRow + 100)
        Set charges = ChargeGateway.getOrders(initRow, midRow)
        Set resp = ChargeGateway.createCharges(charges)
        
        If resp.Status = 200 Then
            respMessage = resp.json()("message")
            MsgBox "Linhas " + CStr(initRow) + " a " + CStr(midRow) + ": " + resp.json()("message"), , "Sucesso"
            
        ElseIf resp.error().Exists("errors") Then
            Dim errors As Collection: Set errors = resp.error()("errors")
            Dim error As Dictionary
            Dim errorList As String
            Dim errorDescription As String
            
            For Each error In errors
                errorDescription = Utils.correctErrorLine(error("message"), CLng(initRow) - 1)
                errorList = errorList & errorDescription & Chr(10)
            Next
            
            Dim messageBox As String
            messageBox = resp.error()("message") & Chr(10) & Chr(10) & errorList
            MsgBox messageBox, , "Erro"
            End
        Else
            MsgBox resp.error()("message"), , "Erro"
            
        End If
        
        initRow = initRow + 100
    Loop Until (midRow >= lastRow)
     
End Sub

Public Sub createInvoices()

    If Not isSignedin Then
        MsgBox "Acesso negado. Faça login novamente.", , "Erro"
        Exit Sub
    End If
    
    Dim invoices As Collection
    Dim resp As response
    Dim initRow As Long
    Dim midRow As Long
    Dim lastRow As Long
    Dim respMessage As String
    
    Call Utils.applyStandardLayout("M")
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Nome do Cliente"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "CPF/CNPJ do Cliente"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Data de Vencimento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Multa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Juros ao Mês"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Expiração em Horas"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Descrição 1"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Valor 1"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Descrição 2"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Valor 2"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Descrição 3"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = "Valor 3"
    
    With ActiveWindow
        .SplitColumn = 13
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    initRow = 10
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    midRow = initRow - 1
    Do
        midRow = IIf(midRow + 100 >= lastRow, lastRow, midRow + 100)
        Set invoices = v2InvoiceGateway.getOrders(initRow, midRow)
        Set resp = v2InvoiceGateway.createInvoices(invoices)
        
        If resp.Status = 200 Then
            respMessage = resp.json()("message")
            MsgBox "Linhas " + CStr(initRow) + " a " + CStr(midRow) + ": " + resp.json()("message"), , "Sucesso"
            
        ElseIf resp.errors().Exists("errors") Then
            Dim errors As Collection: Set errors = resp.errors()("errors")
            Dim error As Dictionary
            Dim errorList As String
            Dim errorDescription As String
            
            For Each error In errors
                errorDescription = Utils.correctErrorLine(error("message"), CLng(initRow))
                errorList = errorList & errorDescription & Chr(10)
            Next
            
            Dim messageBox As String
            messageBox = errorList
            MsgBox messageBox, , "Erro"
            End
        Else
            MsgBox resp.error()("message"), , "Erro"
            
        End If
        
        initRow = initRow + 100
    Loop Until (midRow >= lastRow)
     
End Sub

Public Sub sendCustomers()
    Dim customers As Collection
    Dim resp As response
    Dim initRow As Long
    Dim midRow As Long
    Dim lastRow As Long
    Dim respMessage As String
    Dim errorDescription As String
    
    Call Utils.applyStandardLayout("K")
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "E-mail"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Telefone"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Logradouro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Complemento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Bairro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Cidade"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Estado"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "CEP"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 11
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True

    initRow = 10
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    midRow = initRow - 1
    Do
        midRow = IIf(midRow + 100 >= lastRow, lastRow, midRow + 100)
        Set customers = ChargeGateway.getCustomerOrders(initRow, midRow)
        Set resp = ChargeGateway.createCustomers(customers)
        
        If resp.Status = 200 Then
            respMessage = resp.json()("message")
            MsgBox "Linhas " + CStr(initRow) + " a " + CStr(midRow) + ": " + resp.json()("message"), , "Sucesso"
            
        ElseIf resp.error().Exists("errors") Then
            Dim errors As Collection: Set errors = resp.error()("errors")
            Dim error As Dictionary
            Dim errorList As String
            
            For Each error In errors
                errorDescription = Utils.correctErrorLine(error("message"), initRow - 1)
                errorList = errorList & errorDescription & Chr(10)
            Next
            
            Dim messageBox As String
            messageBox = resp.error()("message") & Chr(10) & Chr(10) & errorList
            MsgBox messageBox, , "Erro"
            End
        Else
            Dim errorMessage As String: errorMessage = resp.error()("message")
            errorDescription = Utils.correctErrorLine(errorMessage, initRow - 1)
            MsgBox errorDescription, , "Erro"
            
        End If
        
        initRow = initRow + 100
    Loop Until (midRow >= lastRow)
    
End Sub

Public Sub executeTransfers()
    SendTransferForm.Show
End Sub

Public Sub executeInternalTransfers()
    SendInternalTransferForm.Show
End Sub

Public Sub payCharges()
    SendChargePaymentForm.Show
End Sub

Public Sub searchChargePayments()
    ViewChargePaymentForm.Show
End Sub

Public Sub searchPaymentRequests()
    If Not isSignedin() Then
        MsgBox "Acesso negado. Faça login novamente.", , "Erro"
        Exit Sub
    End If
    ViewRequestForm.Show
End Sub