Public Sub signIn()
    SignInForm.Show
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
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "Credentials" And ws.name <> "InputLog" Then
            ws.Cells(2, 1).Value = ""
            ws.Cells(3, 1).Value = ""
            ws.Cells(4, 1).Value = ""
            ws.Cells(5, 1).Value = ""
            ws.Cells(6, 1).Value = ""
        End If
    Next
    
    clearAll
    Application.ScreenUpdating = True
    MsgBox response("success")("message"), , "Sucesso"
End Sub

Public Sub clearAll()
    For Each ws In ThisWorkbook.Worksheets
        ws.Cells.UnMerge
        ws.Range("A10:Z" & Rows.Count).ClearContents
    Next
End Sub

Public Sub searchStatement()
    SearchForm.Show
End Sub

Public Sub sendOrders()
    On Error Resume Next
    SendOrderForm.Show
End Sub

Public Sub searchCharges()
    ChargeForm.Show
End Sub

Public Sub searchTransfers()
    TransferForm.Show
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
    ActiveSheet.Cells(9, 1).Value = "Id do Cliente"
    ActiveSheet.Cells(9, 2).Value = "Nome"
    ActiveSheet.Cells(9, 3).Value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 4).Value = "E-mail"
    ActiveSheet.Cells(9, 5).Value = "Telefone"
    ActiveSheet.Cells(9, 6).Value = "Logradouro"
    ActiveSheet.Cells(9, 7).Value = "Complemento"
    ActiveSheet.Cells(9, 8).Value = "Bairro"
    ActiveSheet.Cells(9, 9).Value = "Cidade"
    ActiveSheet.Cells(9, 10).Value = "Estado"
    ActiveSheet.Cells(9, 11).Value = "CEP"
    ActiveSheet.Cells(9, 12).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 12
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    row = 10

    Do
        Set respJson = getCustomers(cursor, optionalParam)

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
    On Error Resume Next
    
    Dim charges As Collection
    Dim respMessage As String
    
    Call Utils.applyStandardLayout("L")
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Id do Cliente"
    ActiveSheet.Cells(9, 2).Value = "Valor"
    ActiveSheet.Cells(9, 3).Value = "Data de Vencimento"
    ActiveSheet.Cells(9, 4).Value = "Multa"
    ActiveSheet.Cells(9, 5).Value = "Juros ao Mês"
    ActiveSheet.Cells(9, 6).Value = "Dias para Baixa Automática"
    ActiveSheet.Cells(9, 7).Value = "Descrição 1"
    ActiveSheet.Cells(9, 8).Value = "Valor 1"
    ActiveSheet.Cells(9, 9).Value = "Descrição 2"
    ActiveSheet.Cells(9, 10).Value = "Valor 2"
    ActiveSheet.Cells(9, 11).Value = "Descrição 3"
    ActiveSheet.Cells(9, 12).Value = "Valor 3"
    
    With ActiveWindow
        .SplitColumn = 12
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True

    Set charges = ChargeGateway.getOrders()
    
    respMessage = ChargeGateway.createCharges(charges)
     
End Sub

Public Sub executeTransfers()
    ExecuteTransfersForm.Show
End Sub

Public Sub payCharges()
    PayChargesForm.Show
End Sub

Public Sub searchChargePayments()
    ChargePaymentForm.Show
End Sub