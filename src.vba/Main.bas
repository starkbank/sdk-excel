Public Sub signIn()

SignInForm.Show

End Sub

Public Sub signOut()
Dim response As Dictionary

Set response = AuthGateway.deleteSession(SessionGateway.getAccessToken())

If response("error").Count <> 0 And response("error")("code") <> "invalidAccessToken" Then
    MsgBox response("error")("message"), , "Erro"
    Exit Sub
End If

Call SessionGateway.saveSession("", "", "", "", "")
For Each ws In ThisWorkbook.Worksheets
    If ws.name <> "Credentials" And ws.name <> "InputLog" Then
        ws.Cells(2, 1).value = ""
        ws.Cells(3, 1).value = ""
        ws.Cells(4, 1).value = ""
        ws.Cells(5, 1).value = ""
    End If
Next

MsgBox response("success")("message"), , "Sucesso"

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
    ActiveSheet.Cells(9, 1).value = "Id do Cliente"
    ActiveSheet.Cells(9, 2).value = "Nome"
    ActiveSheet.Cells(9, 3).value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 4).value = "E-mail"
    ActiveSheet.Cells(9, 5).value = "Telefone"
    ActiveSheet.Cells(9, 6).value = "Logradouro"
    ActiveSheet.Cells(9, 7).value = "Complemento"
    ActiveSheet.Cells(9, 8).value = "Bairro"
    ActiveSheet.Cells(9, 9).value = "Cidade"
    ActiveSheet.Cells(9, 10).value = "Estado"
    ActiveSheet.Cells(9, 11).value = "CEP"
    ActiveSheet.Cells(9, 12).value = "Tags"
    
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

            ActiveSheet.Cells(row, 1).value = customer("id")
            ActiveSheet.Cells(row, 2).value = customer("name")
            ActiveSheet.Cells(row, 3).value = customer("taxId")
            ActiveSheet.Cells(row, 4).value = customer("email")
            ActiveSheet.Cells(row, 5).value = customer("phone")
            
            Dim address As Dictionary: Set address = customer("address")
            
            ActiveSheet.Cells(row, 6).value = address("streetLine1")
            ActiveSheet.Cells(row, 7).value = address("streetLine2")
            ActiveSheet.Cells(row, 8).value = address("district")
            ActiveSheet.Cells(row, 9).value = address("city")
            ActiveSheet.Cells(row, 10).value = address("stateCode")
            ActiveSheet.Cells(row, 11).value = address("zipCode")

            Dim tags As Collection: Set tags = customer("tags")
            ActiveSheet.Cells(row, 12).value = CollectionToString(tags, ",")

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
    ActiveSheet.Cells(9, 1).value = "Id do Cliente"
    ActiveSheet.Cells(9, 2).value = "Valor"
    ActiveSheet.Cells(9, 3).value = "Data de Vencimento"
    ActiveSheet.Cells(9, 4).value = "Multa"
    ActiveSheet.Cells(9, 5).value = "Juros ao Mês"
    ActiveSheet.Cells(9, 6).value = "Dias para Baixa Automática"
    ActiveSheet.Cells(9, 7).value = "Descrição 1"
    ActiveSheet.Cells(9, 8).value = "Valor 1"
    ActiveSheet.Cells(9, 9).value = "Descrição 2"
    ActiveSheet.Cells(9, 10).value = "Valor 2"
    ActiveSheet.Cells(9, 11).value = "Descrição 3"
    ActiveSheet.Cells(9, 12).value = "Valor 3"
    
    With ActiveWindow
        .SplitColumn = 12
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True

    Set charges = ChargeGateway.getOrders()
    
    respMessage = ChargeGateway.createCharges(charges)
     
End Sub

