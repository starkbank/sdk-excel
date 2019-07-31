Private Sub AfterTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    AfterTextBox.Text = Utils.formatDateInUserForm(AfterTextBox.Text)
    reentry = False
End Sub

Private Sub BeforeTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    BeforeTextBox.Text = Utils.formatDateInUserForm(BeforeTextBox.Text)
    reentry = False
End Sub

Private Sub UserForm_Initialize()
    Me.StatusComboBox.AddItem "Todos"
    Me.StatusComboBox.AddItem "Sucesso"
    Me.StatusComboBox.AddItem "Processando"
    Me.StatusComboBox.AddItem "Falha"
    
    Me.StatusComboBox.value = "Todos"
End Sub

Private Sub SearchButton_Click()
    'On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.value
    Dim beforeInput As String: beforeInput = BeforeTextBox.value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)

    Dim transactionId As String: transactionId = TransactionIdBox.value
    Dim statusString As String: statusString = StatusComboBox.value
    Dim cursor As String
    Dim transfers As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    'Table layout
    Utils.applyStandardLayout ("J")
    Range("A10:J" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(9, 1).value = "Data de Criação"
    ActiveSheet.Cells(9, 2).value = "Valor"
    ActiveSheet.Cells(9, 3).value = "Status"
    ActiveSheet.Cells(9, 4).value = "Nome"
    ActiveSheet.Cells(9, 5).value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 6).value = "Código do Banco"
    ActiveSheet.Cells(9, 7).value = "Agência"
    ActiveSheet.Cells(9, 8).value = "Número de Conta"
    ActiveSheet.Cells(9, 9).value = "Id da Transação"
    ActiveSheet.Cells(9, 10).value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 10
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    '--------------- If transactionId has been defined ------------------
    If transactionId <> "" Then
        optionalParam.Add "transactionId", transactionId
    End If
    
    '--------------- If status and dates have been defined ------------------
    Debug.Print "after: "
    Debug.Print after
    Dim Status As String: Status = TransferGateway.getStatus(statusString)
    If Status <> "all" And Status <> "" Then
        optionalParam.Add "status", Status
    End If
    If after <> "--" Then
        optionalParam.Add "after", after
    End If
    If before <> "--" Then
        optionalParam.Add "before", before
    End If

    row = 10

    Do
        Set respJson = getTransfers(cursor, optionalParam)

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set transfers = respJson("transfers")

        For Each transfer In transfers

            Dim created As String: created = transfer("created")
            ActiveSheet.Cells(row, 1).value = Utils.ISODATEZ(created)

            ActiveSheet.Cells(row, 2).value = transfer("amount") / 100

            Dim transferStatus As String: transferStatus = transfer("status")
            ActiveSheet.Cells(row, 3).value = TransferGateway.getStatusInPt(transferStatus)
            ActiveSheet.Cells(row, 4).value = transfer("name")
            ActiveSheet.Cells(row, 5).value = transfer("taxId")
            ActiveSheet.Cells(row, 6).value = transfer("bankCode")
            ActiveSheet.Cells(row, 7).value = transfer("branchCode")
            ActiveSheet.Cells(row, 8).value = transfer("accountNumber")
            ActiveSheet.Cells(row, 9).value = transfer("transactionId")

            Dim tags As Collection: Set tags = transfer("tags")
            ActiveSheet.Cells(row, 10).value = CollectionToString(tags, ",")

            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub