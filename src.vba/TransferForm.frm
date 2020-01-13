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
    
    Me.StatusComboBox.Value = "Todos"
End Sub

Private Sub SearchButton_Click()
    On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)

    Dim transactionId As String: transactionId = TransactionIdBox.Value
    Dim statusString As String: statusString = StatusComboBox.Value
    Dim cursor As String
    Dim transfers As Collection
    Dim transfer As Object
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    'Table layout
    Utils.applyStandardLayout ("J")
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":J" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Data de Criação"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Id da Transferência"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Status"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Código do Banco"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Agência"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Número de Conta"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Ids de Transação (Saída, Estorno)"
    
    With ActiveWindow
        .SplitColumn = 10
        .SplitRow = TableFormat.HeaderRow()
    End With
    ActiveWindow.FreezePanes = True
    
    '--------------- If transactionId has been defined ------------------
    If transactionId <> "" Then
        optionalParam.Add "transactionIds", transactionId
    End If
    
    '--------------- If status and dates have been defined ------------------
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

    row = TableFormat.HeaderRow() + 1

    Do
        Set respJson = getTransfers(cursor, optionalParam)

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set transfers = respJson("transfers")

        For Each transfer In transfers

            Dim created As String: created = transfer("created")
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(created)
            ActiveSheet.Cells(row, 2).Value = transfer("id")
            ActiveSheet.Cells(row, 3).Value = transfer("amount") / 100

            Dim transferStatus As String: transferStatus = transfer("status")
            ActiveSheet.Cells(row, 4).Value = TransferGateway.getStatusInPt(transferStatus)
            ActiveSheet.Cells(row, 5).Value = transfer("name")
            ActiveSheet.Cells(row, 6).Value = transfer("taxId")
            ActiveSheet.Cells(row, 7).Value = transfer("bankCode")
            ActiveSheet.Cells(row, 8).Value = transfer("branchCode")
            ActiveSheet.Cells(row, 9).Value = transfer("accountNumber")
            ActiveSheet.Cells(row, 10).Value = Utils.CollectionToString(Utils.correctTransferTags(transfer("transactionIds")), ",")
            
            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub