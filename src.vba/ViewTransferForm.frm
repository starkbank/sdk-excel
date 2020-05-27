Private Sub AfterTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    AfterTextBox.text = Utils.formatDateInUserForm(AfterTextBox.text)
    reentry = False
End Sub

Private Sub BeforeTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    BeforeTextBox.text = Utils.formatDateInUserForm(BeforeTextBox.text)
    reentry = False
End Sub

Private Sub UserForm_Initialize()
    Me.StatusComboBox.AddItem "Todos"
    Me.StatusComboBox.AddItem "Sucesso"
    Me.StatusComboBox.AddItem "Processando"
    Me.StatusComboBox.AddItem "Falha"
    
    Me.StatusComboBox.Value = "Todos"
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
End Sub

Private Sub SearchButton_Click()
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)

    Dim transactionId As String: transactionId = TransactionIdBox.Value
    Dim statusString As String: statusString = StatusComboBox.Value
    Dim cursor As String
    Dim transfers As Collection
    Dim transferLogs As Collection
    Dim logsFailedByTransfer As Dictionary: Set logsFailedByTransfer = New Dictionary
    Dim transfer As Object
    Dim transferLog As Object
    Dim row As Long
    Dim initialRow As Long
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    If beforeInput = "" Then
        beforeInput = Format(Date, "dd/mm/yyyy")
        If afterInput = "" Then
            afterInput = Format(DateAdd("d", -30, Date), "dd/mm/yyyy")
        End If
    ElseIf afterInput = "" Then
        afterInput = "01/01/2018"
    End If
    
    after = Utils.DateToSendingFormat(afterInput)
    before = Utils.DateToSendingFormat(beforeInput)
    
    If DateDiff("d", after, before) > 30 Then
        If MsgBox("O período selecionado é superior a 30 dias. A operação pode demorar. Continuar?", vbOKCancel) = vbCancel Then
            Exit Sub
        End If
    End If
    
    'Table layout
    Utils.applyStandardLayout ("K")
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":K" & Rows.Count).ClearContents
    
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
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Detalhamento de falha"
    
    Call FreezeHeader
    
    '--------------- If transactionId has been defined ------------------
    If transactionId <> "" Then
        optionalParam.Add "transactionIds", transactionId
    End If
    
    '--------------- If status and dates have been defined ------------------
    Dim Status As String: Status = TransferGateway.getStatus(statusString)
    If Status <> "all" And Status <> "" Then
        optionalParam.Add "status", Status
    End If
    If after <> "--" And transactionId = "" Then
        optionalParam.Add "after", after
    End If
    If before <> "--" And transactionId = "" Then
        optionalParam.Add "before", before
    End If

    row = TableFormat.HeaderRow() + 1

    Do
        logRow = row
        Set respJson = getTransfers(cursor, optionalParam)
        If respJson.Exists("error") Then
            Unload Me
            Exit Sub
        End If

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set transfers = respJson("transfers")

        For Each transfer In transfers
            transferId = transfer("id")
            Dim created As String: created = transfer("created")
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(created)
            ActiveSheet.Cells(row, 2).Value = transferId
            ActiveSheet.Cells(row, 3).Value = transfer("amount") / 100

            Dim transferStatus As String: transferStatus = transfer("status")
            ActiveSheet.Cells(row, 4).Value = TransferGateway.getStatusInPt(transferStatus)
            ActiveSheet.Cells(row, 5).Value = transfer("name")
            ActiveSheet.Cells(row, 6).Value = transfer("taxId")
            ActiveSheet.Cells(row, 7).Value = transfer("bankCode")
            ActiveSheet.Cells(row, 8).Value = transfer("branchCode")
            ActiveSheet.Cells(row, 9).Value = transfer("accountNumber")
            ActiveSheet.Cells(row, 10).Value = Utils.CollectionToString(Utils.correctTransferTags(transfer("transactionIds")), ",")
            
            If DetailedCheckBox.Value = True And transferStatus = "failed" Then
                logsFailedByTransfer.Add transferId, New Dictionary
            End If
            row = row + 1
        Next
        
        If DetailedCheckBox.Value = True Then
            
            Dim logsParam As Dictionary
            Dim keys As String
            Dim sep As String
            Set logsParam = New Dictionary
            logsParam.Add "types", "failed"
            keys = ""
            sep = ""
            For Each transferId In logsFailedByTransfer.keys()
                keys = keys + sep + transferId
                sep = ","
            Next
            logsParam.Add "transferIds", keys
            Set respJson = getTransferLogs("", logsParam)
            If respJson.Exists("error") Then
                MsgBox "Erro ao obter dados detalhados de transferências com falha!"
                Exit Sub
            End If
    
            Set transferLogs = respJson("logs")
            
            For Each transferLog In transferLogs
                Set logsFailedByTransfer(transferLog("transferId")) = transferLog
            Next
            
            For Each transfer In transfers
                If transfer("status") = "failed" Then
                    Dim errors As String: errors = ""
                    Dim errorMessage As Variant
                    sep = ""
                    For Each errorMessage In logsFailedByTransfer(transfer("id"))("errors")
                        errors = errors + sep + errorMessage
                        sep = ","
                    Next
                    ActiveSheet.Cells(logRow, 11).Value = errors
                End If
                logRow = logRow + 1
            Next
            
            Set logsFailedByTransfer = New Dictionary
        End If
        
    Loop While cursor <> ""
       
    Unload Me
End Sub