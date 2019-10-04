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

Private Sub DownloadButton_Click()
    On Error Resume Next
    Application.DisplayAlerts = False
    
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim cursor As String
    Dim transactions As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    Dim transactionCreated As String
    Dim transactionId As String
    Dim transactionFee As Double
    
    ActiveSheet.Cells.UnMerge
    Call Utils.applyStandardLayout("F")
    ActiveSheet.Range("A10:F" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Data"
    ActiveSheet.Cells(9, 2).Value = "Valor"
    ActiveSheet.Cells(9, 3).Value = "Descrição"
    ActiveSheet.Cells(9, 4).Value = "Id da Transação"
    ActiveSheet.Cells(9, 5).Value = "Tarifa"
    ActiveSheet.Cells(9, 6).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 6
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    'Optional parameters
    optionalParam.Add "after", after
    optionalParam.Add "before", before
    
    row = 10
    
    Do
        Set respJson = getTransaction(cursor, optionalParam)
        
        If respJson.Count() = 0 Then
            Unload Me
            Exit Sub
        End If
        
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If
            
        Set transactions = respJson("transactions")
        
        For Each Transaction In transactions
            Dim tags As Collection: Set tags = Transaction("tags")
            Dim tagsStr As String: tagsStr = CollectionToString(tags, ",")
            Dim created As String: created = Transaction("created")
            Dim initialRow As Integer
            transactionCreated = Utils.ISODATEZ(created)
            transactionId = Transaction("id")
            
            transactionFee = CDbl(Transaction("fee")) / 100
            
            conditionTeam = InStr(Transaction("path"), "team")
            conditionTransferRequest = InStr(Transaction("path"), "transfer-request")
            If (conditionTeam) And DetailedCheckBox.Value = True Then
                initialRow = row
                row = getOrdersInTransaction(transactionCreated, transactionId, row, transactionFee)
            ElseIf conditionTransferRequest And DetailedCheckBox.Value = True Then
                initialRow = row
                row = getTransfersInTransaction(transactionCreated, transactionId, row, transactionFee)
            Else
                Dim signal As Integer: signal = 1
                If Transaction("flow") = "out" Then
                    signal = -1
                End If
    
                ActiveSheet.Cells(row, 1).Value = transactionCreated
                ActiveSheet.Cells(row, 2).Value = Transaction("amount") / 100 * signal
                ActiveSheet.Cells(row, 3).Value = Transaction("description")
                ActiveSheet.Cells(row, 4).Value = transactionId
                ActiveSheet.Cells(row, 5).Value = transactionFee
                ActiveSheet.Cells(row, 6).Value = CollectionToString(tags, ",")
                
                row = row + 1
            End If
        Next
    
    Loop While cursor <> ""
    
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
    
End Sub

Private Function getOrdersInTransaction(transactionCreated As String, transactionId As String, row As Integer, transactionFee As Double) As Integer
    Dim transfers As Collection
    Dim transfer As Object
    Dim transferTags As Collection
    Dim transferTagsStr As String
    Dim transferDescription As String
    Dim initialRow As Integer
    Dim transferFee As Double
    Dim getTransferParam As Dictionary: Set getTransferParam = New Dictionary
    getTransferParam.Add "transactionId", transactionId
    Set transferRespJson = TransferGateway.getTransfers("", getTransferParam)
    
    Set transfers = transferRespJson("transfers")
    numberOfTransfers = transfers.Count()
    transferFee = transactionFee / numberOfTransfers
    For Each transfer In transfers
        Set transferTags = transfer("tags")
        transferTagsStr = CollectionToString(transferTags, ",")
        transferDescription = createTransferDescription(transfer("name"), transfer("taxId"))
        
        initialRow = row
        row = getOrdersInTransfer(transferTagsStr, transferDescription, transactionCreated, transactionId, row, transferFee)
    Next
    
    getOrdersInTransaction = row
End Function

Private Function getTransfersInTransaction(transactionCreated As String, transactionId As String, row As Integer, transactionFee As Double) As Integer
    Dim transfers As Collection
    Dim transfer As Object
    Dim transferTags As Collection
    Dim transferTagsStr As String
    Dim transferDescription As String
    Dim initialRow As Integer
    Dim transferFee As Double
    Dim getTransferParam As Dictionary: Set getTransferParam = New Dictionary
    getTransferParam.Add "transactionId", transactionId
    Set transferRespJson = TransferGateway.getTransfers("", getTransferParam)
    
    Set transfers = transferRespJson("transfers")
    numberOfTransfers = transfers.Count()
    transferFee = transactionFee / numberOfTransfers
    For Each transfer In transfers
        Set transferTags = transfer("tags")
        transferTagsStr = CollectionToString(transferTags, ",")
        transferDescription = createTransferDescription(transfer("name"), transfer("taxId"))
        
        ActiveSheet.Cells(row, 1).Value = transactionCreated
        ActiveSheet.Cells(row, 2).Value = transfer("amount") / 100 * (-1)
        ActiveSheet.Cells(row, 3).Value = transferDescription
        ActiveSheet.Cells(row, 4).Value = transactionId
        ActiveSheet.Cells(row, 5).Value = transferFee
        ActiveSheet.Cells(row, 6).Value = transferTagsStr
        
        row = row + 1
        initialRow = row
    Next
    
    getTransfersInTransaction = row
End Function

Private Function getOrdersInTransfer(transferTags As String, transferDescription As String, transactionCreated As String, transactionId As String, row As Integer, transferFee As Double) As Integer
    Dim orders As Collection
    Dim order As Object
    Dim orderTags As Collection
    
    Dim teamId As String
    Dim listId As String
    Dim transferId As String
    Dim orderFee As Double
    result = Split(transferTags, "/")
    teamId = result(1)
    listId = result(3)
    transferId = result(5)
    
    Dim getOrdersParam As Dictionary: Set getOrdersParam = New Dictionary
    getOrdersParam.Add "teamId", teamId
    getOrdersParam.Add "listId", listId
    getOrdersParam.Add "transferId", transferId
    
    Set orderRespJson = TeamGateway.getOrdersByTransfer("", getOrdersParam)
    
    Set orders = orderRespJson("orders")
    numberOfOrders = orders.Count()
    orderFee = transferFee / numberOfOrders
    For Each order In orders
        Set orderTags = order("tags")
        
        ActiveSheet.Cells(row, 1).Value = transactionCreated
        ActiveSheet.Cells(row, 2).Value = order("amount") / 100 * (-1)
        ActiveSheet.Cells(row, 3).Value = transferDescription
        ActiveSheet.Cells(row, 4).Value = transactionId
        ActiveSheet.Cells(row, 5).Value = orderFee
        ActiveSheet.Cells(row, 6).Value = CollectionToString(orderTags, ",")
        
        row = row + 1
    Next
    
    getOrdersInTransfer = row
End Function

Private Function createTransferDescription(name As String, taxId As String) As String
    createTransferDescription = "Transferência para " & name & ". CPF/CNPJ: " & taxId & "."
End Function