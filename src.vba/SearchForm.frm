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
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
    
End Sub

Private Sub DownloadButton_Click()
    
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    Dim afterInput As String: afterInput = AfterTextBox.Value
    
    Call Statement.getStatement(beforeInput, afterInput)
    
    Unload Me
End Sub
'
'Private Function getOrdersInTransaction(path As String, transactionCreated As String, transactionId As String, transactionFee As Double, row As Integer) As Integer
'    Dim orders As Collection
'    Dim cursor As String
'    Dim orderTags As Collection
'    Dim order As Object
'    Dim orderTagsStr As String
'    Dim orderDescription As String
'    Dim teamId As String
'    Dim listId As String
'    Dim initialRow As Integer
'    Dim getOrderParam As Dictionary: Set getOrderParam = New Dictionary
'    Debug.Print path
'    splitPath = Split(path, "/")
'    teamId = splitPath(1)
'    listId = splitPath(3)
'
'    getOrderParam.Add "teamId", teamId
'    getOrderParam.Add "listId", listId
'
'    Do
'        Set orderRespJson = TransferGateway.getOrders(cursor, getOrderParam)
'
'        cursor = ""
'        If orderRespJson("cursor") <> "" Then
'            cursor = orderRespJson("cursor")
'        End If
'
'        Set orders = orderRespJson("orders")
'
'        numberOfOrders = orders.Count()
'        orderFee = transactionFee / numberOfOrders
'
'        For Each order In orders
'            If order("status") <> "disapproved" Then
'                Set orderTags = order("tags")
'                ActiveSheet.Cells(row, 1).Value = transactionCreated
'                ActiveSheet.Cells(row, 2).Value = order("amount") / 100 * (-1)
'                ActiveSheet.Cells(row, 3).Value = createDescription(order("name"), order("taxId"))
'                ActiveSheet.Cells(row, 4).Value = transactionId
'                ActiveSheet.Cells(row, 5).Value = orderFee
'                ActiveSheet.Cells(row, 6).Value = CollectionToString(orderTags, ",")
'
'                row = row + 1
'            End If
'        Next
'    Loop While cursor <> ""
'
'    getOrdersInTransaction = row
'End Function
'
'Private Function getTransfersInTransaction(path As String, transactionCreated As String, transactionId As String, transactionFee As Double, row As Integer) As Integer
'    Dim transfers As Collection
'    Dim cursor As String
'    Dim transfer As Object
'    Dim transferDescription As String
'    Dim requestId As String
'    Dim initialRow As Integer
'    Dim transferFee As Double
'    Dim transferTags As Collection
'    Dim getTransferParam As Dictionary: Set getTransferParam = New Dictionary
'    Dim splitPath() As String
'
'    splitPath = Split(path, "/")
'    requestId = splitPath(1)
'
'    getTransferParam.Add "requestId", requestId
'
'    Do
'        Set transferRespJson = TransferGateway.getTransfers(cursor, getTransferParam)
'
'        cursor = ""
'        If transferRespJson("cursor") <> "" Then
'            cursor = transferRespJson("cursor")
'        End If
'
'        Set transfers = transferRespJson("transfers")
'
'        numberOfTransfers = transfers.Count()
'        transferFee = transactionFee / numberOfTransfers
'
'        For Each transfer In transfers
'            Set transferTags = transfer("tags")
'            transferTagsStr = CollectionToString(transferTags, ",")
'            transferDescription = createDescription(transfer("name"), transfer("taxId"))
'
'            ActiveSheet.Cells(row, 1).Value = transactionCreated
'            ActiveSheet.Cells(row, 2).Value = transfer("amount") / 100 * (-1)
'            ActiveSheet.Cells(row, 3).Value = transferDescription
'            ActiveSheet.Cells(row, 4).Value = transactionId
'            ActiveSheet.Cells(row, 5).Value = transferFee
'            ActiveSheet.Cells(row, 6).Value = transferTagsStr
'
'            row = row + 1
'            initialRow = row
'        Next
'    Loop While cursor <> ""
'
'    getTransfersInTransaction = row
'End Function
'
'Private Function createDescription(name As String, taxId As String) As String
'    createDescription = "Transferência para " & name & ". CPF/CNPJ: " & taxId & "."
'End Function
'
'Private Function transactionSign(flow As String) As Integer
'    transactionSign = 1
'
'    If flow = "out" Then
'        transactionSign = -1
'    End If
'End Function
'
'Private Function transferSign(isChargeBack As Boolean) As Integer
'    transferSign = -1
'
'    If isChargeBack Then
'        transferSign = 1
'    End If
'End Function
'
'Private Function isChargeBack(splitPath() As String) As Boolean
'    isChargeBack = False
'
'    splitLen = UBound(splitPath, 1) - LBound(splitPath, 1) + 1
'
'    If splitLen > 2 Then
'        If splitPath(UBound(splitPath, 1)) = "chargeback" Then
'            isChargeBack = True
'        End If
'    End If
'End Function