
Public Sub getStatement(beforeInput As String, afterInput As String)
    On Error Resume Next
    
    Worksheets("Extrato").Activate
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim cursor As String
    Dim transact As Dictionary
    Dim transactions As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    Dim respJson As Dictionary: Set respJson = New Dictionary
    Dim sign As Integer
    
    Dim transactionCreated As String
    Dim transactionId As String
    Dim transactionFee As Double
    
    Dim searchedLists() As String
    ReDim Preserve searchedLists(0)
    searchedLists(0) = ""
    
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
        If respJson.Count = 0 Then
            Exit Sub
        End If
        
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If
            
        Set transactions = respJson("transactions")
        
        Debug.Print "transactions", transactions.Count()
        
        For Each transact In transactions
        
            Dim initialRow As Integer
            Dim created As String: created = transact("created")
            Dim path As String:  path = transact("path")
            
            Dim tags As Collection: Set tags = transact("tags")
            Dim tagsStr As String: tagsStr = CollectionToString(tags, ",")
            Dim splitPath() As String: splitPath = Split(path, "/")
            
            transactionCreated = Utils.ISODATEZ(created)
            transactionId = transact("id")
            transactionFee = CDbl(transact("fee")) / 100
            
            conditionTeam = (splitPath(0) = "team")
            conditionTransferRequest = (splitPath(0) = "transfer-request")
            
            If (conditionTeam) And DetailedCheckBox.Value = True And (Not isChargeBack(splitPath)) Then
                If (Not Utils.IsInArray(path, searchedLists)) Then
                    initialRow = row
                    row = getOrdersInTransaction(path, transactionCreated, transactionId, transactionFee, row)
                    
                    ReDim Preserve searchedLists(UBound(searchedLists) + 1)
                    searchedLists(UBound(searchedLists)) = path
                End If
                
            ElseIf conditionTransferRequest And DetailedCheckBox.Value = True And (Not isChargeBack(splitPath)) Then
                initialRow = row
                row = getTransfersInTransaction(path, transactionCreated, transactionId, transactionFee, row)
                
            Else
                sign = transactionSign(transact("flow"))
                ActiveSheet.Cells(row, 1).Value = transactionCreated
                ActiveSheet.Cells(row, 2).Value = CDbl(transact("amount")) / 100 * sign
                ActiveSheet.Cells(row, 3).Value = transact("description")
                ActiveSheet.Cells(row, 4).Value = transactionId
                ActiveSheet.Cells(row, 5).Value = transactionFee
                ActiveSheet.Cells(row, 6).Value = CollectionToString(tags, ",")
                
                row = row + 1
            End If
        
        Next
    
    Loop While cursor <> ""
    
End Sub

Private Function getOrdersInTransaction(path As String, transactionCreated As String, transactionId As String, transactionFee As Double, row As Integer) As Integer
    Dim orders As Collection
    Dim cursor As String
    Dim orderTags As Collection
    Dim order As Object
    Dim orderTagsStr As String
    Dim orderDescription As String
    Dim teamId As String
    Dim listId As String
    Dim initialRow As Integer
    Dim getOrderParam As Dictionary: Set getOrderParam = New Dictionary
    Debug.Print path
    splitPath = Split(path, "/")
    teamId = splitPath(1)
    listId = splitPath(3)
    
    getOrderParam.Add "teamId", teamId
    getOrderParam.Add "listId", listId
    
    Do
        Set orderRespJson = TransferGateway.getOrders(cursor, getOrderParam)
        
        cursor = ""
        If orderRespJson("cursor") <> "" Then
            cursor = orderRespJson("cursor")
        End If
            
        Set orders = orderRespJson("orders")
        
        numberOfOrders = orders.Count()
        orderFee = transactionFee / numberOfOrders
        
        For Each order In orders
            If order("status") <> "disapproved" Then
                Set orderTags = order("tags")
                ActiveSheet.Cells(row, 1).Value = transactionCreated
                ActiveSheet.Cells(row, 2).Value = order("amount") / 100 * (-1)
                ActiveSheet.Cells(row, 3).Value = createDescription(order("name"), order("taxId"))
                ActiveSheet.Cells(row, 4).Value = transactionId
                ActiveSheet.Cells(row, 5).Value = orderFee
                ActiveSheet.Cells(row, 6).Value = CollectionToString(orderTags, ",")
                
                row = row + 1
            End If
        Next
    Loop While cursor <> ""
    
    getOrdersInTransaction = row
End Function

Private Function getTransfersInTransaction(path As String, transactionCreated As String, transactionId As String, transactionFee As Double, row As Integer) As Integer
    Dim transfers As Collection
    Dim cursor As String
    Dim transfer As Object
    Dim transferDescription As String
    Dim requestId As String
    Dim initialRow As Integer
    Dim transferFee As Double
    Dim transferTags As Collection
    Dim getTransferParam As Dictionary: Set getTransferParam = New Dictionary
    Dim splitPath() As String
    
    splitPath = Split(path, "/")
    requestId = splitPath(1)
    
    getTransferParam.Add "requestId", requestId
    
    Do
        Set transferRespJson = TransferGateway.getTransfers(cursor, getTransferParam)
        
        cursor = ""
        If transferRespJson("cursor") <> "" Then
            cursor = transferRespJson("cursor")
        End If
            
        Set transfers = transferRespJson("transfers")
        
        numberOfTransfers = transfers.Count()
        transferFee = transactionFee / numberOfTransfers
        
        For Each transfer In transfers
            Set transferTags = transfer("tags")
            transferTagsStr = CollectionToString(transferTags, ",")
            transferDescription = createDescription(transfer("name"), transfer("taxId"))
            
            ActiveSheet.Cells(row, 1).Value = transactionCreated
            ActiveSheet.Cells(row, 2).Value = transfer("amount") / 100 * (-1)
            ActiveSheet.Cells(row, 3).Value = transferDescription
            ActiveSheet.Cells(row, 4).Value = transactionId
            ActiveSheet.Cells(row, 5).Value = transferFee
            ActiveSheet.Cells(row, 6).Value = transferTagsStr
            
            row = row + 1
            initialRow = row
        Next
    Loop While cursor <> ""
    
    getTransfersInTransaction = row
End Function



Private Function createDescription(name As String, taxId As String) As String
    createDescription = "Transferência para " & name & ". CPF/CNPJ: " & taxId & "."
End Function

Private Function transactionSign(flow As String) As Integer
    transactionSign = 1
    
    If flow = "out" Then
        transactionSign = -1
    End If
End Function

Private Function transferSign(isChargeBack As Boolean) As Integer
    transferSign = -1
    
    If isChargeBack Then
        transferSign = 1
    End If
End Function

Private Function isChargeBack(splitPath() As String) As Boolean
    isChargeBack = False
    
    splitLen = UBound(splitPath, 1) - LBound(splitPath, 1) + 1
    
    If splitLen > 2 Then
        If splitPath(UBound(splitPath, 1)) = "chargeback" Then
            isChargeBack = True
        End If
    End If
End Function
