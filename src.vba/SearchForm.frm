
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
    'On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value

    Dim after As String
    Dim before As String
    
    Dim cursor As String
    Dim teamCursor As String
    Dim transact As Dictionary
    Dim transactions As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    Dim respJson As Dictionary: Set respJson = New Dictionary
    Dim sign As Integer
    Dim teams As Collection
    
    Dim transactionType As String
    Dim transactionCreated As String
    Dim transactionId As String
    Dim transactionFee As Double

    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    If beforeInput = "" Then
        beforeInput = Date
        If afterInput = "" Then
            afterInput = DateAdd("d", -30, Date)
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
    
    ActiveSheet.Cells.UnMerge
    Call Utils.applyStandardLayout("F")
    ActiveSheet.Range("A10:G" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Data"
    ActiveSheet.Cells(9, 2).Value = "Tipo de transação"
    ActiveSheet.Cells(9, 3).Value = "Valor"
    ActiveSheet.Cells(9, 4).Value = "Descrição"
    ActiveSheet.Cells(9, 5).Value = "Id da Transação"
    ActiveSheet.Cells(9, 6).Value = "Tarifa"
    ActiveSheet.Cells(9, 7).Value = "Tags"

    With ActiveWindow
        .SplitColumn = 7
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True

    'Optional parameters
    If after <> "--" Then
        optionalParam.Add "after", after
    End If
    If before <> "--" Then
        optionalParam.Add "before", before
    End If

    row = 10
    
    Set teams = getTeams("", New Dictionary)("teams")
    
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
            transactionType = getTransactionType(splitPath, teams)
            
            conditionTeam = (splitPath(0) = "team")
            conditionTransferRequest = (splitPath(0) = "transfer-request")
            If (conditionTeam Or conditionTransferRequest) And DetailedCheckBox.Value = True Then
                initialRow = row
                row = getTransfersInTransaction(path, transactionCreated, transactionId, transactionFee, transactionType, row)
                
            Else
                sign = transactionSign(transact("flow"))
                ActiveSheet.Cells(row, 1).Value = transactionCreated
                ActiveSheet.Cells(row, 2).Value = transactionType
                ActiveSheet.Cells(row, 3).Value = CDbl(transact("amount")) / 100 * sign
                If sign > 0 Then
                    ActiveSheet.Cells(row, 3).Font.Color = RGB(0, 140, 0)
                Else
                    ActiveSheet.Cells(row, 3).Font.Color = RGB(180, 0, 0)
                End If
                ActiveSheet.Cells(row, 4).Value = transact("description")
                ActiveSheet.Cells(row, 5).Value = transactionId
                ActiveSheet.Cells(row, 6).Value = transactionFee
                ActiveSheet.Cells(row, 7).Value = CollectionToString(tags, ",")
                
                row = row + 1
            End If

        Next

    Loop While cursor <> ""

    Unload Me

End Sub

Private Function getTransfersInTransaction(path As String, transactionCreated As String, transactionId As String, transactionFee As Double, transactionType As String, row As Integer) As Integer
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
    Dim sign As Integer
    Dim chargebackBool As Boolean

    sign = -1
    splitPath = Split(path, "/")
    requestId = splitPath(1)
    chargebackBool = False
    getTransferParam.Add "transactionIds", transactionId
    If isChargeback(splitPath) Then
        sign = 1
        getTransferParam.Add "status", "failed"
        chargebackBool = True
    End If

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
            If transferTags.Count() <> 0 Then
                Set transferTags = correctTransferTags(transferTags)
            End If
            transferTagsStr = CollectionToString(transferTags, ",")
            transferDescription = createDescription(transfer("name"), transfer("taxId"), chargebackBool)

            ActiveSheet.Cells(row, 1).Value = transactionCreated
            ActiveSheet.Cells(row, 2).Value = transactionType
            ActiveSheet.Cells(row, 3).Value = transfer("amount") / 100 * sign
            If sign > 0 Then
                ActiveSheet.Cells(row, 3).Font.Color = RGB(0, 140, 0)
            Else
                ActiveSheet.Cells(row, 3).Font.Color = RGB(180, 0, 0)
            End If
            ActiveSheet.Cells(row, 4).Value = transferDescription
            ActiveSheet.Cells(row, 5).Value = transactionId
            ActiveSheet.Cells(row, 6).Value = transferFee
            ActiveSheet.Cells(row, 7).Value = transferTagsStr

            row = row + 1
            initialRow = row
        Next
    Loop While cursor <> ""

    getTransfersInTransaction = row
End Function

Private Function getTransactionType(list() As String, ByRef teams As Collection)
    Dim transactionType As String
    Select Case list(0)
        Case "self"
            transactionType = "Transferência interna"
        Case "charge"
            transactionType = "Recebimento de boleto pago"
        Case "charge-payment"
            transactionType = "Pagamento de boleto"
        Case "transfer-request"
            transactionType = "Transf. sem aprovação"
        Case "team"
            Dim teamName As String
            Dim team As Dictionary
            teamName = ""
            For Each team In teams
                If team("id") = list(1) Then
                    teamName = team("name")
                End If
            Next
            transactionType = "Transf. com aprovação: Time " & teamName
        Case Else
            transactionType = "Outros"
    End Select
    
    If isChargeback(list) Then
        transactionType = "Estorno: " & transactionType
    End If
    getTransactionType = transactionType
End Function

Private Function createDescription(name As String, taxId As String, isChargeback As Boolean) As String
    createDescription = "Transferência para " & name & ". CPF/CNPJ: " & taxId & "."
    If isChargeback Then
        createDescription = "Estorno de saldo por falha de " + createDescription
    End If
End Function

Private Function transactionSign(flow As String) As Integer
    transactionSign = 1

    If flow = "out" Then
        transactionSign = -1
    End If
End Function

Private Function isChargeback(splitPath() As String) As Boolean
    isChargeback = False

    splitLen = UBound(splitPath, 1) - LBound(splitPath, 1) + 1

    If splitLen > 2 Then
        If splitPath(UBound(splitPath, 1)) = "chargeback" Then
            isChargeback = True
        End If
    End If
End Function
