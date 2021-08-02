
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
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
    
End Sub

Private Sub DownloadButton_Click()

    If Not isSignedin Then
        MsgBox "Acesso negado. Faça login novamente.", , "Erro"
        Exit Sub
    End If
    
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
    
    Dim transactionType As String
    Dim transactionCreated As String
    Dim transactionId As String
    Dim transactionFee As Double
    Dim balance As Double

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
    
    ActiveSheet.Cells.UnMerge
    Call Utils.applyStandardLayout("I")
    ActiveSheet.Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":I" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Data"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Username"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Número da Conta (Workspace ID)"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Tipo de transação"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Descrição"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Id da Transação"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Tarifa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Tags"

    Call FreezeHeader

    'Optional parameters
    If after <> "--" Then
        optionalParam.Add "after", after
    End If
    If before <> "--" Then
        optionalParam.Add "before", before
    End If
    
    row = TableFormat.HeaderRow() + 1
    
    Dim workspaceList As Collection: Set workspaceList = SheetParser.dict("Listar Contas")
    If workspaceList.Count() = 0 Then
        MsgBox "É necessário listar as contas antes de baixar o extrato!", vbExclamation
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo eh:
    
    For Each workspace In workspaceList
        Dim workspaceId As String: workspaceId = workspace("Número da Conta (Workspace ID)")
        Dim workspaceUsername As String: workspaceUsername = workspace("Username")
        
        Call postSessionV1(True, CStr(workspaceId))
        
        Do
            Set respJson = V2BankGateway.getTransaction(cursor, optionalParam)
            If respJson.Exists("error") Then
                Unload Me
                Exit Sub
            End If
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
                Dim path As String:  path = transact("source")
    
                Dim tags As Collection: Set tags = transact("tags")
                Dim tagsStr As String: tagsStr = CollectionToString(tags, ",")
                Dim splitPath() As String: splitPath = Split(path, "/")
                
                transactionCreated = created
                transactionId = transact("id")
                transactionFee = CDbl(transact("fee")) / 100
                transactionType = getTransactionType(splitPath, tags)
                
                sign = transactionSign(transact("flow"))
                ActiveSheet.Cells(row, 1).Value = transactionCreated
                ActiveSheet.Cells(row, 2).Value = workspaceUsername
                ActiveSheet.Cells(row, 3).Value = workspaceId
                ActiveSheet.Cells(row, 4).Value = transactionType
                ActiveSheet.Cells(row, 5).Value = CDbl(transact("amount")) / 100 * sign
                If sign > 0 Then
                    ActiveSheet.Cells(row, 5).Font.Color = RGB(0, 140, 120)
                Else
                    ActiveSheet.Cells(row, 5).Font.Color = RGB(180, 0, 150)
                End If
                
                ActiveSheet.Cells(row, 6).Value = transact("description")
                ActiveSheet.Cells(row, 7).Value = transactionId
                ActiveSheet.Cells(row, 8).Value = transactionFee
                ActiveSheet.Cells(row, 9).Value = CollectionToString(tags, ",")
                
                row = row + 1
    
            Next
    
        Loop While cursor <> ""
    Next
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 2).End(xlUp).row
    Range("A10:I" & lastRow).Sort key1:=Range("A10:A" & lastRow), order1:=xlDescending, header:=xlNo
    Unload Me
eh:
    
End Sub

Private Function getTransactionType(list() As String, ByRef tags As Collection)
    Dim transactionType As String
    Select Case list(0)
        Case "self"
            transactionType = "Transferência interna"
        Case "charge"
            transactionType = "Recebimento de boleto pago"
        Case "boleto"
            transactionType = "Recebimento de boleto pago"
        Case "invoice"
            transactionType = "Recebimento de cobrança Pix"
        Case "deposit"
            transactionType = "Recebimento de depósito Pix"
        Case "charge-payment"
            transactionType = "Pag. de boleto"
        Case "boleto-payment"
            transactionType = "Pag. de boleto"
        Case "brcode-payment"
            transactionType = "Pag. de QR Code"
        Case "bar-code-payment"
            transactionType = "Pag. de imposto/concessionária"
        Case "utility-payment"
            transactionType = "Pag. de concessionária com cód. de barras"
        Case "darf-payment"
            transactionType = "Pag. de DARF sem cód. de barras"
        Case "tax-payment"
            transactionType = "Pag. de imposto com cód. de barras"
        Case "transfer-request"
            transactionType = "Transf. sem aprovação"
        Case "transfer"
            Dim tag As Variant
            transactionType = "Transf. sem aprovação"
            For Each tag In tags
                If InStr(1, tag, "payment-request/") Then
                    transactionType = "Transf. com aprovação"
                End If
            Next
        Case "team"
            transactionType = "Transf. com aprovação"
        Case Else
            transactionType = "Outros"
    End Select
    
    If isChargeback(list) Then
        transactionType = "Estorno: " & transactionType
    End If
    getTransactionType = transactionType
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

