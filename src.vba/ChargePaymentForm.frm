
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
    Me.StatusBox.AddItem "Todos"
    Me.StatusBox.AddItem "Criados"
    Me.StatusBox.AddItem "Processando"
    Me.StatusBox.AddItem "Pagos"
    Me.StatusBox.AddItem "Falha"
    
    Me.StatusBox.Value = "Todos"
End Sub

Private Sub SearchButton_Click()
    On Error Resume Next

    Dim statusString As String: statusString = StatusBox.Value
    Dim cursor As String
    Dim payments As Collection
    Dim payment As Object
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    'Table layout
    Utils.applyStandardLayout ("G")
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":G" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Data de Criação"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Status"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Data de Agendamento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Linha Digitável"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Descrição"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 7
        .SplitRow = TableFormat.HeaderRow()
    End With
    ActiveWindow.FreezePanes = True
    
    'Optional parameters
    Dim Status As String: Status = ChargePaymentGateway.getStatus(statusString)
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
        Set respJson = ChargePaymentGateway.getChargePayments(cursor, optionalParam)

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set payments = respJson("payments")

        For Each payment In payments

            Dim created As String: created = payment("created")
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(created)
            
            ActiveSheet.Cells(row, 2).Value = payment("amount") / 100
            
            Dim paymentStatus As String: paymentStatus = payment("status")
            ActiveSheet.Cells(row, 3).Value = ChargePaymentGateway.getStatusInPt(paymentStatus)

            Dim scheduled As String: scheduled = payment("scheduled")
            ActiveSheet.Cells(row, 4).Value = Utils.ISODATEZ(scheduled)

            ActiveSheet.Cells(row, 5).Value = payment("line")
            ActiveSheet.Cells(row, 6).Value = payment("description")

            Dim tags As Collection: Set tags = payment("tags")
            ActiveSheet.Cells(row, 7).Value = CollectionToString(tags, ",")

            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub
