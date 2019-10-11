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
    Me.StatusBox.AddItem "Pagos"
    Me.StatusBox.AddItem "Pendentes de Registro"
    Me.StatusBox.AddItem "Registrados"
    Me.StatusBox.AddItem "Vencidos"
    Me.StatusBox.AddItem "Cancelados"
    
    Me.StatusBox.Value = "Todos"
End Sub

Private Sub SearchButton_Click()
    On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim statusString As String: statusString = StatusBox.Value
    Dim cursor As String
    Dim charges As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    'Table layout
    Utils.applyStandardLayout ("I")
    Range("A10:I" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = " Data de Emissão"
    ActiveSheet.Cells(9, 2).Value = "Valor"
    ActiveSheet.Cells(9, 3).Value = "Vencimento"
    ActiveSheet.Cells(9, 4).Value = "Status"
    ActiveSheet.Cells(9, 5).Value = "Nome"
    ActiveSheet.Cells(9, 6).Value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 7).Value = "Linha Digitável"
    ActiveSheet.Cells(9, 8).Value = "Id do Boleto"
    ActiveSheet.Cells(9, 9).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 9
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    'Optional parameters
    Dim Status As String: Status = ChargeGateway.getStatus(statusString)
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
        Set respJson = getCharges(cursor, optionalParam)

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set charges = respJson("charges")

        For Each charge In charges

            Dim issueDate As String: issueDate = charge("issueDate")
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(issueDate)

            ActiveSheet.Cells(row, 2).Value = charge("amount") / 100

            Dim dueDate As String: dueDate = charge("dueDate")
            ActiveSheet.Cells(row, 3).Value = Utils.ISODATEZ(dueDate)

            Dim chargeStatus As String: chargeStatus = charge("status")
            ActiveSheet.Cells(row, 4).Value = ChargeGateway.getStatusInPt(chargeStatus)
            ActiveSheet.Cells(row, 5).Value = charge("name")
            ActiveSheet.Cells(row, 6).Value = charge("taxId")
            ActiveSheet.Cells(row, 7).Value = charge("line")
            ActiveSheet.Cells(row, 8).Value = charge("id")

            Dim tags As Collection: Set tags = charge("tags")
            ActiveSheet.Cells(row, 9).Value = CollectionToString(tags, ",")

            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub