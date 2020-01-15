
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
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
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
    Dim rng As Range
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    'Table layout
    Utils.applyStandardLayout ("J")
    ActiveSheet.Hyperlinks.Delete
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":J" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = " Data de Emissão"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Vencimento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Status"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Linha Digitável"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Id do Boleto"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Tags"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Link PDF"
    
    With ActiveWindow
        .SplitColumn = 10
        .SplitRow = TableFormat.HeaderRow()
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
    
    row = TableFormat.HeaderRow() + 1

    Do
        Set respJson = getCharges(cursor, optionalParam)
        If respJson.Exists("error") Then
            Unload Me
            Exit Sub
        End If
        
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
            
            ActiveSheet.Cells(row, 10).Value = "Link"
            
            Set rng = ActiveSheet.Range("J" + CStr(row))
            rng.Parent.Hyperlinks.Add Anchor:=rng, address:=StarkBankApi.baseUrl() + "/v1/charge/" + charge("id") + "/pdf", SubAddress:="", TextToDisplay:="PDF"

            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub