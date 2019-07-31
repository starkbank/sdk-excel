Private Sub UserForm_Initialize()
    Me.StatusBox.AddItem "Todos"
    Me.StatusBox.AddItem "Pagos"
    Me.StatusBox.AddItem "Pendentes de Registro"
    Me.StatusBox.AddItem "Registrados"
    Me.StatusBox.AddItem "Vencidos"
    Me.StatusBox.AddItem "Cancelados"
    
    Me.StatusBox.value = "Todos"
End Sub

Private Sub SearchButton_Click()
    On Error Resume Next
    Dim statusString As String: statusString = StatusBox.value
    Dim cursor As String
    Dim charges As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    'Table layout
    Utils.applyStandardLayout ("I")
    Range("A10:I" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(9, 1).value = " Data de Emissão"
    ActiveSheet.Cells(9, 2).value = "Valor"
    ActiveSheet.Cells(9, 3).value = "Vencimento"
    ActiveSheet.Cells(9, 4).value = "Status"
    ActiveSheet.Cells(9, 5).value = "Nome"
    ActiveSheet.Cells(9, 6).value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 7).value = "Linha Digitável"
    ActiveSheet.Cells(9, 8).value = "Id da Transação"
    ActiveSheet.Cells(9, 9).value = "Tags"
    
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
            ActiveSheet.Cells(row, 1).value = Utils.ISODATEZ(issueDate)

            ActiveSheet.Cells(row, 2).value = charge("amount") / 100

            Dim dueDate As String: dueDate = charge("dueDate")
            ActiveSheet.Cells(row, 3).value = Utils.ISODATEZ(dueDate)

            Dim chargeStatus As String: chargeStatus = charge("status")
            ActiveSheet.Cells(row, 4).value = ChargeGateway.getStatusInPt(chargeStatus)
            ActiveSheet.Cells(row, 5).value = charge("name")
            ActiveSheet.Cells(row, 6).value = charge("taxId")
            ActiveSheet.Cells(row, 7).value = charge("line")
            ActiveSheet.Cells(row, 8).value = charge("id")

            Dim tags As Collection: Set tags = charge("tags")
            ActiveSheet.Cells(row, 9).value = CollectionToString(tags, ",")

            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub