
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
    Utils.applyStandardLayout ("V")
    ActiveSheet.Hyperlinks.Delete
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":V" & Rows.count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = " Data de Emissão"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Status"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Valor de Emissão"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Desconto"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Multa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Juros"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Data de Crédito"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Vencimento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Linha Digitável"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = "Id do Boleto"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 14).Value = "Tarifa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 15).Value = "Tags"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 16).Value = "Link PDF"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 17).Value = "Logradouro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 18).Value = "Complemento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 19).Value = "Bairro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 20).Value = "Cidade"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 21).Value = "Estado"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 22).Value = "CEP"
    
    With ActiveWindow
        .SplitColumn = 6
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

            Dim id As String: id = charge("id")
            Dim amount As Double: amount = charge("amount") / 100
            Dim fee As Double: fee = charge("fee") / 100
            Dim issueDate As String: issueDate = charge("issueDate")
            Dim chargeStatus As String: chargeStatus = charge("status")
            Dim dueDate As String: dueDate = charge("dueDate")
            Dim tags As Collection: Set tags = charge("tags")
            
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(issueDate)
            ActiveSheet.Cells(row, 2).Value = charge("name")
            ActiveSheet.Cells(row, 3).Value = charge("taxId")
            ActiveSheet.Cells(row, 4).Value = ChargeGateway.getStatusInPt(chargeStatus)
            ActiveSheet.Cells(row, 5).Value = amount
            
            ActiveSheet.Cells(row, 11).Value = Utils.ISODATEZ(dueDate)
            ActiveSheet.Cells(row, 12).Value = charge("line")
            ActiveSheet.Cells(row, 13).Value = id
            ActiveSheet.Cells(row, 14).Value = fee
            ActiveSheet.Cells(row, 15).Value = CollectionToString(tags, ",")
            
            ActiveSheet.Cells(row, 16).Value = "PDF"
            Set rng = ActiveSheet.Range("P" + CStr(row))
            rng.Parent.Hyperlinks.Add Anchor:=rng, address:=StarkBankApi.baseUrl() + "/v1/charge/" + charge("id") + "/pdf", SubAddress:="", TextToDisplay:="PDF"

            If DetailedCheckBox.Value = True Then
                ActiveSheet.Cells(row, 17).Value = charge("streetLine1")
                ActiveSheet.Cells(row, 18).Value = charge("streetLine2")
                ActiveSheet.Cells(row, 19).Value = charge("district")
                ActiveSheet.Cells(row, 20).Value = charge("city")
                ActiveSheet.Cells(row, 21).Value = charge("stateCode")
                ActiveSheet.Cells(row, 22).Value = charge("zipCode")
                
                If chargeStatus = "paid" Then
                    Dim nominalAmount As Double
                    Dim fine As Double
                    Dim interest As Double
                    Dim discount As Double
                    Dim deltaAmount As Double
                    Dim paidDate As String
                    Dim logs As Collection
                    Dim paidLog As Dictionary
                    Dim createdLog As Dictionary
                    
                    Set logs = ChargeGateway.getEventLog(id, "register,paid", New Dictionary)("logs")
                    Set createdLog = logs(2)
                    Set paidLog = logs(1)
                    
                    paidDate = paidLog("created")
                    ActiveSheet.Cells(row, 10).Value = Utils.ISODATEZ(paidDate)
                    
                    nominalAmount = createdLog("charge")("amount") / 100
                    deltaAmount = amount - nominalAmount
                    
                    ActiveSheet.Cells(row, 6).Value = nominalAmount
                    If deltaAmount < 0 Then
                        discount = deltaAmount
                        ActiveSheet.Cells(row, 7).Value = discount
                        ActiveSheet.Cells(row, 7).Font.Color = RGB(180, 0, 0)
                    End If
                    If deltaAmount > 0 Then
                        fine = charge("fine") / 100 * nominalAmount
                        interest = amount - fine - nominalAmount
                        ActiveSheet.Cells(row, 8).Value = fine
                        ActiveSheet.Cells(row, 9).Value = interest
                        ActiveSheet.Cells(row, 8).Font.Color = RGB(0, 140, 0)
                        ActiveSheet.Cells(row, 9).Font.Color = RGB(0, 140, 0)
                    End If
                End If
            End If
            
            row = row + 1
        Next

    Loop While cursor <> ""
    
    Unload Me
     
End Sub