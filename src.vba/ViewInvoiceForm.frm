
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
    Me.StatusBox.AddItem "Criados"
    Me.StatusBox.AddItem "Vencidos"
    Me.StatusBox.AddItem "Cancelados"
    Me.StatusBox.AddItem "Expirados"
    
    
    Me.StatusBox.Value = "Todos"
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
End Sub

Private Sub SearchButton_Click()
    ' On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim statusString As String: statusString = StatusBox.Value
    Dim cursor As String
    Dim invoices As Collection
    Dim invoiceLogs As Collection
    Dim row As Long
    Dim logRow As Long
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    Dim rng As Range
    
    Dim logsPaidByInvoice As Dictionary: Set logsPaidByInvoice = New Dictionary
    Dim logsRegisteredByInvoice As Dictionary: Set logsRegisteredByInvoice = New Dictionary
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    'Table layout
    Utils.applyStandardLayout ("P")
    ActiveSheet.Hyperlinks.Delete
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":P" & Rows.Count).ClearContents
    
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
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Vencimento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Pagável até"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Copia e Cola (BR Code)"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = "Id do Boleto"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 14).Value = "Tarifa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 15).Value = "Tags"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 16).Value = "Link PDF"
    
    Call FreezeHeader
    
    'Optional parameters
    Dim Status As String: Status = v2InvoiceGateway.getStatus(statusString)
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
        logRow = row
        Set respJson = getInvoices(cursor, optionalParam)
        If respJson.Exists("error") Then
            Unload Me
            Exit Sub
        End If
        
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set invoices = respJson("invoices")

        For Each invoice In invoices

            Dim id As String: id = invoice("id")
            Dim amount As Double: amount = invoice("amount") / 100
            Dim nominalAmount As Double: nominalAmount = invoice("nominalAmount") / 100
            Dim discountAmount As Double: discountAmount = invoice("discountAmount") / 100
            Dim fineAmount As Double: fineAmount = invoice("fineAmount") / 100
            Dim interestAmount As Double: interestAmount = invoice("interestAmount") / 100
            Dim fee As Double: fee = invoice("fee") / 100
            Dim issueDate As String: issueDate = invoice("created")
            Dim updated As String: updated = invoice("updated")
            Dim invoiceStatus As String: invoiceStatus = invoice("status")
            Dim dueDate As String: dueDate = invoice("due")
            Dim expiration As Long: expiration = invoice("expiration")
            Dim tags As Collection: Set tags = invoice("tags")
            
            Dim expirationDate As Date: expirationDate = DateAdd("s", expiration, Utils.DatefromIsoString(dueDate))
            
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(issueDate)
            ActiveSheet.Cells(row, 2).Value = invoice("name")
            ActiveSheet.Cells(row, 3).Value = invoice("taxId")
            ActiveSheet.Cells(row, 4).Value = v2InvoiceGateway.getStatusInPt(invoiceStatus)
            ActiveSheet.Cells(row, 5).Value = amount
            ActiveSheet.Cells(row, 6).Value = nominalAmount
            ActiveSheet.Cells(row, 7).Value = discountAmount
            ActiveSheet.Cells(row, 8).Value = fineAmount
            ActiveSheet.Cells(row, 9).Value = interestAmount
            ActiveSheet.Cells(row, 10).Value = Utils.ISODATEZ(dueDate)
            ActiveSheet.Cells(row, 11).Value = Utils.ISODATEZ(Format(expirationDate, "yyyy-mm-ddThh:mm:ss"))
            ActiveSheet.Cells(row, 12).Value = invoice("brcode")
            ActiveSheet.Cells(row, 13).Value = id
            ActiveSheet.Cells(row, 14).Value = fee
            ActiveSheet.Cells(row, 15).Value = CollectionToString(tags, ",")
            
            ActiveSheet.Cells(row, 16).Value = "PDF"
            Set rng = ActiveSheet.Range("P" + CStr(row))
            rng.Parent.Hyperlinks.Add Anchor:=rng, address:=invoice("pdf"), SubAddress:="", TextToDisplay:="PDF"
            
            row = row + 1
        Next
        
    Loop While cursor <> ""
    
    Unload Me
End Sub