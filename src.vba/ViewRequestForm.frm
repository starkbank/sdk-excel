
Private Sub UserForm_Initialize()
    Dim cursor As String
    Dim centers As Collection
    
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
    
    Do
        Set respJson = getCostCenters(cursor, New Dictionary)
    
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If
            
        Set centers = respJson("centers")
        
        For Each center In centers
            Me.CenterBox.AddItem center("name") + " (id = " + center("id") + ")"
        Next
    
    Loop While cursor <> ""
    
End Sub
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

Private Sub SearchButton_Click()
    ' On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim id As String
    Dim amount As Double
    Dim tags As Collection
    Dim cursor As String
    Dim paymentRequest As Object
    Dim requests As Collection
    Dim row As Long
    Dim logRow As Long
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    Dim rng As Range
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    'Table layout
    Utils.applyStandardLayout ("M")
    ActiveSheet.Hyperlinks.Delete
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":M" & Rows.Count).ClearContents
    
    'Date parameters
    If after <> "--" Then
        optionalParam.Add "after", after
    End If
    If before <> "--" Then
        optionalParam.Add "before", before
    End If
    
    'Cost Center parameters
    Dim centerInfo As String: centerInfo = CenterBox.Value
    With CreateObject("VBScript.RegExp")
        .Pattern = "\= ([^)]+)\)"
        .Global = True
        For Each M In .Execute(centerInfo)
            centerId = M.submatches(0)
        Next
    End With
    optionalParam.Add "centerId", centerId
    
    'Type parameters
    Dim paymentType As String: paymentType = ""
    If OptionButtonTypeTransfer.Value Then
        paymentType = paymentType + "transfer"
    End If
    If OptionButtonTypeBoleto.Value Then
        paymentType = paymentType + "boleto-payment"
    End If
    If OptionButtonTypeUtility.Value Then
        paymentType = paymentType + "utility-payment"
    End If
    If OptionButtonTypeTax.Value Then
        paymentType = paymentType + "tax-payment"
    End If
    If OptionButtonTypeQrcode.Value Then
        paymentType = paymentType + "brcode-payment"
    End If
    
    If paymentType <> "" Then
        optionalParam.Add "type", paymentType
    End If
    
    'Status parameters
    Dim RequestStatus As String: RequestStatus = ""
    If OptionButtonStatusPending.Value Then
        RequestStatus = RequestStatus + "pending,approved"
    End If
    If OptionButtonStatusScheduled.Value Then
        RequestStatus = RequestStatus + "scheduled,processing,success,failed"
    End If
    If OptionButtonStatusDenied.Value Then
        RequestStatus = RequestStatus + "denied,canceled"
    End If
    
    If RequestStatus <> "" Then
        optionalParam.Add "status", RequestStatus
    End If
    
    ViewRequestHeaderInitialize (paymentType)
    
    
    row = TableFormat.HeaderRow() + 1
    
    Do
        Set respJson = getPaymentRequests(cursor, optionalParam)
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If
        
        Set requests = respJson("requests")
        
        
        For Each paymentRequest In requests
            Dim requestDate As String: requestDate = paymentRequest("created")
            Dim paymentRequestType As String: paymentRequestType = paymentRequest("type")
            Dim paymentRequestStatus As String: paymentRequestStatus = paymentRequest("status")
            Dim actions As Collection: Set actions = paymentRequest("actions")
            Dim paymentRequestedBy As String: paymentRequestedBy = actions(1)("name")
            Dim payment As Dictionary: Set payment = paymentRequest("payment")
            
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(requestDate)
            ActiveSheet.Cells(row, 2).Value = getRequestTypeInPt(paymentRequestType)
            ActiveSheet.Cells(row, 3).Value = paymentRequest("description")
            ActiveSheet.Cells(row, 4).Value = paymentRequest("amount") / 100
            ActiveSheet.Cells(row, 5).Value = paymentRequestedBy
            ActiveSheet.Cells(row, 6).Value = getRequestStatusInPt(paymentRequestStatus)
            ActiveSheet.Cells(row, 7).Value = payment("id")
            ActiveSheet.Cells(row, 8).Value = CollectionToString(paymentRequest("tags"))
            
            Select Case paymentType
                Case "transfer":
                    ActiveSheet.Cells(row, 9).Value = payment("name")
                    ActiveSheet.Cells(row, 10).Value = payment("taxId")
                    ActiveSheet.Cells(row, 11).Value = payment("bankCode")
                    ActiveSheet.Cells(row, 12).Value = payment("branchCode")
                    ActiveSheet.Cells(row, 13).Value = payment("accountNumber")
                Case "boleto-payment":
                    ActiveSheet.Cells(row, 9).Value = payment("taxId")
                    ActiveSheet.Cells(row, 10).Value = getBarcodeOrLine(payment)
                Case "utility-payment":
                    ActiveSheet.Cells(row, 9).Value = getBarcodeOrLine(payment)
                Case "tax-payment":
                    ActiveSheet.Cells(row, 9).Value = getBarcodeOrLine(payment)
                Case "brcode-payment":
                    ActiveSheet.Cells(row, 9).Value = payment("taxId")
                    ActiveSheet.Cells(row, 10).Value = payment("brcode")
                Case Else:
            End Select
            row = row + 1
        Next
        
    Loop While cursor <> ""
    
    Unload Me
End Sub

Sub ViewRequestHeaderInitialize(paymentType As String)

    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Data da Solicitação"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Tipo de Pagamento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Descrição"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Solicitado por"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Status"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "ID do pagamento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Tags"
    
    Select Case paymentType
        Case "transfer":
            ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Nome"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "CPF / CNPJ"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Código do Banco / ISPB"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Agência"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = "Conta"
        Case "boleto-payment":
            ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "CPF / CNPJ"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Linha Digitável / Código de Barras"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = ""
        Case "utility-payment":
            ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Linha Digitável / Código de Barras"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = ""
        Case "tax-payment":
            ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Linha Digitável / Código de Barras"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = ""
        Case "brcode-payment":
            ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "CPF / CNPJ"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Copia e Cola (BR Code)"
            ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = ""
            ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = ""
        Case Else:
    End Select
    Call FreezeHeader
    
End Sub

Function getBarcodeOrLine(payment As Dictionary)
    Dim barcodeOrLine As String
    barcodeOrLine = payment("line")
    If barcodeOrLine = "" Or barcodeOrLine = Null Then
        barcodeOrLine = payment("barCode")
    End If
    getBarcodeOrLine = barcodeOrLine
End Function

Function getBrcodeType(brcodeType As String)
    Select Case brcodeType
        Case "dynamic": getBrcodeType = "Dinâmico"
        Case "static": getBrcodeType = "Estático"
        Case Else: getBrcodeType = "Outro"
    End Select
End Function