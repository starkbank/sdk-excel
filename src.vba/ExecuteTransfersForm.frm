Private Sub UserForm_Initialize()
    Me.PathBox.text = InputLogGateway.getPath()
End Sub

Private Sub ConfirmButton_Click()
    Dim myFile As String: myFile = PathBox.value
    Dim externalId As String: externalId = ExternalIdBox.value
    Dim description As String: description = DescriptionBox.value
    
    Dim privateKey As String, textLine As String
    Dim response As Dictionary
    
    Call InputLogGateway.savePath(myFile)
    Call Utils.applyStandardLayout("G")
    
    'Headers definition
    ActiveSheet.Cells(9, 1).value = "Nome"
    ActiveSheet.Cells(9, 2).value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 3).value = "Valor"
    ActiveSheet.Cells(9, 4).value = "Código do Banco"
    ActiveSheet.Cells(9, 5).value = "Agência"
    ActiveSheet.Cells(9, 6).value = "Conta"
    ActiveSheet.Cells(9, 7).value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 7
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    '--------------- Read privateKey -----------------
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textLine
        privateKey = privateKey & textLine
    Loop
    
    Close #1
    
    '--------------- Create body -----------------
    Dim payload As String, tags() As String
    Dim dict As New Dictionary, transactionDict As New Dictionary
    Dim transfers As Collection
    
    Set transfers = TransferGateway.getTransfersFromSheet()
    
    transactionDict.Add "externalId", externalId
    transactionDict.Add "description", description
    transactionDict.Add "tags", tags
    
    dict.Add "transaction", transactionDict
    dict.Add "transfers", transfers
    
    payload = JsonConverter.ConvertToJson(dict)
    
    '--------------- Sign body -----------------
    Dim signatureResp As Dictionary
    Set signatureResp = SignatureGateway.signMessage(privateKey, payload)
    
    If signatureResp("error").Count <> 0 Then
        MsgBox response("error")("message"), , "Erro"
        Unload Me
        Exit Sub
    End If

    Dim signature As String: signature = signatureResp("success")("signature")
'    Debug.Print "signature:"
'    Debug.Print Signature
'    Debug.Print "********************************"
    
    '--------------- Create transfers -----------------
    Dim respMessage As String
    respMessage = TransferGateway.createTransfers(payload, signature)
    
    Unload Me
     
End Sub

Private Sub BrowseButton_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave privada")
    Me.PathBox.value = myFile
End Sub
