





Private Sub UserForm_Initialize()
    Me.PathBox.Text = InputLogGateway.getPath()
End Sub

Private Sub ConfirmButton_Click()
    On Error Resume Next
    Dim myFile As String: myFile = PathBox.Value
    Dim externalId As String: externalId = ExternalIdBox.Value
    Dim description As String: description = DescriptionBox.Value
    
    Dim privkeyStr As String, textLine As String
    Dim response As Dictionary
    
    Call InputLogGateway.savePath(myFile)
    Call Utils.applyStandardLayout("G")
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Nome"
    ActiveSheet.Cells(9, 2).Value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 3).Value = "Valor"
    ActiveSheet.Cells(9, 4).Value = "Código do Banco"
    ActiveSheet.Cells(9, 5).Value = "Agência"
    ActiveSheet.Cells(9, 6).Value = "Conta"
    ActiveSheet.Cells(9, 7).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 7
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    '----------- Sign in again -----------
    Dim password As String: password = PasswordBox.Value
    Dim workspace As String: workspace = SessionGateway.getWorkspace()
    Dim email As String: email = SessionGateway.getEmail()
    Set response = AuthGateway.createNewSession(workspace, email, password)
    
    If response("error").Count <> 0 Then
        MsgBox "Senha incorreta!", , "Erro"
        Unload Me
        Exit Sub
    End If
        
    Dim accessToken As String: accessToken = response("success")("accessToken")
    Call SessionGateway.saveAccessToken(accessToken)
    
    '----------- Validate mandatory inputs -----------
    If externalId = vbNullString Then
        MsgBox "Por favor, adicione um identificador único", , "Erro"
        Unload Me
        Exit Sub
    End If
    
    '--------------- Read privateKey -----------------
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textLine
        privkeyStr = privkeyStr & textLine & vbLf
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
    Dim pk As PrivateKey: Set pk = New PrivateKey
    pk.fromPem (privkeyStr)
    
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(payload, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    '--------------- Create transfers -----------------
    Dim respMessage As String
    respMessage = TransferGateway.createTransfers(payload, signature64)
    
    Unload Me
     
End Sub

Private Sub BrowseButton_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave privada")
    Me.PathBox.Value = myFile
End Sub
