Dim AmountBoxoldValue As String

Private Sub SendInternalTransferForm_Initialize()
    PathBox.text = InputLogGateway.getPath()
    AmountBox = "R$ 0,00"
    AmountBox.SelStart = 6
End Sub

Private Sub BrowseButton_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave privada")
    If CStr(myFile) <> "False" Then
        Me.PathBox.Value = myFile
    End If
End Sub

Private Sub AmountBox_Change()
    AmountBox = Utils.formatCurrencyInUserForm(AmountBox)
End Sub

Private Sub ConfirmButton_Click()
    'On Error Resume Next
    Dim myFile As String: myFile = PathBox.Value
    Dim amount As Long: amount = Utils.IntegerFrom(Utils.clearNonNumeric(AmountBox.Value))
    Dim receiverId As String: receiverId = WorkspaceBox.Value
    Dim externalId As String: externalId = ExternalIdBox.Value
    Dim description As String: description = DescriptionBox.Value
    Dim tags() As String: tags = Split(TagsBox.Value, ",")
    
    Dim privkeyStr As String, textLine As String
    Dim response As Dictionary
    
    Call InputLogGateway.savePath(myFile)
    Call Utils.applyStandardLayout("G")
    
    Call FreezeHeader
    
    '----------- Sign in again -----------
    Dim password As String: password = PasswordBox.Value
    Dim workspace As String: workspace = SessionGateway.getWorkspace()
    Dim email As String: email = SessionGateway.getEmail()
    Set response = AuthGateway.createNewSession(workspace, email, password)
    
    If response("error").Count <> 0 Then
        MsgBox "Senha incorreta!", , "Erro"
        Exit Sub
    End If
        
    Dim accessToken As String: accessToken = response("success")("accessToken")
    Call SessionGateway.saveAccessToken(accessToken)
    
    '----------- Validate mandatory inputs -----------
    If externalId = vbNullString Then
        MsgBox "Por favor, adicione um identificador único", , "Erro"
        Exit Sub
    End If
    
    If myFile = vbNullString Then
        MsgBox "Nenhum arquivo selecionado", vbExclamation, "Erro"
        Exit Sub
    End If
    
    '--------------- Read privateKey -----------------
    If dir(myFile) <> vbNullString Then
        Open myFile For Input As #1
        Do Until EOF(1)
            Line Input #1, textLine
            privkeyStr = privkeyStr & textLine & vbLf
        Loop
        
        Close #1
    Else
        MsgBox "Arquivo não encontrado", vbExclamation
        Exit Sub
    End If
    
    '--------------- Create body -----------------
    Dim payload As String
    Dim dict As New Dictionary, transactionDict As New Dictionary
    
    transactionDict.Add "amount", amount
    transactionDict.Add "receiverId", receiverId
    transactionDict.Add "externalId", externalId
    transactionDict.Add "description", description
    transactionDict.Add "tags", tags
    
    dict.Add "transaction", transactionDict
    
    payload = JsonConverter.ConvertToJson(dict)
    
    '--------------- Sign body -----------------
    Dim pk As PrivateKey: Set pk = New PrivateKey
    pk.fromPem (privkeyStr)
    
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(payload, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    '--------------- Create transfers -----------------
    Dim respJson As Dictionary
    Set respJson = BankGateway.postTransaction(payload, signature64)
    
    If respJson.Exists("transaction") Then
        Unload Me
    End If
     
End Sub

Private Sub ExternalIdBox_Change()

End Sub

Private Sub Label8_Click()

End Sub