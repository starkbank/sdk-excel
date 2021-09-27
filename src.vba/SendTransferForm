Private Sub UserForm_Initialize()
    Me.PathBox.text = InputLogGateway.getPath()
End Sub

Private Sub ConfirmButton_Click()
    Dim myFile As String: myFile = PathBox.Value
    Dim externalId As String: externalId = ExternalIdBox.Value
    Dim description As String: description = DescriptionBox.Value
    
    Dim privkeyStr As String, textLine As String
    Dim response As Dictionary
    
    Call InputLogGateway.savePath(myFile)
    Call Utils.applyStandardLayout("G")
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Código do Banco"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Agência"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Conta"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Tags"
    
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
    On Error GoTo eh
    Dim pk As PrivateKey: Set pk = New PrivateKey
    pk.fromPem (privkeyStr)
    
    On Error Resume Next
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(payload, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    '--------------- Create transfers -----------------
    Dim respJson As Dictionary
    Set respJson = TransferGateway.createTransfers(payload, signature64)
    
    Unload Me
    Exit Sub
eh:
    MsgBox "Por favor, selecione uma chave privada válida!", vbCritical, "Falha de assinatura"
End Sub

Private Sub BrowseButton_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave privada")
    If CStr(myFile) <> "False" Then
        Me.PathBox.Value = myFile
    End If
End Sub
