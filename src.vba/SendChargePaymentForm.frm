Private Sub UserForm_Initialize()
    Me.PathBox.text = InputLogGateway.getPath()
End Sub

Private Sub BrowseButton_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave privada")
    If CStr(myFile) <> "False" Then
        Me.PathBox.Value = myFile
    End If
End Sub

Private Sub ConfirmButton_Click()
    'On Error Resume Next
    Dim myFile As String: myFile = PathBox.Value
    
    Dim privkeyStr As String, textLine As String
    Dim response As Dictionary
    
    Call InputLogGateway.savePath(myFile)
    Call Utils.applyStandardLayout("E")
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Linha Digitável ou Código de Barras"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "CPF/CNPJ do Beneficiário"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Data de Agendamento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Descrição"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Tags"
    
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
    Dim dict As New Dictionary
    Dim payments As Collection
    
    Set payments = ChargePaymentGateway.getChargePaymentsFromSheet()
    
    dict.Add "payments", payments
    
    payload = JsonConverter.ConvertToJson(dict)
    
    '--------------- Sign body -----------------
    On Error GoTo eh
    Dim pk As PrivateKey: Set pk = New PrivateKey
    pk.fromPem (privkeyStr)
    
    On Error Resume Next
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(payload, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    '--------------- Create transfers -----------------
    Dim respMessage As String
    respMessage = ChargePaymentGateway.createPayments(payload, signature64)
    
    Unload Me
    Exit Sub
eh:
    MsgBox "Por favor, selecione uma chave privada válida!", vbCritical, "Falha de assinatura"
End Sub