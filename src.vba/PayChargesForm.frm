Private Sub UserForm_Initialize()
    Me.PathBox.Text = InputLogGateway.getPath()
End Sub

Private Sub BrowseButton_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave privada")
    Me.PathBox.Value = myFile
End Sub

Private Sub ConfirmButton_Click()
    On Error Resume Next
    Dim myFile As String: myFile = PathBox.Value
    
    Dim privkeyStr As String, textLine As String
    Dim response As Dictionary
    
    Call InputLogGateway.savePath(myFile)
    Call Utils.applyStandardLayout("D")
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Linha Digitável ou Código de Barras"
    ActiveSheet.Cells(9, 2).Value = "Data de Agendamento"
    ActiveSheet.Cells(9, 3).Value = "Descrição"
    ActiveSheet.Cells(9, 4).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 4
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
    
    '--------------- Read privateKey -----------------
    Open myFile For Input As #1
    Do Until EOF(1)
        Line Input #1, textLine
        privkeyStr = privkeyStr & textLine
    Loop
    
    Close #1
    
    '--------------- Create body -----------------
    Dim payload As String, tags() As String
    Dim dict As New Dictionary
    Dim payments As Collection
    
    Set payments = ChargePaymentGateway.getChargePaymentsFromSheet()
    
    dict.Add "payments", payments
    
    payload = JsonConverter.ConvertToJson(dict)
    
    '--------------- Sign body -----------------
    Dim pk As privateKey: Set pk = New privateKey
    pk.fromPem (privkeyStr)
    
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(payload, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    '--------------- Create transfers -----------------
    Dim respMessage As String
    respMessage = ChargePaymentGateway.createPayments(payload, signature64)
    
    Unload Me
     
End Sub