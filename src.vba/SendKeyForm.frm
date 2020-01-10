Private Sub BrowseFile_Click()
    Dim myFile As String
    myFile = Application.GetOpenFilename(Title:="Por favor, selecione a sua chave pública")
    If CStr(myFile) <> "False" Then
        Me.FileBox.Value = myFile
    End If
End Sub

Private Sub RequestToken_Click()
    Set response = DigitalSignature.mailToken()
    
    If response("error").Count <> 0 Then
        MsgBox response("error")("message"), , "Erro"
        Exit Sub
    End If
    MsgBox response("success")("message"), , "Sucesso"
End Sub

Private Sub SendPublic_Click()
    
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
    
    Set response = DigitalSignature.sendPublicKey(SessionGateway.getWorkspaceId(), Me.TokenBox.Value, Me.FileBox.Value)
    
    If response("error").Count <> 0 Then
        MsgBox response("error")("message"), , "Erro"
        Exit Sub
    End If
    MsgBox response("success")("message"), , "Sucesso"
    Unload Me
End Sub
