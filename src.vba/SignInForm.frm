
Private Sub UserForm_Initialize()
    Me.EnvironmentBox.AddItem "Produção"
    Me.EnvironmentBox.AddItem "Sandbox"
    
    Me.EmailBox.Value = SessionGateway.getEmail()
    Me.WorkspaceBox.Value = SessionGateway.getWorkspace()
    Me.EnvironmentBox.Value = SessionGateway.getEnvironmentString()
    
End Sub

Private Sub SendButton_Click()
    Debug.Print "SendButton_Click started"
    ' On Error Resume Next
    Dim workspace As String: workspace = WorkspaceBox.Value
    Dim email As String: email = EmailBox.Value
    Dim password As String: password = PasswordBox.Value
    Dim envString As String: envString = EnvironmentBox.Value
    Dim accessToken As String
    Dim memberName As String
    Dim response As Dictionary
    Dim workspaceId As String
    Dim role As String
    Debug.Print "Variables declared"
    
    Call SessionGateway.saveSession(workspace, email, envString, "", "", "")
    Debug.Print "Session cleared"
    
    Set response = AuthGateway.createNewSession(workspace, email, password)
    Debug.Print "New session created"
    
    If response("error").Count <> 0 Then
        MsgBox response("error")("message"), , "Erro"
        Exit Sub
    End If
    
    accessToken = response("success")("accessToken")
    memberName = response("success")("member")("name")
    workspaceId = response("success")("member")("workspaceId")
    
    Call SessionGateway.saveSession(workspace, email, envString, accessToken, memberName, workspaceId)
    Debug.Print "New session data saved"
    
    Call SessionGateway.displayMemberInfo
    Debug.Print "Information printed"
    
    Unload Me
    
End Sub
