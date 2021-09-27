
Public Sub saveSession(workspace As String, email As String, envString As String, accessToken As String, memberName As String, workspaceId As String)
    Sheets("Credentials").Cells(1, 1) = "Workspace"
    Sheets("Credentials").Cells(1, 2) = workspace
    Sheets("Credentials").Cells(2, 1) = "E-mail"
    Sheets("Credentials").Cells(2, 2) = email
    Sheets("Credentials").Cells(3, 1) = "Environment"
    Sheets("Credentials").Cells(3, 2) = envString
    Sheets("Credentials").Cells(4, 1) = "AccessToken"
    Sheets("Credentials").Cells(4, 2) = accessToken
    Sheets("Credentials").Cells(5, 1) = "Member.name"
    Sheets("Credentials").Cells(5, 2) = memberName
    Sheets("Credentials").Cells(6, 1) = "Workspace ID"
    Sheets("Credentials").Cells(6, 2) = workspaceId
    Sheets("Credentials").Cells(7, 1) = "Approval Date"
    
    Sheets("Credentials").Cells(11, 1) = "Session Private"
    Sheets("Credentials").Cells(12, 1) = "Session Public"
    Sheets("Credentials").Cells(13, 1) = "Access ID"
       
End Sub

Public Sub saveAccessToken(accessToken As String)
    Sheets("Credentials").Cells(4, 1) = "AccessToken"
    Sheets("Credentials").Cells(4, 2) = accessToken
End Sub

Public Sub displayMemberInfo()
    Dim helloMessage As String
    Dim workspaceMessage As String
    Dim emailMessage As String
    Dim envMessage As String
    Dim balanceMessage As String
    
    helloMessage = "Olá " + SessionGateway.getMemberName() + "!"
    workspaceMessage = "Workspace: " + SessionGateway.getWorkspace()
    workspaceId = "ID do Workspace: " + SessionGateway.getWorkspaceId()
    emailMessage = "E-mail: " + SessionGateway.getEmail()
    envMessage = "Ambiente: " + SessionGateway.getEnvironmentString()
    balanceMessage = "Saldo: " + SessionGateway.getBalance()
    
    For Each WS In ThisWorkbook.Worksheets
        If WS.name <> "Credentials" And WS.name <> "InputLog" And WS.name <> "Aux" Then
            WS.Cells(2, 1).Value = helloMessage
            WS.Cells(3, 1).Value = workspaceMessage
            WS.Cells(4, 1).Value = workspaceId
            WS.Cells(5, 1).Value = emailMessage
            WS.Cells(6, 1).Value = envMessage
            WS.Cells(7, 1).Value = balanceMessage
        End If
    Next
End Sub

Public Function getAccessToken()
    Dim accessToken As String: accessToken = CStr(Sheets("Credentials").Cells(4, 2))

    If accessToken = "" Then
        getAccessToken = "Trash"
    Else
        getAccessToken = accessToken
    End If
    
End Function

Public Function getEnvironment()
    Select Case getEnvironmentString()
        Case "Produção": getEnvironment = production
        Case "Sandbox":  getEnvironment = sandbox
    End Select
End Function

Public Function getEnvironmentString()
    getEnvironmentString = CStr(Sheets("Credentials").Cells(3, 2))
End Function

Public Function getWorkspace()
    getWorkspace = CStr(Sheets("Credentials").Cells(1, 2))
End Function

Public Function getEmail()
    getEmail = CStr(Sheets("Credentials").Cells(2, 2))
End Function

Public Function getMemberName()
    getMemberName = CStr(Sheets("Credentials").Cells(5, 2))
End Function

Public Function getWorkspaceId()
    getWorkspaceId = CStr(Sheets("Credentials").Cells(6, 2))
End Function

Public Function getBalance()
    Dim AccountInfo As Dictionary
    Dim balanceMessage As String
    
    Set AccountInfo = BankGateway.getAccount()
    balanceMessage = "-"
    If AccountInfo.Count > 0 Then
        Dim balance As Double
        balance = CDbl(AccountInfo("account")("balance"))
        balanceMessage = Utils.MoneyStringFrom(balance)
    End If
    getBalance = balanceMessage
End Function


Public Function sessionPrivateKeyContent() As String
    Dim sessionPrivateKeyPath
    sessionPrivateKeyPath = getTempDir() + "\" + "sessionPrivateKey.pem"
    Call Shell("""" + getOpensslDir() + """ ecparam -name secp256k1 -genkey -out """ + sessionPrivateKeyPath + """")
    Application.Wait Now + #12:00:02 AM#
    
    Open sessionPrivateKeyPath For Input As #1
    Do Until EOF(1)
        Line Input #1, textLine
        sessionPrivateKeyContent = sessionPrivateKeyContent & textLine & vbLf
    Loop
    Close #1
    
End Function

Public Function sessionPublicKeyContent() As String
    Dim sessionPrivateKeyPath
    Dim sessionPublicKeyPath
    sessionPrivateKeyPath = getTempDir() + "\" + "sessionPrivateKey.pem"
    sessionPublicKeyPath = getTempDir() + "\" + "sessionPublicKey.pem"
    Call Shell("""" + getOpensslDir() + """ ec -in """ + sessionPrivateKeyPath + """ -pubout -out """ + sessionPublicKeyPath + """")
    Application.Wait Now + #12:00:02 AM#
    
    Open sessionPublicKeyPath For Input As #1
    Do Until EOF(1)
        Line Input #1, textLine
        sessionPublicKeyContent = sessionPublicKeyContent & textLine & vbLf
    Loop
    Close #1
    
End Function

Public Function getSessionPrivateKeyContent() As String
    getSessionPrivateKeyContent = Sheets("Credentials").Cells(11, 2)
End Function

Public Function getSessionPublicKeyContent() As String
    getSessionPublicKeyContent = Sheets("Credentials").Cells(12, 2)
End Function

Public Function getAccessId() As String
    getAccessId = Sheets("Credentials").Cells(13, 2)
End Function

Public Sub RenewSessionKeys()
    Sheets("Credentials").Cells(11, 2) = sessionPrivateKeyContent
    Sheets("Credentials").Cells(12, 2) = sessionPublicKeyContent
    DeleteTempKeys
End Sub

Public Sub DeleteSessionKeys()
    Sheets("Credentials").Cells(11, 2) = ""
    Sheets("Credentials").Cells(12, 2) = ""
    DeleteTempKeys
End Sub

Public Sub DeleteTempKeys()
    Dim sessionPrivateKeyPath
    Dim sessionPublicKeyPath
    sessionPrivateKeyPath = getTempDir() + "\" + "sessionPrivateKey.pem"
    sessionPublicKeyPath = getTempDir() + "\" + "sessionPublicKey.pem"
    Kill sessionPrivateKeyPath
    Kill sessionPublicKeyPath
End Sub

Public Sub postSessionV1()
    Dim payload As String
    Dim resp As response
    Dim Result As New Dictionary
    Dim headers As New Dictionary
    Dim dict As New Dictionary
    
    dict.Add "platform", "web"
    dict.Add "expiration", 5184000
    dict.Add "publicKey", getSessionPublicKeyContent()
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set resp = StarkBankApi.postRequest("/v1/auth/session", payload, headers)
    
    If resp.Status = 200 Then
        Result.Add "success", resp.json()
        Result.Add "error", New Dictionary
    Else
        Result.Add "success", New Dictionary
        Result.Add "error", resp.error()
    End If
    accessId = Result("success")("session")("id")
    
    Sheets("Credentials").Cells(13, 2) = "session/" + accessId
    
End Sub