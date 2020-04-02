
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
    workspaceMessage = "Workspace: " + SessionGateway.getWorkspace() + " (" + SessionGateway.getWorkspaceId() + ")"
    emailMessage = "E-mail: " + SessionGateway.getEmail()
    envMessage = "Ambiente: " + SessionGateway.getEnvironmentString()
    balanceMessage = "Saldo: " + SessionGateway.getBalance()
    
    For Each WS In ThisWorkbook.Worksheets
        If WS.name <> "Credentials" And WS.name <> "InputLog" Then
            WS.Cells(2, 1).Value = helloMessage
            WS.Cells(3, 1).Value = workspaceMessage
            WS.Cells(4, 1).Value = emailMessage
            WS.Cells(5, 1).Value = envMessage
            WS.Cells(6, 1).Value = balanceMessage
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
    If AccountInfo.count > 0 Then
        Dim balance As Long
        balance = AccountInfo("account")("balance")
        balanceMessage = Utils.MoneyStringFrom(balance)
    End If
    getBalance = balanceMessage
End Function
