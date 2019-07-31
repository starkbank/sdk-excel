
Public Sub saveSession(workspace As String, email As String, envString As String, accessToken As String, memberName As String)
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
       
End Sub

Public Function displayMemberInfo()
    For Each ws In ThisWorkbook.Worksheets
        If ws.name <> "Credentials" And ws.name <> "InputLog" Then
            ws.Cells(2, 1).value = "Olá " + SessionGateway.getMemberName() + "!"
            ws.Cells(3, 1).value = "Workspace: " + SessionGateway.getWorkspace()
            ws.Cells(4, 1).value = "E-mail: " + SessionGateway.getEmail()
            ws.Cells(5, 1).value = "Ambiente: " + getEnvironmentString()
        End If
    Next
End Function

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
