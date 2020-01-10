Public Function createNewSession(workspace As String, email As String, password As String)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    Dim result As New Dictionary
    
    dict.Add "workspace", workspace
    dict.Add "email", email
    dict.Add "password", password
    dict.Add "platform", "web"
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set resp = StarkBankApi.postRequest("/v1/auth/access-token", payload, New Dictionary)
    
    If resp.Status = 200 Then
        result.Add "success", resp.json()
        result.Add "error", New Dictionary
    Else
        result.Add "success", New Dictionary
        result.Add "error", resp.error()
    End If
    Set createNewSession = result
End Function

Public Function deleteSession(accessToken As String)
    Dim resp As response
    Dim result As New Dictionary
    
    Set resp = StarkBankApi.deleteRequest("/v1/auth/access-token/" + accessToken, "", New Dictionary)
    
    If resp.Status = 200 Then
        result.Add "success", resp.json()
        result.Add "error", New Dictionary
    Else
        result.Add "success", New Dictionary
        result.Add "error", resp.error()
    End If
    Set deleteSession = result
End Function