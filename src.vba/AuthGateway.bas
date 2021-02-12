Public Function createNewSession(workspace As String, email As String, password As String)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    Dim Result As New Dictionary
    
    dict.Add "workspace", workspace
    dict.Add "email", email
    dict.Add "password", password
    dict.Add "platform", "web"
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set resp = StarkBankApi.postRequest("/v1/auth/access-token", payload, New Dictionary)
    
    If resp.Status = 200 Then
        Result.Add "success", resp.json()
        Result.Add "error", New Dictionary
    Else
        Result.Add "success", New Dictionary
        Result.Add "error", resp.error()
    End If
    Set createNewSession = Result
End Function

Public Function deleteSession(accessToken As String)
    Dim resp As response
    Dim Result As New Dictionary
    
    Set resp = StarkBankApi.deleteRequest("/v1/auth/access-token/" + accessToken, "", New Dictionary)
    
    If resp.Status = 200 Then
        Result.Add "success", resp.json()
        Result.Add "error", New Dictionary
    Else
        Result.Add "success", New Dictionary
        Result.Add "error", resp.error()
    End If
    Set deleteSession = Result
End Function