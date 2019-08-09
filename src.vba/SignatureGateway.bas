Public Function signMessage(privateKey As String, message As String)
    Dim resp As response
    Dim payload As String
    Dim dict As New Dictionary
    Dim result As New Dictionary
    
    dict.Add "privateKey", privateKey
    dict.Add "message", message
    
    payload = JsonConverter.ConvertToJson(dict)
    
    Set resp = StarkBankApi.externalPostRequest("https://us-central1-api-ms-auth-sbx.cloudfunctions.net/ecdsaSigner", payload)
    
    If resp.Status = 200 Then
        result.Add "success", resp.json()
        result.Add "error", New Dictionary
        Set signMessage = result
    Else
        result.Add "success", New Dictionary
        result.Add "error", resp.error()
        Set signMessage = result
    End If

End Function
