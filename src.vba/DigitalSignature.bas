Public Function mailToken()
    Dim resp As response
    Dim dict As New Dictionary
    Dim Result As New Dictionary
    
    Set resp = StarkBankApi.postRequest("/v1/auth/public-key/token", "", New Dictionary)
    
    If resp.Status = 200 Then
        Result.Add "success", resp.json()
        Result.Add "error", New Dictionary
    Else
        Result.Add "success", New Dictionary
        Result.Add "error", resp.error()
    End If
    Set mailToken = Result
End Function

Public Function sendPublicKey(workspaceId As String, memberId As String, token As String, keyPath As String)
    Dim resp As response
    Dim boundary As String
    Dim payload As String
    Dim publicKeyContent As String
    Dim headers As New Dictionary
    Dim dict As New Dictionary
    Dim Result As New Dictionary
    Dim file_name As String
    Dim file_length As Long
    Dim fnum As Integer
    Dim bytes() As Byte
    
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Charset = "utf-8"
    oStream.Type = 2
    oStream.Open
    oStream.LoadFromFile keyPath
    publicKeyContent = oStream.ReadText
    
    dict.Add "workspaceId", workspaceId
    dict.Add "memberId", memberId
    dict.Add "token", token
    dict.Add "publicKey", publicKeyContent
    
    boundary = String(6, "-") & "publicKeyRequestBoundary"
    payload = ""
    For Each sName In dict
        payload = payload & "--" & boundary & vbCrLf
        payload = payload & "Content-Disposition: form-data; name=""" & sName & """" & vbCrLf & vbCrLf
        payload = payload & dict(sName) & vbCrLf
    Next
    payload = payload & "--" & boundary & "--"
    Set resp = StarkBankApi.uploadRequest("/v1/auth/public-key", payload, headers, boundary)
    
    If resp.Status = 200 Then
        Result.Add "success", resp.json()
        Result.Add "error", New Dictionary
    Else
        Result.Add "success", New Dictionary
        Result.Add "error", resp.error()
    End If
    Set sendPublicKey = Result
End Function