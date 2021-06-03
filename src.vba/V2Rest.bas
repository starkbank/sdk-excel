
Public Function toUnix(dt) As Long
    toUnix = DateDiff("s", "1/1/1970 00:00:00", dt) + 10800
End Function

Public Function baseUrl()
    Select Case SessionGateway.getEnvironment()
        Case development: baseUrl = "https://development.api.starkbank.com"
        Case sandbox:     baseUrl = "https://sandbox.api.starkbank.com"
        Case production:  baseUrl = "https://api.starkbank.com"
    End Select
End Function

Public Function defaultHeaders(payload As String)
    Dim Result As Dictionary
    Set Result = New Dictionary
    Dim accessTime As Long
    Dim message As String
    
    accessId = getAccessId()
    accessTime = toUnix(Now)
    message = accessId + ":" + CStr(accessTime) + ":" + payload
    
    Dim pk As PrivateKey: Set pk = New PrivateKey
    pk.fromPem (getSessionPrivateKeyContent)
    
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(message, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    Result.Add "Content-Type", "Application/json"
    Result.Add "Accept-Language", "pt-BR"
    Result.Add "Access-Time", accessTime
    Result.Add "Access-Id", accessId
    Result.Add "Access-Signature", signature64
    
    If DebugModeOn() Then
        DebugPrint "headers", JsonConverter.ConvertToJson(Result)
    End If
    
    Set defaultHeaders = Result
End Function

Public Function pdfHeaders(payload As String)
    Dim Result As Dictionary
    Set Result = New Dictionary
    Dim accessTime As Long
    Dim message As String
    
    accessTime = toUnix(Now)
    message = getAccessId() + ":" + CStr(accessTime) + ":" + payload
    
    Dim pk As PrivateKey: Set pk = New PrivateKey
    pk.fromPem (getSessionPrivateKeyContent)
    
    Dim signature As signature: Set signature = EllipticCurve_Ecdsa.sign(message, pk)
    Dim signature64 As String: signature64 = signature.toBase64()
    
    Result.Add "Accept-Language", "pt-BR"
    Result.Add "Access-Time", accessTime
    Result.Add "Access-Id", getAccessId()
    Result.Add "Access-Signature", signature64
    
    Set pdfHeaders = Result
End Function

Public Function getRequest(path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    Dim defHeaders As New Dictionary
    
    Set defHeaders = defaultHeaders("")
    
    For Each key In defHeaders.keys()
        headers.Add key, defHeaders(key)
    Next
    
    Set getRequest = Request.fetch(url, "GET", headers, "")
End Function

Public Function postRequest(path As String, payload As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    Dim defHeaders As New Dictionary
    
    Set defHeaders = defaultHeaders(payload)
    For Each key In defHeaders.keys()
        headers.Add key, defHeaders(key)
    Next
    
    Set postRequest = Request.fetch(url, "POST", headers, payload)
End Function

Public Function deleteRequest(path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    Dim defHeaders As New Dictionary
    
    Set defHeaders = defaultHeaders("")
    
    For Each key In defHeaders.keys()
        headers.Add key, defHeaders(key)
    Next
    
    Set deleteRequest = Request.fetch(url, "DELETE", headers, "")
End Function

Public Function downloadRequest(path As String, filepath As String, headers As Dictionary, fallbackName As String) As Boolean
    Dim url As String: url = baseUrl() + path
    Dim defHeaders As New Dictionary
    
    Set defHeaders = pdfHeaders("")
    
    For Each key In defHeaders.keys()
        headers.Add key, defHeaders(key)
    Next
    
    downloadRequest = Request.download(url, filepath, headers, fallbackName)
End Function

Public Function uploadRequest(path As String, payload As String, headers As Dictionary, boundary As String)
    Dim url As String: url = baseUrl() + path + query
    
    For Each key In defaultHeaders().keys()
        headers.Add key, defaultHeaders()(key)
    Next
    headers("Content-Type") = "multipart/form-data; boundary=" + boundary
    
    Set uploadRequest = Request.fetch(url, "POST", headers, payload)
End Function

Public Function externalPostRequest(url As String, payload As String)
    Dim headers As Dictionary: Set headers = New Dictionary
    headers.Add "Content-Type", "Application/json"
    
    Set externalPostRequest = Request.fetch(url, "POST", headers, payload)
End Function

