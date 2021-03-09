Public Enum Environment
    development
    sandbox
    production
End Enum

Public Function baseUrl()
    Select Case SessionGateway.getEnvironment()
        Case development: baseUrl = "https://development.api.starkbank.com"
        Case sandbox:     baseUrl = "https://sandbox.api.starkbank.com"
        Case production:  baseUrl = "https://api.starkbank.com"
    End Select
End Function

Public Function defaultHeaders()
    Dim Result As Dictionary
    Set Result = New Dictionary
    Result.Add "Content-Type", "Application/json"
    Result.Add "Accept-Language", "pt-BR"
    Result.Add "Access-Token", SessionGateway.getAccessToken()
    
    Set defaultHeaders = Result
End Function

Public Function pdfHeaders()
    Dim Result As Dictionary
    Set Result = New Dictionary
    Result.Add "Accept-Language", "pt-BR"
    Result.Add "Access-Token", SessionGateway.getAccessToken()
    
    Set pdfHeaders = Result
End Function

Public Function getRequest(path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    
    For Each key In defaultHeaders().keys()
        headers.Add key, defaultHeaders()(key)
    Next
    
    Set getRequest = Request.fetch(url, "GET", headers, "")
End Function

Public Function postRequest(path As String, payload As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    
    For Each key In defaultHeaders().keys()
        headers.Add key, defaultHeaders()(key)
    Next
    
    Set postRequest = Request.fetch(url, "POST", headers, payload)
End Function

Public Function deleteRequest(path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    
    For Each key In defaultHeaders().keys()
        headers.Add key, defaultHeaders()(key)
    Next
    
    Set deleteRequest = Request.fetch(url, "DELETE", headers, "")
End Function

Public Function downloadRequest(path As String, filepath As String, headers As Dictionary, fallbackName As String) As Boolean
    Dim url As String: url = baseUrl() + path
    
    For Each key In pdfHeaders().keys()
        headers.Add key, pdfHeaders()(key)
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
