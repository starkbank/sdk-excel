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
    Dim result As Dictionary
    Set result = New Dictionary
    result.Add "Content-Type", "Application/json"
    result.Add "Accept-Language", "pt-BR"
    result.Add "Access-Token", SessionGateway.getAccessToken()
    
    Set defaultHeaders = result
End Function

Public Function getRequest(path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    
    For Each Key In defaultHeaders().keys()
        headers.Add Key, defaultHeaders()(Key)
    Next
    
    Set getRequest = Request.fetch(url, "GET", headers, "")
End Function

Public Function postRequest(path As String, payload As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    
    For Each Key In defaultHeaders().keys()
        headers.Add Key, defaultHeaders()(Key)
    Next
    
    Set postRequest = Request.fetch(url, "POST", headers, payload)
End Function

Public Function deleteRequest(path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseUrl() + path + query
    
    For Each Key In defaultHeaders().keys()
        headers.Add Key, defaultHeaders()(Key)
    Next
    
    Set deleteRequest = Request.fetch(url, "DELETE", headers, "")
End Function

Public Function externalPostRequest(url As String, payload As String)
    Dim headers As Dictionary: Set headers = New Dictionary
    headers.Add "Content-Type", "Application/json"
    
    Set externalPostRequest = Request.fetch(url, "POST", headers, payload)
End Function
