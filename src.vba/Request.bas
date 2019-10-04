Public Function fetch(url As String, method As String, headers As Dictionary, payload As String)
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open method, url, False
    
    For Each key In headers.keys()
        objHTTP.setRequestHeader key, headers(key)
    Next
    
    objHTTP.send payload
    
    Dim resp As response
    Set resp = New response
    
    resp.Status = objHTTP.Status
    resp.content = objHTTP.responseText
    
    Set fetch = resp

End Function