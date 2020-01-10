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

Public Function download(url As String, path As String) As Boolean
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    objHTTP.Open "GET", url, False
    
    objHTTP.send
    
    If objHTTP.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write objHTTP.responseBody
        oStream.SaveToFile path, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        download = True
    Else
        download = False
    End If
End Function
