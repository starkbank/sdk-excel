Public Function fetch(url As String, method As String, headers As Dictionary, payload As String)
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open method, url, False
    
'    Debug.Print url
'    Debug.Print payload
    For Each key In headers.keys()
        objHTTP.setRequestHeader key, headers(key)
    Next
    
    Dim resp As response
    Set resp = New response
    On Error GoTo eh:
    objHTTP.send payload
    
    resp.Status = objHTTP.Status
    resp.content = objHTTP.responseText
'    Debug.Print resp.content
    Set fetch = resp
    Exit Function
eh:
    resp.Status = 404
    resp.content = "{""error"":{""code"":""connectionError"",""message"":""Verifique sua conexão de internet!""}}"
    Set fetch = resp
End Function

Public Function download(url As String, path As String, headers As Dictionary) As Boolean
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", url, False
    Dim filepath As String
    
    For Each key In headers.keys()
        objHTTP.setRequestHeader key, headers(key)
    Next
    
    On Error GoTo eh:
    objHTTP.send
    If objHTTP.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write objHTTP.responseBody
        filepath = path + Split(objHTTP.getResponseHeader("Content-Disposition"), "filename=")(1)
        oStream.SaveToFile filepath, 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
        download = True
    Else
eh:
        download = False
    End If
End Function
