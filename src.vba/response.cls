Public Status As Integer
Public content As String

Public Function json()
    Set json = JsonConverter.ParseJson(content)
End Function

Public Function error()
    If Status = 200 Then
        Set error = JsonConverter.ParseJson("{}")
    End If
    
    Set error = json()("error")
End Function

Public Function errors()
    If Status = 200 Then
        Set errors = JsonConverter.ParseJson("{}")
    End If
    
    Set errors = json()
End Function