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