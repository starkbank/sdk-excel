

Public Function getOwnerName(workspaceId As String, optionalParam As Dictionary)
    Dim query As String
    Dim resp As response
    Dim elem As Variant
    Dim path As String
    
    query = ""
    
    path = "/v1/workspace/" + workspaceId + "/owner"
    
    Set resp = StarkBankApi.getRequest(path, query, New Dictionary)
    
    If resp.Status = 200 Then
        Set logArray = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro" + " Response status: " + CStr(resp.Status)
    End If
    getOwnerName = logArray("owner")("name")

End Function