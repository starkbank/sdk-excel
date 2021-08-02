Public Sub listBalance()
    Dim warningMessage As String
    Dim minutesDuration As Long
    Dim workspaceId As String
    Dim workspaceList As Collection: Set workspaceList = SheetParser.dict("Listar Contas")
    Dim workspaceListLength As Long
    Dim row As Long
    
    On Error GoTo eh
    
    workspaceListLength = workspaceList.Count()
    If workspaceListLength = 0 Then
        MsgBox "É necessário listar as contas antes de atualizar os saldos de conta!", vbExclamation
        Exit Sub
    End If
    
    minutesDuration = 5# * (workspaceListLength) / 60# + 1
    warningMessage = "A operação de atualizar o saldo das " + CStr(workspaceListLength) + " contas deve demorar cerca de " + CStr(minutesDuration) + " minutos. Continuar?"
    If MsgBox(warningMessage, vbYesNo, "Operação lenta") = vbNo Then
        Exit Sub
    End If
    
    row = 10
    For Each workspace In workspaceList
        workspaceId = workspace("Número da Conta (Workspace ID)")
        
        Call postSessionV1(True, CStr(workspaceId))
        
        Do
            Set respJson = V2BankGateway.getBalance(workspaceId)
            If respJson.Exists("error") Then
                Exit Sub
            End If
    
            Set balance = respJson("balance")
            
            ActiveSheet.Cells(row, 6).Value = CDbl(balance("amount")) / 100
            
            row = row + 1
    
        Loop While cursor <> ""
    Next
eh:

End Sub

Public Function getBalance(id As String)
    Dim query As String
    Dim resp As response
    
    query = ""
    Set resp = V2Rest.getRequest("/v2/balance/" + id, query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.errors()("errors")(1)("message"), , "Erro"
    End If
    Set getBalance = resp.json()

End Function

Public Function getTransaction(cursor As String, optionalParam As Dictionary)
    Dim query As String
    Dim resp As response
    
    query = ""
    If cursor <> "" Then
        query = "?cursor=" + cursor
    End If
    
    If optionalParam.Count > 0 Then
        For Each key In optionalParam
            If query = "" Then
                query = "?" + key + "=" + CStr(optionalParam(key))
            Else
                query = query + "&" + key + "=" + CStr(optionalParam(key))
            End If
        Next
    End If
    
    Set resp = V2Rest.getRequest("/v2/transaction", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.errors()("errors")(1)("message"), , "Erro"
    End If
    Set getTransaction = resp.json()

End Function

