Public Sub ListWorkspace()
    Dim cursor As String
    Dim workspaces As Collection
    Dim row As Integer
    Dim workspaceCreated As String
    Dim workspaceId As String
    
    If Not isSignedin Then
        MsgBox "Acesso negado. Faça login novamente.", , "Erro"
        Exit Sub
    End If
    
    Call postSessionV1(True, "")
    
    'Table layout
    Utils.applyStandardLayout ("F")
    Range("A10:F" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Número da Conta (Workspace ID)"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Username"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Data"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "CPF / CNPJ permitidos"
    
    With ActiveWindow
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    row = HeaderRow() + 1
    
    On Error GoTo eh
    Do
        Set respJson = getWorkspace(cursor)

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set workspaces = respJson("workspaces")

        For Each workspace In workspaces
            workspaceCreated = workspace("created")
            workspaceId = workspace("id")
            Dim allowedTaxIds As Collection: Set allowedTaxIds = workspace("allowedTaxIds")
            ActiveSheet.Cells(row, 1).Value = workspaceId
            ActiveSheet.Cells(row, 2).Value = workspace("name")
            ActiveSheet.Cells(row, 3).Value = workspace("username")
            ActiveSheet.Cells(row, 4).Value = Utils.ISODATEZ(workspaceCreated)
            ActiveSheet.Cells(row, 5).Value = CollectionToString(allowedTaxIds, ", ")
            
            row = row + 1
        Next

    Loop While cursor <> ""
    
    Exit Sub
eh:
    
End Sub

Public Function getWorkspace(cursor As String)
    Dim query As String
    Dim resp As response
    
    query = ""
    If cursor <> "" Then
        query = "?cursor=" + cursor
    End If
    
    Set resp = V2Rest.getRequest("/v2/workspace", query, New Dictionary)
    
    If resp.Status >= 300 Then
        MsgBox resp.errors()("errors")(1)("message"), , "Erro"
    End If
    Set getWorkspace = resp.json()

End Function
