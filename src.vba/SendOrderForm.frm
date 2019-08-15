Private Sub UserForm_Initialize()
    Dim cursor As String
    Dim teams As Collection
    
    Do
        Set respJson = getTeams(cursor, New Dictionary)
    
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If
            
        Set teams = respJson("teams")
        
        For Each team In teams
            Me.TeamBox.AddItem team("name") + " (id = " + team("id") + ")"
        Next
    
    Loop While cursor <> ""
    
End Sub

Private Sub SendButton_Click()
    On Error Resume Next
    Dim teamId As String
    Dim orders As Collection
    Dim respMessage As String
    Dim teamInfo As String: teamInfo = TeamBox.Value
    
    Call Utils.applyStandardLayout("H")
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Nome"
    ActiveSheet.Cells(9, 2).Value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 3).Value = "Valor"
    ActiveSheet.Cells(9, 4).Value = "Código do Banco"
    ActiveSheet.Cells(9, 5).Value = "Agência"
    ActiveSheet.Cells(9, 6).Value = "Conta"
    ActiveSheet.Cells(9, 7).Value = "Tags"
    ActiveSheet.Cells(9, 8).Value = "Descrição"
    
    With ActiveWindow
        .SplitColumn = 8
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "\= ([^)]+)\)"
        .Global = True
        For Each M In .Execute(teamInfo)
            teamId = M.SubMatches(0)
        Next
    End With

    Set orders = TeamGateway.getOrders()
    respMessage = TeamGateway.createOrders(teamId, orders)
    
    Unload Me
     
End Sub
