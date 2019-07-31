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
    Dim teamInfo As String: teamInfo = TeamBox.value
    
    Call Utils.applyStandardLayout("H")
    
    'Headers definition
    ActiveSheet.Cells(9, 1).value = "Nome"
    ActiveSheet.Cells(9, 2).value = "CPF/CNPJ"
    ActiveSheet.Cells(9, 3).value = "Valor"
    ActiveSheet.Cells(9, 4).value = "Código do Banco"
    ActiveSheet.Cells(9, 5).value = "Agência"
    ActiveSheet.Cells(9, 6).value = "Conta"
    ActiveSheet.Cells(9, 7).value = "Tags"
    ActiveSheet.Cells(9, 8).value = "Descrição"
    
    With ActiveWindow
        .SplitColumn = 8
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    With CreateObject("VBScript.RegExp")
        .Pattern = "\= ([^)]+)\)"
        .Global = True
        For Each m In .Execute(teamInfo)
            teamId = m.SubMatches(0)
        Next
    End With

    Set orders = TeamGateway.getOrders()
    respMessage = TeamGateway.createOrders(teamId, orders)
    
    Unload Me
     
End Sub
