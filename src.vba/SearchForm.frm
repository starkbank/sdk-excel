Private Sub AfterTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    AfterTextBox.Text = Utils.formatDateInUserForm(AfterTextBox.Text)
    reentry = False
End Sub

Private Sub BeforeTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    BeforeTextBox.Text = Utils.formatDateInUserForm(BeforeTextBox.Text)
    reentry = False
End Sub

Private Sub DownloadButton_Click()
    On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim cursor As String
    Dim transactions As Collection
    Dim row As Integer
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    
    Call Utils.applyStandardLayout("F")
    ActiveSheet.Range("A10:F" & Rows.Count).ClearContents
    
    'Headers definition
    ActiveSheet.Cells(9, 1).Value = "Data"
    ActiveSheet.Cells(9, 2).Value = "Valor"
    ActiveSheet.Cells(9, 3).Value = "Descrição"
    ActiveSheet.Cells(9, 4).Value = "Id da Transação"
    ActiveSheet.Cells(9, 5).Value = "Tarifa"
    ActiveSheet.Cells(9, 6).Value = "Tags"
    
    With ActiveWindow
        .SplitColumn = 6
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    'Optional parameters
    optionalParam.Add "after", after
    optionalParam.Add "before", before
    
    row = 10
    
    Do
        Set respJson = getTransaction(cursor, optionalParam)
        
        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If
            
        Set transactions = respJson("transactions")
        
        For Each Transaction In transactions
            Dim signal As Integer: signal = 1
            If Transaction("flow") = "out" Then
                signal = -1
            End If
            
            Dim created As String: created = Transaction("created")
            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(created)
            
            ActiveSheet.Cells(row, 2).Value = Transaction("amount") / 100 * signal
            ActiveSheet.Cells(row, 3).Value = Transaction("description")
            ActiveSheet.Cells(row, 4).Value = Transaction("id")
            ActiveSheet.Cells(row, 5).Value = Transaction("fee") / 100
            
            Dim tags As Collection: Set tags = Transaction("tags")
            ActiveSheet.Cells(row, 6).Value = CollectionToString(tags, ",")
            
            row = row + 1
        Next
    
    Loop While cursor <> ""
    
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
    
End Sub