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
    Dim afterInput As String: afterInput = AfterTextBox.value
    Dim beforeInput As String: beforeInput = BeforeTextBox.value
    
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
    ActiveSheet.Cells(9, 1).value = "Data"
    ActiveSheet.Cells(9, 2).value = "Valor"
    ActiveSheet.Cells(9, 3).value = "Descrição"
    ActiveSheet.Cells(9, 4).value = "Id da Transação"
    ActiveSheet.Cells(9, 5).value = "Tarifa"
    ActiveSheet.Cells(9, 6).value = "Tags"
    
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
            ActiveSheet.Cells(row, 1).value = Utils.ISODATEZ(created)
            
            ActiveSheet.Cells(row, 2).value = Transaction("amount") / 100 * signal
            ActiveSheet.Cells(row, 3).value = Transaction("description")
            ActiveSheet.Cells(row, 4).value = Transaction("id")
            ActiveSheet.Cells(row, 5).value = Transaction("fee") / 100
            
            Dim tags As Collection: Set tags = Transaction("tags")
            ActiveSheet.Cells(row, 6).value = CollectionToString(tags, ",")
            
            row = row + 1
        Next
    
    Loop While cursor <> ""
    
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Me.AfterTextBox.value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.value = InputLogGateway.getBeforeDate()
    
End Sub