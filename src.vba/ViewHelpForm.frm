
Private Sub CreateLogButton_Click()
    Dim initRow As Long
    Dim lastRow As Long
    Dim filename As String
    Dim filepath As String
    Dim datetimenow As String
    
    initRow = 10
    lastRow = CurrentDebugRow() - 1
    
    If lastRow < 10 Then
        MsgBox "Não há nenhuma entrada de log para ser salva. Ative o Modo Debug e execute alguma operação.", vbExclamation
        Exit Sub
    End If
    datetimenow = Format(Now, "yyyy-mm-ddTh-nn-ss")
    filename = "starkbank-sdk-excel-log-" + getCurrentVersion() + "-" + datetimenow + ".log"
    filepath = ThisWorkbook.path + "\" + filename
    Open filepath For Output As #1
    Print #1, "Datetime;Type;Data"
    With Sheets("InputLog")
        For i = initRow To lastRow
            Print #1, CStr(.Cells(i, 1)) + ";" + CStr(.Cells(i, 2)) + ";" + CStr(.Cells(i, 3))
        Next
    End With
    Close #1
    MsgBox "Arquivo salvo com sucesso em:" + vbNewLine + filepath
End Sub

Public Sub UserForm_Activate():
    If DebugModeOn() Then
        DebugCheckBox.Value = True
    End If
End Sub

Private Sub DebugCheckBox_Click()
    If DebugCheckBox.Value = True Then
        Sheets("InputLog").Cells(4, 2) = True
    Else
        Sheets("InputLog").Cells(4, 2) = False
    End If
End Sub

Private Sub GeneratePublicPrivateButton_Click()
    ActiveWorkbook.FollowHyperlink address:="https://starkbank.com/br/faq/how-to-create-ecdsa-keys"
End Sub

Private Sub HelpProductionWebButton_Click()
    ActiveWorkbook.FollowHyperlink address:="https://web.starkbank.com/signup/email"
End Sub

Private Sub HelpSandboxWebButton_Click()
    ActiveWorkbook.FollowHyperlink address:="https://starkbank.com/sandbox"
End Sub


Private Sub SendPublicKeyButton_Click()
    SendKeyForm.Show
End Sub
