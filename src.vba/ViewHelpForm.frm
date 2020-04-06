
Private Sub GeneratePublicPrivateButton_Click()
    ActiveWorkbook.FollowHyperlink address:="https://starkbank.com/br/faq/how-to-create-ecdsa-keys"
End Sub

Private Sub HelpProductionWebButton_Click()
    ActiveWorkbook.FollowHyperlink address:="https://web.starkbank.com"
End Sub

Private Sub HelpSandboxWebButton_Click()
    ActiveWorkbook.FollowHyperlink address:="https://sandbox.web.starkbank.com"
End Sub

Private Sub Label1_Click()

End Sub

Private Sub SendPublicKeyButton_Click()
    SendKeyForm.Show
End Sub