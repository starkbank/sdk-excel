

Private Sub BrowseFolderKeys_Click()
    Dim myFile As String
    myFile = Utils.GetFolder()
    If CStr(myFile) <> "False" Then
        Me.DirBoxKeys.Value = myFile
    End If
End Sub

Private Sub BrowseFolderOpenssl_Click()
    Dim myFile As String
    myFile = Utils.GetFolder()
    If CStr(myFile) <> "False" Then
        Me.DirBoxOpenssl.Value = myFile
    End If
End Sub

Private Sub keyGeneration(diropenssl As String, dirkeys As String)
    Dim opensslPath As String
    Dim commandPrivate As String
    Dim commandPublic As String
    Dim namePublicKey As String
    Dim namePrivateKey As String
    Dim pathPublicKey As String
    Dim pathPrivateKey As String
    
    If dir(diropenssl + "\openssl.exe") <> "" Then
        opensslPath = diropenssl + "\openssl.exe"
    ElseIf dir(diropenssl + "\bin\openssl.exe") <> "" Then
        opensslPath = diropenssl + "\bin\openssl.exe"
    Else
        GoTo eh
    End If
    
    opensslPath = """" + opensslPath + """"
    namePrivateKey = "chavePrivada.pem"
    namePublicKey = "chavePublica.pem"
    
    pathPrivateKey = """" + dirkeys + "\" + namePrivateKey + """"
    pathPublicKey = """" + dirkeys + "\" + namePublicKey + """"
    
    commandPrivate = opensslPath + " ecparam -name secp256k1 -genkey -out " + pathPrivateKey
    commandPublic = opensslPath + " ec -in " + pathPrivateKey + " -pubout -out " + pathPublicKey
    
    Utils.ShellRun (commandPrivate)
    MsgBox "Chave privada gerada. Lembre-se: Sua chave privada NUNCA deve ser compartilhada com qualquer pessoa!", vbExclamation
    Utils.ShellRun (commandPublic)
    MsgBox "Chave pública gerada."
    
    MsgBox "Chaves geradas com sucesso em " + dirkeys
    Exit Sub
eh:
    MsgBox "openssl.exe não encontrado em " + diropenssl, vbExclamation
End Sub

Private Sub GenerateKeys_Click()
    Dim keyFolder As String
    Dim opensslFolder As String
    
    keyFolder = ""
    
    If Me.DirBoxOpenssl.Value <> "" And Me.DirBoxKeys.Value <> "" Then
        Call keyGeneration(Me.DirBoxOpenssl.Value, Me.DirBoxKeys.Value)
    Else
        MsgBox "Nenhuma pasta selecionada!", vbExclamation
    End If
End Sub

Private Sub HelpButton_Click()
    With ViewHelpForm
        .MultiPage1.Value = 2
        .Show
    End With
End Sub