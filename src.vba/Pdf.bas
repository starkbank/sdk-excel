Public Function downloadChargePdf(chargeId As String, folder As String)
    Dim success As Boolean
    Dim path As String
    Dim filepath As String
    
    path = "/charge/" + chargeId + "/pdf"
    filepath = folder + "/boleto-" + chargeId + ".pdf"
    
    downloadChargePdf = StarkBankApi.downloadRequest(path, filepath)
End Function

Public Sub downloadAllPdfs()
    Dim lastRow As Integer
    Dim numberCharges As Integer
    Dim folder As String
    Dim chargeId As String
    Dim success As Boolean
    Dim anyFailed As Boolean
    Dim tooMany As Boolean
    anyFailed = False
    tooMany = True
    
    lastRow = ActiveSheet.Range("H9").CurrentRegion.Rows.Count + 8
    Worksheets("Consulta de Boletos Emitidos").Activate
    folder = ActiveWorkbook.path + "/starkbank-boletos"
    createNewDirectory (folder)
    
    numberCharges = lastRow - 9
    If numberCharges = 0 Then
        MsgBox "Nenhum boleto para baixar. Clique em Consultar Boletos", vbExclamation
        Exit Sub
    ElseIf numberCharges >= 10 Then
        Dim longTimeMessage As String
        Dim downloadTime As Double
        downloadTime = 3.2 * numberCharges / 60
        longTimeMessage = "Há " + CStr(numberCharges) + " boletos para baixar. Esta operação deve levar cerca de " + CStr(Round(downloadTime)) + " minuto(s). Continuar?"
        If MsgBox(longTimeMessage, vbExclamation + vbYesNo) = vbYes Then
            tooMany = False
        End If
    Else
        tooMany = False
    End If
    
    If Not tooMany Then
        For i = 10 To lastRow
            chargeId = CStr(Cells(i, "H").Value)
            success = downloadChargePdf(chargeId, folder)
            If Not success Then
                anyFailed = True
            End If
        Next
        
        If anyFailed Then
            MsgBox "Alguns boletos tiveram falha no download!", vbExclamation
        Else
            MsgBox "Arquivos salvos com sucesso em:" + vbNewLine + folder
        End If
    End If
End Sub

Public Sub createNewDirectory(directoryName As String)
    On Error Resume Next
    If Not DirExists(directoryName) Then
        MkDir (directoryName)
    End If
End Sub
 
Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    Exit Function
End Function