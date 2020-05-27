
Public Sub downloadAllTransferPdfs()
    downloadAllPdfs ("transfer")
End Sub

Public Sub downloadAllChargePdfs()
    downloadAllPdfs ("charge")
End Sub

Public Sub downloadSelectedTransferPdfs()
    Dim initRow As Long
    Dim lastRow As Long
    Dim service As String
    service = "transfer"
    
    initRow = Utils.Max(Selection.row, 10)
    lastRow = Utils.Min(Selection.row + Selection.Rows.Count - 1, ActiveSheet.Range(ColumnId(service) + "9").CurrentRegion.Rows.Count + 8)
    
    If initRow > lastRow Then
        MsgBox "Nenhuma transferência válida selecionada"
        Exit Sub
    End If
    
    Call downloadPdfRange(service, initRow, lastRow)
End Sub

Public Sub downloadSelectedChargePdfs()
    Dim initRow As Long
    Dim lastRow As Long
    Dim service As String
    service = "charge"
    
    initRow = Utils.Max(Selection.row, 10)
    lastRow = Utils.Min(Selection.row + Selection.Rows.Count - 1, ActiveSheet.Range(ColumnId(service) + "9").CurrentRegion.Rows.Count + 8)
    
    If initRow > lastRow Then
        MsgBox "Nenhum boleto válido selecionado"
        Exit Sub
    End If
    
    Call downloadPdfRange(service, initRow, lastRow)
End Sub

Public Function downloadSinglePdf(service, id As String, folder As String)
    Dim success As Boolean
    Dim path As String
    Dim filepath As String
    
    path = "/v1/" + service + "/" + id + "/pdf"
    filepath = folder + "/" + service + "-" + id + ".pdf"
    
    downloadSinglePdf = StarkBankApi.downloadRequest(path, filepath, New Dictionary)
End Function

Public Sub downloadPdfRange(service As String, initRow, lastRow)
    Dim numberEntities As Integer
    Dim folder As String
    Dim idColumn As String
    Dim entityId As String
    Dim success As Boolean
    Dim anyFailed As Boolean
    Dim tooMany As Boolean
    anyFailed = False
    tooMany = True
    
    idColumn = ColumnId(service)
    
    folder = ActiveWorkbook.path + "/starkbank-pdf-" + service
    numberEntities = lastRow - initRow + 1
    
    createNewDirectory (folder)
    If numberEntities = 0 Then
        MsgBox "Nenhum arquivo para baixar. Clique em Consultar", vbExclamation
        Exit Sub
    ElseIf numberEntities >= 10 Then
        Dim longTimeMessage As String
        Dim downloadTime As Double
        downloadTime = 3.2 * numberEntities / 60
        longTimeMessage = "Há " + CStr(numberEntities) + " arquivos para baixar. Esta operação deve levar cerca de " + CStr(Round(downloadTime)) + " minuto(s). Continuar?"
        If MsgBox(longTimeMessage, vbExclamation + vbYesNo) = vbYes Then
            tooMany = False
        End If
    Else
        tooMany = False
    End If
    
    If Not tooMany Then
        For i = initRow To lastRow
            entityId = CStr(Cells(i, idColumn).Value)
            success = downloadSinglePdf(service, entityId, folder)
            If Not success Then
                anyFailed = True
            End If
        Next
        
        If anyFailed Then
            MsgBox "Alguns arquivos tiveram falha no download!", vbExclamation
        Else
            MsgBox "Arquivos salvos com sucesso em:" + vbNewLine + folder
        End If
    End If
End Sub


Public Sub downloadAllPdfs(service As String)
    Dim initRow As Long
    Dim lastRow As Long
    Dim idColumn As String
    Dim abcd As Range
    
    idColumn = ColumnId(service)
    initRow = 10
    lastRow = ActiveSheet.Range(idColumn + "9").CurrentRegion.Rows.Count + 8

    Call downloadPdfRange(service, initRow, lastRow)
End Sub

Public Function ColumnId(service As String)
    
    Select Case service
        Case "transfer"
            Worksheets("Consulta de Transferências").Activate
            ColumnId = "B"
        Case "charge"
            Worksheets("Consulta de Boletos Emitidos").Activate
            ColumnId = "M"
        Case "charge-payment"
            Worksheets("Consulta de Pagamento Boletos").Activate
            ColumnId = "H"
        Case Else
            ColumnId = "A"
    End Select

End Function

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