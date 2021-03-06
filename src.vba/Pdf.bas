
Public Sub downloadAllTransferPdfs()
    downloadAllPdfs ("transfer")
End Sub

Public Sub downloadAllChargePdfs()
    downloadAllPdfs ("charge")
End Sub

Public Sub downloadAllChargePaymentPdfs()
    downloadAllPdfs ("charge-payment")
End Sub

Public Sub downloadAllInvoicePdfs()
    downloadAllPdfs ("invoice")
End Sub

Public Sub downloadSelectedTransferPdfs()
    Dim initRow As Long
    Dim lastRow As Long
    Dim service As String
    service = "transfer"
    
    initRow = Utils.max(Selection.row, 10)
    lastRow = Utils.min(Selection.row + Selection.Rows.Count - 1, ActiveSheet.Range(ColumnId(service) + "9").CurrentRegion.Rows.Count + 8)
    
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
    
    initRow = Utils.max(Selection.row, 10)
    lastRow = Utils.min(Selection.row + Selection.Rows.Count - 1, ActiveSheet.Range(ColumnId(service) + "9").CurrentRegion.Rows.Count + 8)
    
    If initRow > lastRow Then
        MsgBox "Nenhum boleto válido selecionado"
        Exit Sub
    End If
    
    Call downloadPdfRange(service, initRow, lastRow)
End Sub

Public Sub downloadSelectedChargePaymentPdfs()
    Dim initRow As Long
    Dim lastRow As Long
    Dim service As String
    service = "charge-payment"
    
    initRow = Utils.max(Selection.row, 10)
    lastRow = Utils.min(Selection.row + Selection.Rows.Count - 1, ActiveSheet.Range(ColumnId(service) + "9").CurrentRegion.Rows.Count + 8)
    
    If initRow > lastRow Then
        MsgBox "Nenhum pagamento válido selecionado"
        Exit Sub
    End If
    
    Call downloadPdfRange(service, initRow, lastRow)
End Sub

Public Sub downloadSelectedInvoicePdfs()
    Dim initRow As Long
    Dim lastRow As Long
    Dim service As String
    service = "invoice"
    
    initRow = Utils.max(Selection.row, 10)
    lastRow = Utils.min(Selection.row + Selection.Rows.Count - 1, ActiveSheet.Range(ColumnId(service) + "9").CurrentRegion.Rows.Count + 8)
    
    If initRow > lastRow Then
        MsgBox "Nenhuma Invoice válida selecionada"
        Exit Sub
    End If
    
    Call downloadPdfRange(service, initRow, lastRow)
End Sub


Public Function downloadSinglePdf(service, id As String, folder As String, fallbackName As String)
    Dim success As Boolean
    Dim path As String
    Dim filepath As String
    Dim version As String
    
    Select Case service
        Case "invoice"
            version = "v2"
        Case Else
            version = "v1"
    End Select
    
    path = "/" + version + "/" + service + "/" + id + "/pdf"
    filepath = folder + "/"
    
    Select Case service
        Case "invoice"
            downloadSinglePdf = V2Rest.downloadRequest(path, filepath, New Dictionary, fallbackName)
        Case Else
            downloadSinglePdf = StarkBankApi.downloadRequest(path, filepath, New Dictionary, fallbackName)
    End Select
    
End Function

Public Sub downloadPdfRange(service As String, initRow As Long, lastRow As Long)
    Dim numberEntities As Integer
    Dim i As Long
    Dim folder As String
    Dim idColumn As String
    Dim entityId As String
    Dim fallbackName As String
    Dim success As Boolean
    Dim anyFailed As Boolean
    Dim tooMany As Boolean
    anyFailed = False
    anySuccess = False
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
            fallbackName = getFallbackName(i, entityId, service)
            success = downloadSinglePdf(service, entityId, folder, fallbackName)
            If Not success Then
                anyFailed = True
            Else
                anySuccess = True
            End If
        Next
        
        If anyFailed Then
            MsgBox "Alguns arquivos tiveram falha no download!" + vbNewLine + "Atenção: Não é possível baixar comprovantes de operações com falha ou canceladas!", vbExclamation
        End If
        If anySuccess Then
            MsgBox "Arquivos salvos em:" + vbNewLine + folder
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

Public Function getFallbackName(row As Long, id As String, service As String)
    getFallbackName = id + ".pdf"
    If service = "transfer" Then
        getFallbackName = Format(ActiveSheet.Cells(row, 1), "yyyy-mm-dd") + " - " + Replace(ActiveSheet.Cells(row, 3).text, "R$", "R$ ") + " - " + CStr(ActiveSheet.Cells(row, 5).Value) + ".pdf"
    End If
End Function

Public Function ColumnId(service As String)
    
    Select Case service
        Case "transfer"
            Worksheets("Consulta de Transferências").Activate
            ColumnId = "B"
        Case "charge"
            Worksheets("Consulta de Boletos Emitidos").Activate
            ColumnId = "M"
        Case "invoice"
            Worksheets("Consulta de Invoices Emitidas").Activate
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