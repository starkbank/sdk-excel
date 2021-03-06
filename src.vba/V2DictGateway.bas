Public Sub ListKeys()
    
    If Not isSignedin Then
        MsgBox "Acesso negado. Faça login novamente.", , "Erro"
        Exit Sub
    End If
    
    Dim dictKeys As Collection
    Dim resp As response
    Dim initRow As Long
    Dim row As Long
    Dim lastRow As Long
    Dim dictKeyCount As Long
    Dim respMessage As String
    
    Call Utils.applyStandardLayout("D")
    Call Utils.applyLockedLayout("E", "J")
    
    'Headers definition
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Chave Pix"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Tags"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "Descrição"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "ISPB"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Agência"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Conta"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Tipo de Conta"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "externalId"
    
    With ActiveWindow
        .SplitColumn = 10
        .SplitRow = 9
    End With
    ActiveWindow.FreezePanes = True
    
    initRow = HeaderRow() + 1
    row = initRow
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
    
    dictKeyCount = lastRow - initRow + 1
    If dictKeyCount > 50 Then
        If MsgBox("Por restrição do Banco Central, só é possível consultar um máximo de 50 Chaves Pix sem finalizar o pagamento a cada 150 minutos. Apenas as primeiras 50 Chaves serão consultadas. Continuar?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    lastRow = IIf(dictKeyCount > 50, 50, dictKeyCount) + initRow - 1
    
    Dim errorMessage As String: errorMessage = ""
    For Each obj In SheetParser.longDict(initRow, lastRow)
        Dim keyId As String
        Dim dictKey As Variant
        keyId = obj("Chave Pix")
        If ActiveSheet.Cells(row, 11).Value <> keyId Then
            Set resp = V2DictGateway.getDictKey(keyId)
            If resp.Status >= 300 Then
                errorMessage = errorMessage & vbNewLine & "Linha " & row & ": " & resp.errors()("errors")(1)("message")
                ActiveSheet.Cells(row, 5).Value = ""
                ActiveSheet.Cells(row, 6).Value = ""
                ActiveSheet.Cells(row, 7).Value = ""
                ActiveSheet.Cells(row, 8).Value = ""
                ActiveSheet.Cells(row, 9).Value = ""
                ActiveSheet.Cells(row, 10).Value = ""
            Else
                Set dictKey = resp.json()("key")
                ActiveSheet.Cells(row, 5).Value = dictKey("name")
                ActiveSheet.Cells(row, 6).Value = dictKey("taxId")
                ActiveSheet.Cells(row, 7).Value = dictKey("ispb")
                ActiveSheet.Cells(row, 8).Value = dictKey("branchCode")
                ActiveSheet.Cells(row, 9).Value = dictKey("accountNumber")
                ActiveSheet.Cells(row, 10).Value = dictKey("accountType")
            End If
        End If
        ActiveSheet.Cells(row, 11).Value = keyId
        
        row = row + 1
    Next
    If errorMessage <> "" Then
        MsgBox errorMessage
    End If
    
    Call moveToTransfer
End Sub

Public Sub moveToTransfer()
    Dim transfers As Collection: Set transfers = SheetParser.dict()
    Dim transfer As Variant
    Dim validTransfers As Collection: Set validTransfers = New Collection
    For Each transfer In transfers
        If transfer("Nome") <> "" And transfer("Tipo de Conta") <> "" Then
            validTransfers.Add transfer
        End If
    Next
    If validTransfers.Count = 0 Then
        MsgBox "Não há nenhuma Chave Pix válida para mover para a aba de Transferências com Aprovação", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Foram encontradas " & validTransfers.Count & " Chaves Pix válidas. Deseja mover para a aba de Transferências com Aprovação? Dados na aba de Transferências com Aprovação serão apagados.", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Worksheets("Transferências com Aprovação").Activate
    Sheets("Transferências com Aprovação").Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":I" & Rows.Count).ClearContents
    
    row = HeaderRow() + 1
    For Each transfer In validTransfers
        Sheets("Transferências com Aprovação").Cells(row, 1).Value = transfer("Nome")
        Sheets("Transferências com Aprovação").Cells(row, 2).Value = transfer("CPF/CNPJ")
        Sheets("Transferências com Aprovação").Cells(row, 3).Value = transfer("Valor")
        Sheets("Transferências com Aprovação").Cells(row, 4).Value = transfer("ISPB")
        Sheets("Transferências com Aprovação").Cells(row, 5).Value = transfer("Agência")
        Sheets("Transferências com Aprovação").Cells(row, 6).Value = transfer("Conta")
        Sheets("Transferências com Aprovação").Cells(row, 7).Value = transfer("Tipo de Conta")
        Sheets("Transferências com Aprovação").Cells(row, 8).Value = transfer("Tags")
        Sheets("Transferências com Aprovação").Cells(row, 9).Value = transfer("Descrição")
        row = row + 1
    Next
    
    MsgBox "Transferências movidas para a aba de Transferências com Aprovação"
End Sub

Public Function getDictKey(keyId As String)
    Set getDictKey = V2Rest.getRequest("/v2/dict-key/" & keyId, "", New Dictionary)
End Function

Public Sub clearDictKeys()
    Sheets("Consulta de Chaves Pix").Range("A10:Z" & Rows.Count).ClearContents
End Sub