Private registerNumber As Integer
Private formattedAmount As String
Private registeredAmount As String
Private today As String
Private bankCode As String

Private totalNumber As Long
Private totalAmount As Long

Private occurrenceDateDict As Dictionary
Private numberDict As Dictionary
Private amountDict As Dictionary

Public Sub ExportFile()
    On Error GoTo ExportFail
    Dim outputFile As Integer
    Dim lastRow As Integer
    Dim i As Integer
    
    initializeOccurrences
    today = Date
    today = dateFormatter(today, 1, 4, 9)
    
    lastRow = ActiveSheet.Range("A9").CurrentRegion.Rows.count + 8
    
    outputFileName = "CNAB400_" & "20" & dateFormatter(today, 5, 3, 1, "-") & ".RET"
    outputPath = Application.DefaultFilePath & outputFileName
    
    dialog = Application.GetSaveAsFilename(outputFileName, FileFilter:="Text Files (*.ret), *.RET")
    Call exportMessageCanceled(dialog)
    
    outputFile = 1
    Set occurrenceDateDict = New Dictionary
    Call getLogOccurrenceDates(lastRow, "paid")
    Call getLogOccurrenceDates(lastRow, "canceled")
    
    registerNumber = 1
    
    Open dialog For Output As #outputFile
    
    outputPrintHeader (outputFile)
    
    For i = 10 To lastRow
        registerNumber = registerNumber + 1
        Call outputPrintTransactionOne(outputFile, i)
    Next
    
    For Each key In numberDict
        totalNumber = totalNumber + numberDict(key)
        totalAmount = totalAmount + amountDict(key)
    Next
    
    registerNumber = registerNumber + 1
    outputPrintTrailler (outputFile)
    
    Close #outputFile
    Call exportMessageSuccess(dialog)
    
    DebugDict
    Exit Sub
ExportFail:
    Call exportMessageCanceled(False)
End Sub

Private Sub exportMessageSuccess(dialog As Variant)
    If dialog <> False Then
        MsgBox "Arquivo exportado com sucesso!", , "Sucesso"
    Else
        MsgBox "Arquivo não foi salvo", , "Erro ao salvar"
    End If
End Sub

Private Sub exportMessageCanceled(dialog As Variant)
    If dialog = False Then
        MsgBox "Falha ao salvar o arquivo!", vbExclamation, "Erro ao salvar"
        End
    End If
End Sub

Public Function TaxIdFormatting(taxId As String) As String
    taxId = Replace(taxId, ".", "")
    taxId = Replace(taxId, "/", "")
    taxId = Replace(taxId, "-", "")
    
    TaxIdFormatting = taxId
End Function

Public Function getTaxIdType(taxId As String) As String
    lenTaxId = Len(taxId)
    Select Case Len(taxId)
        Case 14
            idType = "01"
        Case 18
            idType = "02"
        Case Else
            idType = "99"
    End Select
    getTaxIdType = idType
End Function

Public Function ZeroPad(s As Variant, n As Integer) As String
    ZeroPad = Format(CStr(s), String(n, "0"))
End Function

Private Sub DebugDict()
    For Each key In numberDict
        If numberDict(key) > 0 Then
            Debug.Print "Ocorr.:", key, "Quant.:", numberDict(key), "Valor:", amountDict(key)
        End If
    Next
    Debug.Print "Total Number:", totalNumber
    Debug.Print "Total Amount:", totalAmount
End Sub

Private Function getAmountLong(amount As Variant)
    amount = FormatCurrency(amount, 2)
    amount = Replace(amount, ",", "")
    amount = Replace(amount, ".", "")
    amount = Replace(amount, "R$", "")
    amount = CLng(amount)
    getAmountLong = amount
End Function

Private Function getLogOccurrenceDate(statusCode As String, chargeId As String) As String
    Dim respMessage As Variant
    Dim logevent As String
    Set respMessage = ChargeGateway.getChargeLog(chargeId, New Dictionary)
    
    For Each elem In respMessage("logs")
        logevent = elem("event")
        If (statusCode = ChargeGateway.getStatusInPt(logevent)) Then
            getLogOccurrenceDate = elem("created")
            Exit Function
        End If
    Next
    
End Function

Private Sub getLogOccurrenceDates(lastRow As Integer, logevent As String)
    Dim chunk As String
    Dim respMessage As Dictionary
    Dim i As Integer
    Dim j As Integer
    Dim occurrenceId As String
    Dim statusCode As String
    Dim dictOccurrenceDate As Dictionary
    Set dictOccurrenceDate = New Dictionary
    
    chunk = ""
    j = 0
    For i = 10 To lastRow
        statusCode = Cells(i, "D").Value
        chargeId = CStr(Cells(i, "M").Value)
        
        occurrenceId = ChargeGateway.getOccurrenceId(statusCode)
        If ChargeGateway.getStatusFromId(occurrenceId) = logevent Then
            j = j + 1
            chunk = chunk & chargeId & ","
            If j >= 100 Then
                insertOccurrenceDict chunk, logevent
                chunk = ""
                j = 0
            End If
        End If
    Next
    If chunk <> "" Then
        insertOccurrenceDict chunk, logevent
    End If
End Sub

Private Sub insertOccurrenceDict(chunk As String, logevent As String)
    Set respMessage = ChargeGateway.getEventLog(chunk, logevent, New Dictionary)
    
    For Each elem In respMessage("logs")
        occurrenceDateDict.Add CStr(elem("charge")("id")), elem("created")
    Next
End Sub

Public Function dateFormatter(inputDate As String, p1 As Integer, p2 As Integer, p3 As Integer, Optional sep As String = "") As String
    dateFormatter = Mid(inputDate, p1, 2) & sep & Mid(inputDate, p2, 2) & sep & Mid(inputDate, p3, 2)
End Function

Private Sub initializeOccurrences()
    Dim key As String
    Set numberDict = New Dictionary
    Set amountDict = New Dictionary
    
    totalNumber = 0
    totalAmount = 0
    For i = 0 To 99
        key = ZeroPad(i, 2)
        numberDict.Add key, 0
        amountDict.Add key, 0
    Next
End Sub

Private Sub outputPrintHeader(outputFile As Integer)
    ' Registro Header do lote
    
    Dim workspaceId As String
    Dim companyName As String
    
    workspaceId = SessionGateway.getWorkspaceId()
    companyId = ZeroPad(workspaceId, 20)
    
    companyName = OwnerGateway.getOwnerName(workspaceId, New Dictionary)
    companyName = Left(UCase(companyName), 30)
    
    creditDate = "000000"
    formattedCreditDate = CStr(creditDate)
    
    Print #outputFile, Tab(1); "02RETORNO01COBRANCA";
    Print #outputFile, Tab(27); companyId;
    Print #outputFile, Tab(47); UCase(companyName);
    Print #outputFile, Tab(77); StarkData.GetBankNumber();
    Print #outputFile, Tab(80); StarkData.GetBankName();
    Print #outputFile, Tab(95); today; ' Data da gravação do arquivo
    Print #outputFile, Tab(101); "01600000";
    Print #outputFile, Tab(109); ZeroPad("360", 5);
    ' Em Branco: 114 a 379
    Print #outputFile, Tab(380); creditDate; ' TODO: Data do crédito
    ' Em Branco: 386 a 394
    Print #outputFile, Tab(395); ZeroPad(registerNumber, 6)
End Sub

Private Sub outputPrintTransactionOne(outputFile As Integer, i As Integer)
    Dim issueDate As String
    Dim creditDate As String
    Dim occurrenceDate As String
    Dim amount As String
    Dim dueDate As String
    Dim statusCode As String
    Dim companyName As String
    Dim taxId As String
    Dim chargeId As String
    Dim amountInt As Long
    
    issueDate = Cells(i, "A").Value
    amount = CStr(Cells(i, "E").Value)
    dueDate = Cells(i, "K").Value
    statusCode = Cells(i, "D").Value
    companyName = Cells(i, "B").Value
    taxId = Cells(i, "C").Value
    chargeId = CStr(Cells(i, "M").Value)
    
    taxId = CStr(taxId)
    typeTaxId = getTaxIdType(taxId)
    taxId = TaxIdFormatting(taxId)
    
    occurrenceId = ChargeGateway.getOccurrenceId(statusCode)
    
    numberDict(occurrenceId) = numberDict(occurrenceId) + 1
    
    amountInt = getAmountLong(amount)
    amountDict(occurrenceId) = amountDict(occurrenceId) + amountInt
    
    wallet = ZeroPad(StarkData.GetWallet(), 3)
    branch = ZeroPad(StarkData.GetBranch(), 4)
    accountNumber = ZeroPad(StarkData.GetAccountNumber(), 9)
    companySubscription = "0" & wallet & branch & accountNumber
    formattedAmount = ZeroPad(amountInt, 13)
    
    occurrenceDate = dateFormatter(issueDate, 1, 4, 9)
    formattedPaidAmount = ZeroPad("0", 13)
    creditDate = "      "
    bankCode = "    "
    
    If occurrenceId = "06" Or occurrenceId = "09" Then
        occurrenceDate = dateFormatter(occurrenceDateDict(chargeId), 9, 6, 3)
        If occurrenceId = "06" Then
            formattedPaidAmount = formattedAmount
            creditDate = occurrenceDate
            bankCode = "0000"
        End If
    End If
    registeredAmount = ZeroPad(amountInt, 12)
    
    formattedIssueDate = dateFormatter(issueDate, 1, 4, 9)
    formattedDueDate = dateFormatter(dueDate, 1, 4, 9)
    
    ' Registro de Transação
    Print #outputFile, Tab(1); "1";
    Print #outputFile, Tab(2); typeTaxId; ' TODO: Verificar CPF ou CNPJ, PIS/PASEP
    Print #outputFile, Tab(4); ZeroPad(taxId, 14); ' N. inscricao da Empresa
    Print #outputFile, Tab(18); "000"; ' Zeros
    Print #outputFile, Tab(21); companySubscription; ' Identificacao da Empresa no Banco (Zero, Carteira, Agencia e CC) Obs pag 45
    Print #outputFile, Tab(38); ZeroPad(chargeId, 25); ' "Numero controle do participante, uso da empresa"
    Print #outputFile, Tab(63); ZeroPad("0", 8);
    Print #outputFile, Tab(71); Right(chargeId, 12); ' TODO: Obs Pag 45
    Print #outputFile, Tab(83); ZeroPad("0", 10);
    Print #outputFile, Tab(93); ZeroPad("0", 12);
    Print #outputFile, Tab(105); "0"; ' TODO: Confirmar se Zero ao invés de R; Obs Pag 45
    Print #outputFile, Tab(106); "00"; ' Não foi informado parcelamento
    Print #outputFile, Tab(108); StarkData.GetWallet(); ' Carteira
    Print #outputFile, Tab(109); occurrenceId; ' Identificacao ocorrencia: Obs Pag 45
    Print #outputFile, Tab(111); occurrenceDate; ' Data ocorrencia no banco
    Print #outputFile, Tab(117); ZeroPad("0", 10); ' Numero do documento
    Print #outputFile, Tab(127); ZeroPad(chargeId, 20); ' TODO: Obs Pag 46
    Print #outputFile, Tab(147); formattedDueDate;
    Print #outputFile, Tab(153); formattedAmount;
    Print #outputFile, Tab(166); "341"; ' Codigo do Banco, Camara de Compensacao
    Print #outputFile, Tab(169); ZeroPad(branch, 5); ' TODO: Codigo da Agencia do banco cobrador";
    Print #outputFile, Tab(176); "0000000000000"; ' Despesas de coranca Tarifa de registro
    Print #outputFile, Tab(189); "0000000000000";
    Print #outputFile, Tab(202); "0000000000000"; ' Juros operacao em atraso
    Print #outputFile, Tab(215); "0000000000000"; ' Valor do IOF
    Print #outputFile, Tab(228); "0000000000000"; ' Valor abatido
    Print #outputFile, Tab(241); "0000000000000"; ' Desconto concedido
    Print #outputFile, Tab(254); formattedPaidAmount; ' Valor pago
    Print #outputFile, Tab(267); "0000000000000"; ' Juros de Mora
    Print #outputFile, Tab(280); "0000000000000";
    ' Em Branco: 293 a 294
    Print #outputFile, Tab(295); " "; ' A = Aceito, D = Desprezado
    Print #outputFile, Tab(296); creditDate; ' Data do credito;
    Print #outputFile, Tab(302); "   "; ' Origem pagamento Obs pag 46
    ' Em Branco: 305 a 314
    Print #outputFile, Tab(315); bankCode; ' Codigo do banco
    Print #outputFile, Tab(319); "0000000000"; ' Motivo das rejeicoes Obs pag 47
    ' Em Branco: 329 a 368
    'Print #outputFile, Tab(369); "00"; ' Numero do cartorio
    'Print #outputFile, Tab(371); ZeroPad("0", 10); ' TODO
    Print #outputFile, Tab(395); ZeroPad(registerNumber, 6)
End Sub

Private Sub outputPrintTrailler(outputFile As Integer)
    ' Registro Trailler
    Print #outputFile, Tab(1); "9";
    Print #outputFile, Tab(2); "2";
    Print #outputFile, Tab(3); "01";
    Print #outputFile, Tab(5); ZeroPad(StarkData.GetBankNumber(), 3);
    ' Em Branco: 008 a 017
    Print #outputFile, Tab(18); ZeroPad(totalNumber, 8); ' TODO: Quantidade de Titulos de cobranca
    Print #outputFile, Tab(26); ZeroPad(totalAmount, 14); ' TODO: Valor total em cobranca
    Print #outputFile, Tab(40); ZeroPad("360", 8); ; ' TODO: Aviso Bancário
    ' Em Branco: 048 a 057
    Print #outputFile, Tab(58); ZeroPad(numberDict("02"), 5); ' Quantidade de registros 02
    Print #outputFile, Tab(63); ZeroPad(amountDict("02"), 12); ' Valor dos registros ocorrencia 02
    Print #outputFile, Tab(75); ZeroPad(amountDict("06"), 12); ' Valor dos registros ocorrencia 06
    Print #outputFile, Tab(87); ZeroPad(numberDict("06"), 5); ' Quantidade de registros ocorrencia 06
    Print #outputFile, Tab(92); ZeroPad(amountDict("06"), 12); ; ' Valor dos registros ocorrencia 06
    Print #outputFile, Tab(104); ZeroPad(numberDict("09") + numberDict("10"), 5); ' Quantidade de registros 09 e 10
    Print #outputFile, Tab(109); ZeroPad(amountDict("09") + amountDict("10"), 12); ' Valor dos registros ocorrencia 09 e 10
    Print #outputFile, Tab(121); ZeroPad(numberDict("13"), 5); ' Quantidade de registros 13
    Print #outputFile, Tab(126); ZeroPad(amountDict("13"), 12); ' Valor dos registros ocorrencia 13
    Print #outputFile, Tab(138); ZeroPad(numberDict("14"), 5); ' Quantidade de registros 14
    Print #outputFile, Tab(143); ZeroPad(amountDict("14"), 12); ' Valor dos registros ocorrencia 14
    Print #outputFile, Tab(155); ZeroPad(numberDict("12"), 5); ' Quantidade de registros 12
    Print #outputFile, Tab(160); ZeroPad(amountDict("12"), 12); ' Valor dos registros ocorrencia 12
    Print #outputFile, Tab(172); ZeroPad(numberDict("19"), 5); ' Quantidade de registros 19
    Print #outputFile, Tab(177); ZeroPad(amountDict("19"), 12); ' Valor dos registros ocorrencia 19
    ' Em Branco: 189 a 362
    Print #outputFile, Tab(363); ZeroPad("0", 15); ' TODO: Valor Total Rateios
    Print #outputFile, Tab(378); ZeroPad("0", 8); ' TODO: Quantidade Total Rateios
    Print #outputFile, Tab(395); Format(CStr(registerNumber), "000000");
End Sub
