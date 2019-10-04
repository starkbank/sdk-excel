Private registerNumber As Integer
Private formattedAmount As String
Private registeredAmount As String
Private today As String
Private bankCode As String

Private totalNumber As Long
Private totalAmount As Long

Private numberDict As Dictionary
Private amountDict As Dictionary

Public Sub ExportFile()
    Dim outputFile As Integer
    Dim i As Integer
    
    InitializeOccurrences
    
    today = Date
    today = Mid(today, 1, 2) & Mid(today, 4, 2) & Mid(today, 9, 2)
    
    lastRow = ActiveSheet.Range("A9").CurrentRegion.Rows.Count + 8
    Debug.Print "The last row is: " + CStr(lastRow)
    
    outputFileName = "CNAB400_" & CStr(today) & ".RET"
    outputPath = Application.DefaultFilePath & outputFileName
    
    dialog = Application.GetSaveAsFilename(outputFileName, FileFilter:="Text Files (*.ret), *.RET")
    Call CanceledExportMessage(dialog)
    outputFile = 1
    Open dialog For Output As #outputFile
    
    registerNumber = 1
    
    OutputPrintHeader (outputFile)
    For i = 10 To lastRow
        registerNumber = registerNumber + 1
        Call OutputPrintTransactionOne(outputFile, i)
    Next
    For Each key In numberDict
        totalNumber = totalNumber + numberDict(key)
        totalAmount = totalAmount + amountDict(key)
    Next
    registerNumber = registerNumber + 1
    OutputPrintTrailler (outputFile)
    
    Close #outputFile
    Call SuccessExportMessage(dialog)
    DebugDict
End Sub

Private Sub SuccessExportMessage(dialog As Variant)
    If dialog <> False Then
        Cells(2, 1) = dialog
        MsgBox "Arquivo exportado com sucesso...", , "Sucesso"
    Else
        MsgBox "Arquivo não foi salvo", , "Erro ao salvar"
    End If
End Sub

Private Sub CanceledExportMessage(dialog As Variant)
    If dialog = False Then
        MsgBox "Arquivo não foi salvo", , "Erro ao salvar"
        End
    End If
End Sub
Public Function TaxIdFormatting(taxId As String) As String
    ' Verify and validade taxId
    If taxId = "" Then
        Debug.Print "No taxId to check"
    End If
        
    taxId = Replace(taxId, ".", "")
    taxId = Replace(taxId, "/", "")
    taxId = Replace(taxId, "-", "")
    
    TaxIdFormatting = taxId
End Function

Public Function GetTaxIdType(taxId As String) As String
    ' Only works with TaxIds containing punctuation
    lenTaxId = Len(taxId)
    Debug.Print taxId
    Select Case Len(taxId)
        Case 14
            idType = "01"
        Case 18
            idType = "02"
        Case Else
            Debug.Print "Error to verify taxId type: " & taxId
            idType = "99"
    End Select
    GetTaxIdType = idType
End Function

Public Function GetOccurrenceId(statusCode As String)
    Select Case statusCode
        Case "pendente de registro"
            occurrenceId = "00"
        Case "registrado"
            occurrenceId = "02"
        Case "vencido"
            occurrenceId = "02"
        Case "falha"
            occurrenceId = "03"
        Case "pago"
            occurrenceId = "06"
        Case "cancelado"
            occurrenceId = "09"
        Case Else
            occurrenceId = "99"
    End Select
    GetOccurrenceId = occurrenceId
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

Private Function GetAmountLong(amount As Variant)
    Debug.Print "Amount: " + amount
    amount = FormatCurrency(amount, 2)
    amount = Replace(amount, ",", "")
    amount = Replace(amount, ".", "")
    amount = Replace(amount, "R$", "")
    amount = CLng(amount)
    GetAmountLong = amount
End Function

Private Function GetLogOccurrenceDate(statusCode As String, chargeId As String) As String
    Dim respMessage As Variant
    Dim logEvent As String
    Set respMessage = ChargeGateway.getChargeLog(chargeId, New Dictionary)
    
    For Each elem In respMessage("logs")
        logEvent = elem("event")
        If (statusCode = ChargeGateway.getStatusInPt(logEvent)) Then
            GetLogOccurrenceDate = elem("created")
            Exit Function
        End If
    Next
    
End Function

Private Sub InitializeOccurrences()
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

Private Sub OutputPrintHeader(outputFile As Integer)
    ' Registro Header do lote
    
    Dim workspaceId As String
    Dim companyName As String
    
    workspaceId = SessionGateway.getWorkspaceId()
    companyId = ZeroPad(workspaceId, 20)
    
    companyName = OwnerGateway.getOwnerName(workspaceId, New Dictionary)
    companyName = Left(UCase(companyName), 30)
    
    creditDate = "000000" ' TODO: ????
    ' creditDate = Application.WorkDay(Date, 1) ' TODO: ????
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
    Print #outputFile, Tab(394); ZeroPad(registerNumber, 6)
End Sub

Private Sub OutputPrintTransactionOne(outputFile As Integer, i As Integer)
    Dim issueDate As String
    Dim amount As String
    Dim dueDate As String
    Dim statusCode As String
    Dim companyName As String
    Dim taxId As String
    Dim chargeId As String
    Dim amountInt As Long
    
    issueDate = Cells(i, "A").Value
    amount = CStr(Cells(i, "B").Value)
    dueDate = Cells(i, "C").Value
    statusCode = Cells(i, "D").Value
    companyName = Cells(i, "E").Value
    taxId = Cells(i, "F").Value
    chargeId = CStr(Cells(i, "H").Value)
    
    taxId = CStr(taxId)
    typeTaxId = GetTaxIdType(taxId)
    taxId = TaxIdFormatting(taxId)
    
    occurrenceId = GetOccurrenceId(statusCode)
    
    numberDict(occurrenceId) = numberDict(occurrenceId) + 1
    
    amountInt = GetAmountLong(amount)
    amountDict(occurrenceId) = amountDict(occurrenceId) + amountInt
    
    Debug.Print "chargeId", chargeId, CStr(VarType(chargeId))
    occurrenceDate = GetLogOccurrenceDate(statusCode, chargeId)
    
    If occurrenceId = "06" Then
        creditDate = occurrenceDate
        creditDate = Mid(creditDate, 9, 2) & Mid(creditDate, 6, 2) & Mid(creditDate, 3, 2)
        bankCode = "0000"
    Else
        creditDate = "      "
        bankCode = "    "
    End If
    
    wallet = ZeroPad(StarkData.GetWallet(), 3)
    branch = ZeroPad(StarkData.GetBranch(), 4)
    accountNumber = ZeroPad(StarkData.GetAccountNumber(), 9)
    companySubscription = "0" & wallet & branch & accountNumber
    formattedAmount = ZeroPad(amountInt, 13)
    If occurenceId = "06" Then
        formattedPaidAmount = formmattedAmount
    Else
        formattedPaidAmount = ZeroPad("0", 13)
    End If
    registeredAmount = ZeroPad(amountInt, 12)
    
    formattedIssueDate = Mid(issueDate, 1, 2) & Mid(issueDate, 4, 2) & Mid(issueDate, 9, 2)
    formattedDueDate = Mid(dueDate, 1, 2) & Mid(dueDate, 4, 2) & Mid(dueDate, 9, 2)
    occurrenceDate = Mid(occurrenceDate, 9, 2) & Mid(occurrenceDate, 6, 2) & Mid(occurrenceDate, 3, 2)
    
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
    Print #outputFile, Tab(117); Right(chargeId, 10); ' Numero do documento
    Print #outputFile, Tab(127); ZeroPad(documentNumber, 20); ' TODO: Obs Pag 46
    Print #outputFile, Tab(147); formattedDueDate;
    Print #outputFile, Tab(153); formattedAmount;
    Print #outputFile, Tab(166); "341"; ' Codigo do Banco, Camara de Compensacao
    Print #outputFile, Tab(169); ZeroPad(branch, 5); ' TODO: ???? Codigo da Agencia do banco cobrador";
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
    'Print #outputFile, Tab(371); ZeroPad("0", 10); ' TODO: ????
    Print #outputFile, Tab(394); ZeroPad(registerNumber, 6)
End Sub

Private Sub OutputPrintTrailler(outputFile As Integer)
    ' Registro Trailler
    Print #outputFile, Tab(1); "9";
    Print #outputFile, Tab(2); "2";
    Print #outputFile, Tab(3); "01";
    Print #outputFile, Tab(5); ZeroPad(StarkData.GetBankNumber(), 3);
    ' Em Branco: 008 a 017
    Print #outputFile, Tab(18); ZeroPad(totalNumber, 8); ' TODO: ???? Quantidade de Titulos de cobranca
    Print #outputFile, Tab(26); ZeroPad(totalAmount, 14); ' TODO: ???? Valor total em cobranca
    Print #outputFile, Tab(40); ZeroPad("360", 8); ; ' TODO: ???? Aviso Bancário
    ' Em Branco: 048 a 057
    Print #outputFile, Tab(58); ZeroPad(numberDict("02"), 5); ' Quantidade de registros 02
    Print #outputFile, Tab(63); ZeroPad(amountDict("02"), 12); ' Valor dos registros ocorrencia 02
    Print #outputFile, Tab(75); ZeroPad(amountDict("06"), 12); ' Valor dos registros ocorrencia 06
    Print #outputFile, Tab(87); ZeroPad(numberDict("06"), 5); ' Quantidade de registros ocorrencia 06
    Print #outputFile, Tab(92); ZeroPad(amountDict("06"), 12); ; ' ???? Valor dos registros ocorrencia 06
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
    Print #outputFile, Tab(363); ZeroPad("0", 15); ' TODO: ???? Valor Total Rateios
    Print #outputFile, Tab(378); ZeroPad("0", 8); ' TODO: ???? Quantidade Total Rateios
    Print #outputFile, Tab(394); Format(CStr(registerNumber), "000000");
End Sub