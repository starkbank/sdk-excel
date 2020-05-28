
Private Sub AfterTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    AfterTextBox.text = Utils.formatDateInUserForm(AfterTextBox.text)
    reentry = False
End Sub

Private Sub BeforeTextBox_Change()
    Static reentry As Boolean
    If reentry Then Exit Sub
    
    reentry = True
    BeforeTextBox.text = Utils.formatDateInUserForm(BeforeTextBox.text)
    reentry = False
End Sub

Private Sub UserForm_Initialize()
    Me.OptionButtonEventCredited.Enabled = True
    
    Me.AfterTextBox.Value = InputLogGateway.getAfterDate()
    Me.BeforeTextBox.Value = InputLogGateway.getBeforeDate()
End Sub

Private Sub SearchButton_Click()
    ' On Error Resume Next
    Dim afterInput As String: afterInput = AfterTextBox.Value
    Dim beforeInput As String: beforeInput = BeforeTextBox.Value
    
    Dim after As String: after = Utils.DateToSendingFormat(afterInput)
    Dim before As String: before = Utils.DateToSendingFormat(beforeInput)
    
    Dim id As String
    Dim amount As Double
    Dim fee As Double
    Dim issueDate As String
    Dim eventDate As String
    Dim chargeStatus As String
    Dim logEvent As String
    Dim dueDate As String
    Dim tags As Collection
    
    Dim cursor As String
    Dim charge As Object
    Dim chargeLog As Object
    Dim charges As Collection
    Dim chargeLogs As Collection
    Dim row As Long
    Dim logRow As Long
    Dim optionalParam As Dictionary: Set optionalParam = New Dictionary
    Dim rng As Range
    
    Dim logsPaidByCharge As Dictionary: Set logsPaidByCharge = New Dictionary
    Dim logsRegisteredByCharge As Dictionary: Set logsRegisteredByCharge = New Dictionary
    
    Call InputLogGateway.saveDates(afterInput, beforeInput)
    
    'Table layout
    Utils.applyStandardLayout ("V")
    ActiveSheet.Hyperlinks.Delete
    Range("A" & CStr(TableFormat.HeaderRow() + 1) & ":V" & Rows.Count).ClearContents
    
    'Headers definition
    
    ActiveSheet.Cells(TableFormat.HeaderRow(), 1).Value = "Data do Evento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 2).Value = "Evento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 3).Value = "Nome"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 4).Value = "CPF/CNPJ"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 5).Value = "Valor"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 6).Value = "Valor de Emissão"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 7).Value = "Desconto"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 8).Value = "Multa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 9).Value = "Juros"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 10).Value = "Data de Emissão"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 11).Value = "Vencimento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 12).Value = "Linha Digitável"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 13).Value = "Id do Boleto"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 14).Value = "Tarifa"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 15).Value = "Tags"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 16).Value = "Link PDF"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 17).Value = "Logradouro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 18).Value = "Complemento"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 19).Value = "Bairro"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 20).Value = "Cidade"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 21).Value = "Estado"
    ActiveSheet.Cells(TableFormat.HeaderRow(), 22).Value = "CEP"
    
    Call FreezeHeader
    
    'Optional parameters
    
    If after <> "--" Then
        optionalParam.Add "after", after
    End If
    If before <> "--" Then
        optionalParam.Add "before", before
    End If
    
    Dim Events As String: Events = ""
'    If OptionButtonEventCreated.Value Then
'        Events = Events + "register,"
'    End If
'    If OptionButtonEventRegistered.Value Then
'        Events = Events + "registered,"
'    End If
    If OptionButtonEventCanceled.Value Then
        Events = Events + "canceled,"
    End If
    If OptionButtonEventOverdue.Value Then
        Events = Events + "overdue,"
    End If
    If OptionButtonEventCredited.Value Then
        Events = Events + "bank,"
    End If
    
    row = TableFormat.HeaderRow() + 1
    If Events <> "bank," Then
        optionalParam.Add "events", Events
        Do
            Set respJson = getChargeLogs(cursor, optionalParam)
            cursor = ""
            If respJson("cursor") <> "" Then
                cursor = respJson("cursor")
            End If
    
            Set logs = respJson("logs")
    
            For Each chargeLog In logs
                
                Set charge = chargeLog("charge")
                
                logEvent = chargeLog("event")
                id = charge("id")
                amount = charge("amount") / 100
                fee = charge("fee") / 100
                eventDate = chargeLog("created")
                issueDate = charge("issueDate")
                chargeStatus = charge("status")
                dueDate = charge("dueDate")
                Set tags = charge("tags")
                
                ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(eventDate)
                ActiveSheet.Cells(row, 2).Value = ChargeGateway.getEventInPt(logEvent)
                ActiveSheet.Cells(row, 3).Value = charge("name")
                ActiveSheet.Cells(row, 4).Value = charge("taxId")
                ActiveSheet.Cells(row, 5).Value = amount
    
                ActiveSheet.Cells(row, 10).Value = Utils.ISODATEZ(issueDate)
                ActiveSheet.Cells(row, 11).Value = Utils.ISODATEZ(dueDate)
                ActiveSheet.Cells(row, 12).Value = charge("line")
                ActiveSheet.Cells(row, 13).Value = id
                ActiveSheet.Cells(row, 14).Value = fee
                ActiveSheet.Cells(row, 15).Value = CollectionToString(tags, ",")
    
                ActiveSheet.Cells(row, 16).Value = "PDF"
                Set rng = ActiveSheet.Range("P" + CStr(row))
                rng.Parent.Hyperlinks.Add Anchor:=rng, address:=StarkBankApi.baseUrl() + "/v1/charge/" + charge("id") + "/pdf", SubAddress:="", TextToDisplay:="PDF"
                row = row + 1
            Next
        Loop While cursor <> ""
        Unload Me
        Exit Sub
    End If
    optionalParam.Add "events", Events
    Do
        logRow = row
        Set respJson = getChargeLogs(cursor, optionalParam)
        If respJson.Exists("error") Then
            Unload Me
            Exit Sub
        End If

        cursor = ""
        If respJson("cursor") <> "" Then
            cursor = respJson("cursor")
        End If

        Set logs = respJson("logs")

        For Each chargeLog In logs
            Set charge = chargeLog("charge")
            
            eventDate = chargeLog("created")
            logEvent = chargeLog("event")
            id = charge("id")
            amount = charge("amount") / 100
            fee = charge("fee") / 100
            issueDate = charge("issueDate")
            chargeStatus = charge("status")
            dueDate = charge("dueDate")
            Set tags = charge("tags")

            ActiveSheet.Cells(row, 1).Value = Utils.ISODATEZ(eventDate)
            ActiveSheet.Cells(row, 2).Value = ChargeGateway.getEventInPt(logEvent)
            ActiveSheet.Cells(row, 3).Value = charge("name")
            ActiveSheet.Cells(row, 4).Value = charge("taxId")
            
            ActiveSheet.Cells(row, 5).Value = amount

            ActiveSheet.Cells(row, 10).Value = Utils.ISODATEZ(issueDate)
            ActiveSheet.Cells(row, 11).Value = Utils.ISODATEZ(dueDate)
            ActiveSheet.Cells(row, 12).Value = charge("line")
            ActiveSheet.Cells(row, 13).Value = id
            ActiveSheet.Cells(row, 14).Value = fee
            ActiveSheet.Cells(row, 15).Value = CollectionToString(tags, ",")

            ActiveSheet.Cells(row, 16).Value = "PDF"
            Set rng = ActiveSheet.Range("P" + CStr(row))
            rng.Parent.Hyperlinks.Add Anchor:=rng, address:=StarkBankApi.baseUrl() + "/v1/charge/" + charge("id") + "/pdf", SubAddress:="", TextToDisplay:="PDF"

            ActiveSheet.Cells(row, 17).Value = charge("streetLine1")
            ActiveSheet.Cells(row, 18).Value = charge("streetLine2")
            ActiveSheet.Cells(row, 19).Value = charge("district")
            ActiveSheet.Cells(row, 20).Value = charge("city")
            ActiveSheet.Cells(row, 21).Value = charge("stateCode")
            ActiveSheet.Cells(row, 22).Value = charge("zipCode")

            If chargeStatus = "paid" Then
                logsPaidByCharge.Add id, chargeLog
                logsRegisteredByCharge.Add id, New Dictionary
            End If

            row = row + 1
        Next

        Dim logsParam As Dictionary
        Dim keys As String
        Dim sep As String
        Dim registeredCursor As String: registeredCursor = ""
        keys = ""
        sep = ""
        Set logsParam = New Dictionary
        
        For Each chargeId In logsPaidByCharge.keys()
            keys = keys + sep + chargeId
            sep = ","
        Next
        logsParam.Add "chargeIds", keys
        
        Do
            logsParam("events") = "register"
            logsParam("cursor") = registeredCursor
            Set respJson = getChargeLogs("", logsParam)
            If respJson.Exists("error") Then
                MsgBox "Erro ao obter dados detalhados de boletos registrados!"
                Exit Sub
            End If
            If respJson("cursor") <> "" Then
                registeredCursor = respJson("cursor")
            End If
            Set registeredLogs = respJson("logs")
            For Each registeredLog In registeredLogs
                Set logsRegisteredByCharge(registeredLog("charge")("id")) = registeredLog
            Next
        Loop While registeredCursor <> ""
        
        For Each chargeLog In logs
            Set charge = chargeLog("charge")
            If charge("status") = "paid" Then
                Call setChargeEventInfo(charge, logsPaidByCharge(charge("id")), logsRegisteredByCharge(charge("id")), logRow)
            End If
            logRow = logRow + 1
        Next

        Set logsPaidByCharge = New Dictionary
        Set logsRegisteredByCharge = New Dictionary
        
    Loop While cursor <> ""
    
    Unload Me
End Sub

Public Sub setChargeEventInfo(ByVal charge As Object, ByVal paidLog As Dictionary, ByVal createdLog As Dictionary, row As Long)
    Dim nominalAmount As Double
    Dim fine As Double
    Dim amount As Double
    Dim interest As Double
    Dim discount As Double
    Dim deltaAmount As Double
    Dim logs As Collection
    Dim id As String: id = charge("id")
    
    amount = charge("amount") / 100
    
    nominalAmount = createdLog("charge")("amount") / 100
    deltaAmount = amount - nominalAmount
    
    ActiveSheet.Cells(row, 6).Value = nominalAmount
    If deltaAmount < 0 Then
        discount = deltaAmount
        ActiveSheet.Cells(row, 7).Value = discount
        ActiveSheet.Cells(row, 7).Font.Color = RGB(180, 0, 0)
    End If
    If deltaAmount > 0 Then
        fine = charge("fine") / 100 * nominalAmount
        interest = amount - fine - nominalAmount
        ActiveSheet.Cells(row, 8).Value = fine
        ActiveSheet.Cells(row, 9).Value = interest
        ActiveSheet.Cells(row, 8).Font.Color = RGB(0, 140, 0)
        ActiveSheet.Cells(row, 9).Font.Color = RGB(0, 140, 0)
    End If
End Sub


