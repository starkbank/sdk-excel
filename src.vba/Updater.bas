Public Sub UpdateRoutine()
    Dim currentVersion As String
    Dim repoVersion As String
    Dim updateQuestion As String
    Dim updateAvailable As Boolean
    On Error GoTo Err
    
    currentVersion = getCurrentVersion()
    repoVersion = getLatestRepoVersion()
    
    Debug.Print "Current:", currentVersion
    Debug.Print "Remote:", repoVersion
    
    updateAvailable = SemVerCompare(currentVersion, repoVersion)
    If Not updateAvailable Then
        Exit Sub
    End If
    If AskUpdate(repoVersion) = vbNo Then
        MsgBox "Atualização cancelada."
        Exit Sub
    End If
    If RenameCurrent() = vbYes Then
        Update (repoVersion)
    Else
        MsgBox "Atualização cancelada."
    End If
    Exit Sub
Err:
    MsgBox "Não foi possível verificar atualizações. Favor checar a conexão."
    Exit Sub
End Sub

Public Function OpensslRoutine()
    OpensslRoutine = False
    If getGitDir() = "" Then
        ViewOpensslForm.Show
        Exit Function
    End If
    If dir(getOpensslDir()) = "" Then
        ViewOpensslForm.Show
        Exit Function
    End If
    OpensslRoutine = True
End Function

Public Function UpdateVersionNumber(newVersion As String)
    Sheets("Credentials").Cells(9, 2) = newVersion
    Sheets("Principal").Cells(1, 7) = "SDK Stark Bank - " & newVersion
End Function

Public Function SemVerCompare(currentVersion As String, repoVersion As String)
    Dim allMatches1 As Object
    Dim allMatches2 As Object
    Dim exp As String
    
    Dim currentMajor As Long
    Dim currentMinor As Long
    Dim currentPatch As Long
    
    Dim repoMajor As Long
    Dim repoMinor As Long
    Dim repoPatch As Long
    
    Dim majorBool As Boolean
    Dim minorBool As Boolean
    Dim patchBool As Boolean
    
    On Error Resume Next
    
    exp = "(v?)(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-((?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*)(?:\.(?:0|[1-9]\d*|\d*[a-zA-Z-][0-9a-zA-Z-]*))*))?(?:\+([0-9a-zA-Z-]+(?:\.[0-9a-zA-Z-]+)*))?$"
    With CreateObject("VBScript.RegExp")
        .Pattern = exp
        .Global = True
        .IgnoreCase = True
        Set allSubMatches1 = .Execute(currentVersion)(0).submatches
        Set allSubMatches2 = .Execute(repoVersion)(0).submatches
    End With
    currentMajor = CLng(allSubMatches1(1))
    currentMinor = CLng(allSubMatches1(2))
    currentPatch = CLng(allSubMatches1(3))
    
    repoMajor = CLng(allSubMatches2(1))
    repoMinor = CLng(allSubMatches2(2))
    repoPatch = CLng(allSubMatches2(3))
    
    majorBool = (repoMajor > currentMajor)
    minorBool = (repoMajor = currentMajor) And (repoMinor > currentMinor)
    patchBool = (repoMajor = currentMajor) And (repoMinor = currentMinor) And (repoPatch > currentPatch)
    
    SemVerCompare = False
    If majorBool Or minorBool Or patchBool Then
        SemVerCompare = True
    End If
End Function

Public Sub Update(repoVersion As String)
    Dim successMessage As String
    Dim overwriteLatest As Boolean
    
    DownloadLatest
    OpenDownloaded
    CloseCurrent
    UpdateVersionNumber (repoVersion)
End Sub

Public Function AskUpdate(repoVersion As String)
    updateQuestion1 = "Versão nova disponível para atualização: " + repoVersion + vbNewLine
    updateQuestion2 = "Atualizar agora? Você precisará transferir manualmente seus dados para a nova versão!"
    AskUpdate = MsgBox(updateQuestion1 + updateQuestion2, vbYesNo)
End Function

Private Function RenameCurrent()
    Dim fso As New Scripting.FileSystemObject
    Dim baseName As String
    Dim newName As String
    Dim oldName As String
    Dim overwriteLatest As Integer
    Dim overwriteMessage As String
    
    oldName = ThisWorkbook.FullName
    
    baseName = fso.GetBaseName(ThisWorkbook.name)
    overwriteLatest = vbYes
    Debug.Print dir(ThisWorkbook.path & "/starkbank-sdk.xlsm")
    If baseName = "starkbank-sdk" Then
        newName = baseName + "_OLD.xlsm"
        ThisWorkbook.SaveAs ThisWorkbook.path & "/" & newName
        Kill oldName
    ElseIf dir(ThisWorkbook.path & "/starkbank-sdk.xlsm") <> "" Then
        overwriteMessage = "O arquivo starkbank-sdk.xlsm será sobrescrito. Dados na planilha serão excluídos. Continuar?"
        overwriteLatest = MsgBox(overwriteMessage, vbYesNo)
    End If
    RenameCurrent = overwriteLatest
End Function

Public Sub DownloadLatest()
    Dim path As String: path = "https://github.com/starkbank/sdk-excel/raw/master/starkbank-sdk.xlsm"
    Dim targetFile As String: targetFile = TempFilePath()
    Dim oStream As Object
    Dim content As String
    Dim WinHttpReq As Object
    
    Set WinHttpReq = CreateObject("Msxml2.ServerXMLHTTP")
    WinHttpReq.Open "GET", path
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile targetFile, 2
        oStream.Close
    End If
End Sub

Private Sub OpenDownloaded()
    Workbooks.Open ThisWorkbook.path & "/starkbank-sdk.xlsm"
End Sub

Private Sub CloseCurrent()
    ThisWorkbook.Close False
End Sub

Public Sub UpdateOld()
    Dim successMessage As String
    DownloadLatest
    UpdateModules
    successMessage = "Atualização concluída."
    MsgBox successMessage, , "Sucesso"
End Sub

Public Sub UpdateModules()
    Dim VBProjTo As VBIDE.VBProject
    Dim VBProjFrom As VBIDE.VBProject
    
    Set VBProjTo = ActiveWorkbook.VBProject
    path = Application.ActiveWorkbook.path
    fromname = TempFilePath()
    
    Workbooks.Open fromname
    Set VBProjFrom = ActiveWorkbook.VBProject
    For Each comp In VBProjFrom.VBComponents
        compType = ComponentTypeToString(comp.Type)
        CompName = comp.name
        validName = (CompName <> "Updater") And (CompName <> "JsonConverter") And (CompName <> "Response")
        If (compType = "Code Module" Or compType = "UserForm" Or compType = "Class Module") And validName Then
            Debug.Print "Copying to current sheet: " & comp.name
            Debug.Print CopyModule(comp.name, VBProjFrom, VBProjTo, True)
        End If
    Next
    ActiveWorkbook.Close False
    
End Sub

Private Function TempFilePath()
    TempFilePath = ThisWorkbook.path & "/" + "starkbank-sdk.xlsm"
End Function

Public Function getLatestRepoVersion()
    Dim path As String
    Dim query As String
    Dim getTags As Variant
    
    path = baseRepoPath()
    query = "/tags"
    
    Set resp = getRepoRequest("api", path, query, New Dictionary)
    
    If resp.Status = 200 Then
        Set getTags = resp.json()
    Else
        MsgBox resp.error()("message"), , "Erro"
    End If
    
    getLatestRepoVersion = getTags(1)("name")
End Function

Public Function getLatestRepoFile()
    Dim path As String
    Dim query As String
    Dim content As String
    Dim getTags As Variant
    
    path = baseRepoPath()
    path = "/starkbank/sdk-excel"
    query = "/raw/master/SDK%20Stark%20Bank.xlsb.xlsm"
    
    Set resp = getRepoRequest("", path, query, New Dictionary)
    
    If resp.Status = 200 Then
        content = resp.content
    Else
        MsgBox resp.error()("message"), , "Erro"
    End If
    
    getLatestRepoFile = content
End Function

Public Function getCurrentVersion()
    getCurrentVersion = CStr(Sheets("Credentials").Cells(9, 2))
End Function

Public Function getRepoRequest(base As String, path As String, query As String, headers As Dictionary)
    Dim url As String: url = baseRepoUrl(base) + path + query
    Set getRepoRequest = updateFetch(url, "GET", headers, "")
End Function

Public Function updateFetch(url As String, method As String, headers As Dictionary, payload As String)
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open method, url, False
    
    For Each key In headers.keys()
        objHTTP.setRequestHeader key, headers(key)
    Next
    
    objHTTP.send payload
    
    Dim resp As response
    Set resp = New response
    
    resp.Status = objHTTP.Status
    resp.content = objHTTP.responseText
    
    Set updateFetch = resp

End Function


Public Function baseRepoUrl(repoString As String)
    Select Case repoString
        Case "api"
            baseRepoUrl = "https://api.github.com/repos"
        Case Else
            baseRepoUrl = "https://github.com"
    End Select
End Function

Public Function baseRepoPath()
    baseRepoPath = "/starkbank/sdk-excel"
End Function

Public Sub CopyTestModule()
    Dim moduleName As String
    Dim fromname As String
    Dim path As String
    Dim VBProjTo As VBIDE.VBProject
    Dim VBProjFrom As VBIDE.VBProject
    
    Set VBProjTo = ActiveWorkbook.VBProject
    path = Application.ActiveWorkbook.path
    fromname = TempFilePath()
    moduleName = "TestModule"
    
    Workbooks.Open fromname
    Set VBProjFrom = ActiveWorkbook.VBProject
    Debug.Print CopyModule(moduleName, VBProjFrom, VBProjTo, True)
    ActiveWorkbook.Close (False)
    ImportedTest moduleName
End Sub

Public Sub ImportedTest(moduleName As String)
    Dim zero As Integer
    zero = SimpleTest()
    ActiveWorkbook.VBProject.VBComponents.Remove ActiveWorkbook.VBProject.VBComponents(moduleName)
End Sub

Private Sub ListModules()
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    Set VBProj = ActiveWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        Debug.Print VBComp.name, ComponentTypeToString(VBComp.Type)
    Next VBComp
End Sub

Private Function ComponentTypeToString(ComponentType As VBIDE.vbext_ComponentType) As String
    Select Case ComponentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(ComponentType)
    End Select
End Function
