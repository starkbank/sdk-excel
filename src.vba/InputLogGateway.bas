Public Function saveDates(after As String, before As String)
    Sheets("InputLog").Cells(1, 1) = "AfterDate"
    Sheets("InputLog").Cells(1, 2) = after
    Sheets("InputLog").Cells(2, 1) = "BeforeDate"
    Sheets("InputLog").Cells(2, 2) = before
End Function

Public Function savePath(path As String)
    Sheets("InputLog").Cells(3, 1) = "PrivateKeyPath"
    Sheets("InputLog").Cells(3, 2) = path
End Function

Public Function getAfterDate()
    getAfterDate = CStr(Sheets("InputLog").Cells(1, 2))
End Function

Public Function getBeforeDate()
    getBeforeDate = CStr(Sheets("InputLog").Cells(2, 2))
End Function

Public Function getPath()
    getPath = Sheets("InputLog").Cells(3, 2)
End Function