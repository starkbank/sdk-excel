Public Function headers(sheetName As String)

    If sheetName = "" Then
        sheetName = ActiveSheet.name
    End If
    
    Dim columns As New Collection
    For Each elem In Sheets(sheetName).UsedRange.columns
        columns.Add Sheets(sheetName).Cells(TableFormat.HeaderRow(), elem.column).Value
    Next
    Set headers = columns
End Function

Public Function dict(sheetName As String)
    Dim Result As New Collection
    Dim keys As Collection
    
    If sheetName = "" Then
        sheetName = ActiveSheet.name
    End If
    Set keys = headers(sheetName)
    
    For row = HeaderRow() + 1 To Sheets(sheetName).Cells(Rows.Count, "A").End(xlUp).row
        Dim obj As Object
        Set obj = JsonConverter.ParseJson("{}")
        For Each elem In Sheets(sheetName).UsedRange.columns
            obj(keys(elem.column)) = Sheets(sheetName).Cells(row, elem.column).text
        Next
        Result.Add obj
    Next
    
    Set dict = Result
End Function

Public Function longDict(initRow, lastRow)
    Dim Result As New Collection
    Dim keys As Collection
    
    Set keys = headers()
    
    For row = initRow To lastRow
        Dim obj As Object
        Set obj = JsonConverter.ParseJson("{}")
        For Each elem In ActiveSheet.UsedRange.columns
            obj(keys(elem.column)) = ActiveSheet.Cells(row, elem.column).text
        Next
        Result.Add obj
    Next
    
    Set longDict = Result
End Function