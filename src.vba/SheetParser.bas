Public Function headers()
    Dim columns As New Collection
    For Each elem In ActiveSheet.UsedRange.columns
        columns.Add ActiveSheet.Cells(9, elem.column).value
    Next
    Set headers = columns
End Function

Public Function dict()
    Dim result As New Collection
    Dim keys As Collection
    
    Set keys = headers()
    
    For row = 10 To ActiveSheet.Cells(Rows.Count, "A").End(xlUp).row
        Dim obj As Object
        Set obj = JsonConverter.ParseJson("{}")
        For Each elem In ActiveSheet.UsedRange.columns
            obj(keys(elem.column)) = ActiveSheet.Cells(row, elem.column).text
        Next
        result.Add obj
    Next
    
    Set dict = result
End Function