Public Function HeaderRow()
    HeaderRow = 9
End Function

Public Sub FreezeHeader()
    With ActiveWindow
    If .FreezePanes Then .FreezePanes = False
        .SplitRow = TableFormat.HeaderRow()
        .SplitColumn = 0
        .FreezePanes = True
    End With
End Sub