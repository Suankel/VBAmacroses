Sub delete_empty_lines()

 Dim row As Long
    Dim rowsRange As Range
    Dim rangeOfOneRow As Range
    
    Set rowsRange = Range("74:1300") ' you can choose another range
    
    For row = 1 To rowsRange.Rows.Count
        If Application.CountA(rowsRange.Rows(row)) = 0 Then
            If rangeOfOneRow Is Nothing Then
              Set rangeOfOneRow = rowsRange.Rows(row)
            Else
              Set rangeOfOneRow = Union(rangeOfOneRow, rowsRange.Rows(row))
            End If
        End If
    Next row
    If Not rangeOfOneRow Is Nothing Then rangeOfOneRow.Delete
MsgBox ("Empty lines have been deleted")
End Sub

