
Sub create_new_tabs()

  Dim i As Long
  Dim list As Worksheet
  Dim prevList As Worksheet
  Dim prevListName As String
  Dim thisListName As String
  Dim arrRangesPrev() As String
  Dim arrRanges() As String
  

  Set list = ActiveSheet
  thisListName = list.Name
  
  Set prevList = list.Previous
  prevListName = prevList.Name
  
  arrRangesPrev = Split(prevList.Name, "_")
  arrRanges = Split(list.Name, "_")

If list.Next Is Nothing Then

    prevList.Copy after:=ActiveSheet
    ActiveSheet.Name = arrRangesPrev(0) & "_" & arrRangesPrev(1) + 1
    Set prevList = ActiveSheet
    list.Copy after:=ActiveSheet
    ActiveSheet.Name = arrRanges(0) & "_" & arrRanges(1) + 1
    Set list = ActiveSheet
        
    list.Cells.Replace prevListName, prevList.Name
    prevList.Cells.Replace thisListName, list.Name
    prevList.Select

Else

    MsgBox ("New tabs have been created")
End If

End Sub