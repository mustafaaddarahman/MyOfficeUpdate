Attribute VB_Name = "Module4"
Function rightToLeftListView2()
Dim Columns As Long
Dim i As Long
With ListView2
Columns = .ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next
End With
End Function
