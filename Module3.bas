Attribute VB_Name = "Module3"
Function rightToLeftListView()
Dim Columns As Long
Dim i As Long
With ListView1
Columns1 = ListView1.ColumnHeaders.count
For i1 = 1 To Columns
.ColumnHeaders(i1).Position = Columns1 - i1 + 1
Next
For i1 = 2 To Columns
.ColumnHeaders(i1).Alignment = lvwColumnRight
Next
End With
End Function





