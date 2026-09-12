Attribute VB_Name = "Module13"
Sub auto_num()
Dim i, lr As Integer
lr = Sheets(21).Cells(Rows.count, "a").End(xlUp).row
For i = 1 To lr - 1
Sheets(21).Cells(i + 1, "a") = i
Next
End Sub
