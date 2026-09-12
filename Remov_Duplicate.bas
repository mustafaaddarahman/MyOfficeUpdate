Attribute VB_Name = "Remov_Duplicate"
Sub Remove_Duplicates1()

Dim lastR1 As Integer
Dim i As Integer
For i = 9 To lastR1
If Sheets(35).Cells(i, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(i, "ce").value = "ÇÖÇÝÉ" Then
lastR1 = Sheets(35).Cells(rowS.count, 1).End(xlUp).row
Range("a8:ce" & lastR1).RemoveDuplicates Columns:=Array(2, 5), header:=xlYes

ElseIf Sheets(35).Cells(i, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(i, "ce").value = "ÍÐÝ" Then
lastR1 = Sheets(35).Cells(rowS.count, 1).End(xlUp).row
Range("a8:ce" & lastR1).RemoveDuplicates Columns:=Array(2, 5), header:=xlYes

End If
Next





End Sub
