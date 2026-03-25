Attribute VB_Name = "Remov_Duplicate"
Sub Remove_Duplicates1()

Dim lastr1 As Integer
Dim i As Integer
For i = 9 To lastr1
If Sheets(35).Cells(i, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(i, "ce").value = "ÇÖÇÝÉ" Then
lastr1 = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
Range("a8:ce" & lastr1).RemoveDuplicates Columns:=Array(2, 5), Header:=xlYes

ElseIf Sheets(35).Cells(i, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(i, "ce").value = "ÍÐÝ" Then
lastr1 = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
Range("a8:ce" & lastr1).RemoveDuplicates Columns:=Array(2, 5), Header:=xlYes

End If
Next





End Sub
