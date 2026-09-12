Attribute VB_Name = "Module6"
Sub trans_data()
 Sheets(11).Range("a2:z10000").value = ""
 'Sheets(11).Range("a1").Value = 0
  
Dim lr, frow As Double
lr = Sheets(11).Range("a" & Rows.count).End(xlUp).row + 1
With UserForm1.ListView1
For frow = 1 To UserForm1.ListView1.ListItems.count
Sheets(11).Cells(lr, "a").value = val(Sheets(11).Cells(lr - 1, "a").value) + 1
Sheets(11).Cells(lr, "b").value = UserForm1.ListView1.ListItems(frow).SubItems(1)
Sheets(11).Cells(lr, "c").value = UserForm1.ListView1.ListItems(frow).SubItems(2)
Sheets(11).Cells(lr, "d").value = UserForm1.ListView1.ListItems(frow).SubItems(3)
Sheets(11).Cells(lr, "e").value = UserForm1.ListView1.ListItems(frow).SubItems(4)
Sheets(11).Cells(lr, "f").value = UserForm1.ListView1.ListItems(frow).SubItems(5)
Sheets(11).Cells(lr, "g").value = UserForm1.ListView1.ListItems(frow).SubItems(6)
Sheets(11).Cells(lr, "h").value = UserForm1.ListView1.ListItems(frow).SubItems(7)
Sheets(11).Cells(lr, "i").value = UserForm1.ListView1.ListItems(frow).SubItems(8)
Sheets(11).Cells(lr, "j").value = UserForm1.ListView1.ListItems(frow).SubItems(9)
Sheets(11).Cells(lr, "k").value = UserForm1.ListView1.ListItems(frow).SubItems(10)
Sheets(11).Cells(lr, "l").value = UserForm1.ListView1.ListItems(frow).SubItems(11)
Sheets(11).Cells(lr, "m").value = UserForm1.ListView1.ListItems(frow).SubItems(12)
Sheets(11).Cells(lr, "n").value = UserForm1.ListView1.ListItems(frow).SubItems(13)
Sheets(11).Cells(lr, "o").value = UserForm1.ListView1.ListItems(frow).SubItems(14)
Sheets(11).Cells(lr, "p").value = UserForm1.ListView1.ListItems(frow).SubItems(15)
Sheets(11).Cells(lr, "q").value = UserForm1.ListView1.ListItems(frow).SubItems(16)
Sheets(11).Cells(lr, "r").value = UserForm1.ListView1.ListItems(frow).SubItems(17)
Sheets(11).Cells(lr, "s").value = UserForm1.ListView1.ListItems(frow).SubItems(18)
Sheets(11).Cells(lr, "t").value = UserForm1.ListView1.ListItems(frow).SubItems(19)
Sheets(11).Cells(lr, "u").value = UserForm1.ListView1.ListItems(frow).SubItems(20)
Sheets(11).Cells(lr, "v").value = UserForm1.ListView1.ListItems(frow).SubItems(21)
Sheets(11).Cells(lr, "w").value = UserForm1.ListView1.ListItems(frow).SubItems(22)
Sheets(11).Cells(lr, "x").value = UserForm1.ListView1.ListItems(frow).SubItems(23)
Sheets(11).Cells(lr, "y").value = UserForm1.ListView1.ListItems(frow).SubItems(26)
Sheets(11).Cells(lr, "z").value = UserForm1.ListView1.ListItems(frow).SubItems(25)

'Sheets(11).Cells(lr, "m").Value = UserForm1.ListView1.ListItems(frow).SubItems(1)
'Sheets(11).Cells(lr, "n").Value = UserForm1.ListView1.ListItems(frow).SubItems(1)
lr = lr + 1
Next frow
End With
Sheets(11).Range("a1") = "«·ﬂÊœ"
Sheets(11).Activate
End Sub

