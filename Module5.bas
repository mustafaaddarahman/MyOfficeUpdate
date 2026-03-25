Attribute VB_Name = "Module5"
Sub emad_pro()

'============================================= «·Õ”«»«  «·ﬂ·ÌÂ
Dim item2 As ListItem
Dim last2, frw2 As Integer
Dim simadd As Double
Dim simsubtract As Double
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
On Error Resume Next
With UserForm1.ListView1
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))




item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(2) = Sheets(1).Cells(frw2, "c")
item2.SubItems(3) = Sheets(1).Cells(frw2, "d")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(7) = Sheets(1).Cells(frw2, "h")
item2.SubItems(8) = Sheets(1).Cells(frw2, "i")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(10) = Sheets(1).Cells(frw2, "k")
item2.SubItems(11) = Sheets(1).Cells(frw2, "l")
item2.SubItems(12) = Sheets(1).Cells(frw2, "m")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(14) = Sheets(1).Cells(frw2, "o")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(16) = Sheets(1).Cells(frw2, "q")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(21) = Sheets(1).Cells(frw2, "v")
item2.SubItems(22) = Sheets(1).Cells(frw2, "w")
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(25) = Sheets(1).Cells(frw2, "z")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(1).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(1).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(38) = Sheets(1).Cells(frw2, "am")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(40) = Sheets(1).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(1).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(1).Cells(frw2, "as")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(46) = Sheets(1).Cells(frw2, "au")
item2.SubItems(47) = Sheets(1).Cells(frw2, "av")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(1).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(1).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(1).Cells(frw2, "az")
item2.SubItems(52) = Sheets(1).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(1).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(1).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(1).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(1).Cells(frw2, "be")
item2.SubItems(57) = Sheets(1).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(1).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(1).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(1).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(1).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(1).Cells(frw2, "bk")
item2.SubItems(63) = Sheets(1).Cells(frw2, "bl")

End With
Next

Dim item3 As ListItem
Dim last3, frw3 As Integer
last3 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw3 = 9 To last3
simadd = Application.WorksheetFunction.Sum(Sheets(1).Range("bh9:bh" & frw3))
simsubtract = Application.WorksheetFunction.Sum(Sheets(1).Range("bj9:bj" & frw3))
UserForm1.TextBox73.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("j9:j" & frw3))
UserForm1.TextBox74.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("ad9:ad" & frw3))
UserForm1.TextBox79.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("ae9:ae" & frw3))
UserForm1.TextBox80.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("av9:av" & frw3))
UserForm1.TextBox66.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("aw9:aw" & frw3))
UserForm1.TextBox76.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("x9:x" & frw3))

UserForm1.TextBox66.Text = (TextBox66.Text + simadd) - simsubtract
UserForm1.TextBox75.Text = Application.WorksheetFunction.count(Sheets(1).Range("a9:a" & frw3))


UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")
Next

End Sub
