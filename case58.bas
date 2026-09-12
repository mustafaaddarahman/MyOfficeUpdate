Attribute VB_Name = "case58"
Sub case_58()
On Error Resume Next
Sheets("sheet28").Range("a9:bz8000").ClearContents
ActiveSheet.Range("a9: bz6000").UnMerge
Dim arr1, arr2, lastR As Variant, i As Long, j As Long, k As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets(13).Cells(Rows.count, 1).End(xlUp).row + 1
arr1 = Sheets(13).Range("a9:bz" & lastR)
j = 8
ReDim arr2(9 To UBound(arr1, 1), 1 To UBound(arr1, 2))

For i = 1 To UBound(arr1, 1)

 If arr1(i, 58) Like "*" & UserForm17.TextBox2.Text & "*" Then

 j = j + 1
 For k = 1 To 78
 arr2(j, k) = arr1(i, k)
  
  Next k
  
End If

Next


Sheets("sheet28").Range("a9").Resize(j - 1, 78).value = arr2
Sheets("sheet28").Activate
'=============================
With UserForm17.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer
UserForm17.ListView1.ListItems.Clear
last1 = Sheets(13).Range("A" & Rows.count).End(xlUp).row

For frw1 = 9 To last1
 If Sheets(13).Cells(frw1, "bf").value Like "*" & UserForm17.TextBox2.Text & "*" Then
Set item1 = .ListItems.Add(, , Sheets(13).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(13).Cells(frw1, "b")
item1.SubItems(2) = Sheets(13).Cells(frw1, "c")
item1.SubItems(3) = Sheets(13).Cells(frw1, "d")
item1.SubItems(4) = Sheets(13).Cells(frw1, "e")
item1.SubItems(5) = Sheets(13).Cells(frw1, "f")
item1.SubItems(6) = Sheets(13).Cells(frw1, "g")
item1.SubItems(7) = Sheets(13).Cells(frw1, "h")
item1.SubItems(8) = Sheets(13).Cells(frw1, "i")
item1.SubItems(9) = Sheets(13).Cells(frw1, "j")
item1.SubItems(10) = Sheets(13).Cells(frw1, "k")
item1.SubItems(11) = Sheets(13).Cells(frw1, "l")
item1.SubItems(12) = Sheets(13).Cells(frw1, "m")
item1.SubItems(13) = Sheets(13).Cells(frw1, "n")
item1.SubItems(14) = Sheets(13).Cells(frw1, "o")
item1.SubItems(15) = Sheets(13).Cells(frw1, "p")
item1.SubItems(16) = Sheets(13).Cells(frw1, "q")
item1.SubItems(17) = Sheets(13).Cells(frw1, "r")
item1.SubItems(18) = Sheets(13).Cells(frw1, "s")
item1.SubItems(19) = Sheets(13).Cells(frw1, "t")
item1.SubItems(20) = Sheets(13).Cells(frw1, "u")
item1.SubItems(21) = Sheets(13).Cells(frw1, "v")
item1.SubItems(22) = Sheets(13).Cells(frw1, "w")
item1.SubItems(23) = Sheets(13).Cells(frw1, "x")
item1.SubItems(24) = Sheets(13).Cells(frw1, "y")
item1.SubItems(25) = Sheets(13).Cells(frw1, "z")
item1.SubItems(26) = Sheets(13).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(13).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(13).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(13).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(13).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(13).Cells(frw1, "af")
item1.SubItems(32) = Sheets(13).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(13).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(13).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(13).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(13).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(13).Cells(frw1, "al")
item1.SubItems(38) = Sheets(13).Cells(frw1, "am")
item1.SubItems(39) = Sheets(13).Cells(frw1, "an")
item1.SubItems(40) = Sheets(13).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(13).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(13).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(13).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(13).Cells(frw1, "as")
item1.SubItems(45) = Sheets(13).Cells(frw1, "at")
item1.SubItems(46) = Sheets(13).Cells(frw1, "au")
item1.SubItems(47) = Sheets(13).Cells(frw1, "av")
item1.SubItems(48) = Sheets(13).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(13).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(13).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(13).Cells(frw1, "az")
item1.SubItems(52) = Sheets(13).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(13).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(13).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(13).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(13).Cells(frw1, "be")
item1.SubItems(57) = Sheets(13).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(13).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(13).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(13).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(13).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(13).Cells(frw1, "bk")
item1.SubItems(63) = Sheets(13).Cells(frw1, "bl")
item1.SubItems(64) = Sheets(13).Cells(frw1, "bm")
item1.SubItems(65) = Sheets(13).Cells(frw1, "bn")
item1.SubItems(66) = Sheets(13).Cells(frw1, "bo")
item1.SubItems(67) = Sheets(13).Cells(frw1, "bp")
item1.SubItems(68) = Sheets(13).Cells(frw1, "bq")
item1.SubItems(69) = Sheets(13).Cells(frw1, "br")
item1.SubItems(70) = Sheets(13).Cells(frw1, "bs")
item1.SubItems(71) = Sheets(13).Cells(frw1, "bt")
item1.SubItems(72) = Sheets(13).Cells(frw1, "bu")
item1.SubItems(73) = Sheets(13).Cells(frw1, "bv")
End If
Next
End With
End Sub
