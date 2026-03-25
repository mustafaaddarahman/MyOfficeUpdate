Attribute VB_Name = "case_update_chang_salary001"
Sub case_update_chang_salary()
On Error Resume Next
Sheets(36).Range("a9:cd8000").ClearContents
ActiveSheet.Range("a9: cd6000").UnMerge
Dim arr001 As Variant, i As Long, j As Long, k As Long
Dim arr002 As Variant
Dim lastR As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
arr001 = Sheets(35).Range("a9:cd" & lastR)
j = 0
ReDim arr002(1 To UBound(arr001, 1), 1 To UBound(arr001, 2))

For i = 1 To UBound(arr001, 1)

 If arr001(i, 66) >= CDate(UserForm28.TextBox2.Text) Then

 j = j + 1
 For k = 1 To 82
 arr002(j, k) = arr001(i, k)
  
  Next k
  
End If

Next


Sheets(36).Range("a9").Resize(j, 82).value = arr002
Sheets(36).Activate
'=============================
With UserForm28.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer
UserForm28.ListView1.ListItems.Clear
last1 = Sheets(35).Range("A" & Rows.count).End(xlUp).row

For frw1 = 9 To last1
 If Sheets(35).Cells(frw1, "bn") >= CDate(UserForm28.TextBox2.Text) Then
Set item1 = .ListItems.Add(, , Sheets(35).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(35).Cells(frw1, "b")
item1.SubItems(2) = Sheets(35).Cells(frw1, "c")
item1.SubItems(3) = Sheets(35).Cells(frw1, "d")
item1.SubItems(4) = Sheets(35).Cells(frw1, "e")
item1.SubItems(5) = Sheets(35).Cells(frw1, "f")
item1.SubItems(6) = Sheets(35).Cells(frw1, "g")
item1.SubItems(7) = Sheets(35).Cells(frw1, "h")
item1.SubItems(8) = Sheets(35).Cells(frw1, "i")
item1.SubItems(9) = Sheets(35).Cells(frw1, "j")
item1.SubItems(10) = Sheets(35).Cells(frw1, "k")
item1.SubItems(11) = Sheets(35).Cells(frw1, "l")
item1.SubItems(12) = Sheets(35).Cells(frw1, "m")
item1.SubItems(13) = Sheets(35).Cells(frw1, "n")
item1.SubItems(14) = Sheets(35).Cells(frw1, "o")
item1.SubItems(15) = Sheets(35).Cells(frw1, "p")
item1.SubItems(16) = Sheets(35).Cells(frw1, "q")
item1.SubItems(17) = Sheets(35).Cells(frw1, "r")
item1.SubItems(18) = Sheets(35).Cells(frw1, "s")
item1.SubItems(19) = Sheets(35).Cells(frw1, "t")
item1.SubItems(20) = Sheets(35).Cells(frw1, "u")
item1.SubItems(21) = Sheets(35).Cells(frw1, "v")
item1.SubItems(22) = Sheets(35).Cells(frw1, "w")
item1.SubItems(23) = Sheets(35).Cells(frw1, "x")
item1.SubItems(24) = Sheets(35).Cells(frw1, "y")
item1.SubItems(25) = Sheets(35).Cells(frw1, "z")
item1.SubItems(26) = Sheets(35).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(35).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(35).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(35).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(35).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(35).Cells(frw1, "af")
item1.SubItems(32) = Sheets(35).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(35).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(35).Cells(frw1, "ai")










End If
Next
End With
End Sub
