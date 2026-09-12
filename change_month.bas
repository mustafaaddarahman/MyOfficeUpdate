Attribute VB_Name = "change_month"
Sub change_salary()

Dim iRow As Long, i As Long
'MsgBox "Õ«Ê· «œŒ«· «·„Êﬁ› «·ÊŸÌ›Ì ·Â–« «·„‰ ”» ·Ì ”‰Ï ··‰Ÿ«„  —ÕÌ·Â ﬂ‹‹ „Êﬁ› ‰ﬁ· «Ê «Õ«·… «·Ï «· ﬁ«⁄œ «Ê «Ì „Êﬁ› «Œ— Ì ·«∆„ „⁄ Õ«·… «· —ÕÌ·", vbDefaultButton1, "—”«·Â  ‰»ÌÂ"

   Sheets(35).Activate

    iRow = Range("A" & Rows.count).End(xlUp).row
'=======================================================
Dim lasr, Y As Integer
lasr = Sheets(55).Range("A" & Rows.count).End(xlUp).row
 For Y = 9 To lasr
 If Sheets(55).Cells(Y, 2).value = UserForm1.TextBox2.Text And Sheets(55).Cells(Y, 5).value = UserForm1.TextBox5 Then
'==============================================
Sheets(35).Activate
 Range("A" & iRow + 1).Offset(0, 0).value = Sheets(55).Cells(Y, 1).value
    Range("A" & iRow + 1).Offset(0, 1).value = Sheets(55).Cells(Y, 2).value
    Range("A" & iRow + 1).Offset(0, 2).value = Sheets(55).Cells(Y, 3).value
    Range("A" & iRow + 1).Offset(0, 3).value = Sheets(55).Cells(Y, 4).value
    Range("A" & iRow + 1).Offset(0, 4).value = Sheets(55).Cells(Y, 5)
    Range("A" & iRow + 1).Offset(0, 5).value = Sheets(55).Cells(Y, 6)
    Range("A" & iRow + 1).Offset(0, 6).value = Sheets(55).Cells(Y, 7)
    Range("A" & iRow + 1).Offset(0, 7).value = Sheets(55).Cells(Y, 8)
    Range("A" & iRow + 1).Offset(0, 8).value = Sheets(55).Cells(Y, 9)
    Range("A" & iRow + 1).Offset(0, 9).value = Sheets(55).Cells(Y, 10)
    Range("A" & iRow + 1).Offset(0, 10).value = Sheets(55).Cells(Y, 11)
    Range("A" & iRow + 1).Offset(0, 11).value = Sheets(55).Cells(Y, 12)
    Range("A" & iRow + 1).Offset(0, 12).value = Sheets(55).Cells(Y, 13)
   Range("A" & iRow + 1).Offset(0, 13).value = Sheets(55).Cells(Y, 14)
    Range("A" & iRow + 1).Offset(0, 14).value = Sheets(55).Cells(Y, 15)
    Range("A" & iRow + 1).Offset(0, 15).value = Sheets(55).Cells(Y, 16)
    Range("A" & iRow + 1).Offset(0, 16).value = Sheets(55).Cells(Y, 17)
   Range("A" & iRow + 1).Offset(0, 17).value = Sheets(55).Cells(Y, 18)
    Range("A" & iRow + 1).Offset(0, 18).value = Sheets(55).Cells(Y, 19)
   Range("A" & iRow + 1).Offset(0, 19).value = Sheets(55).Cells(Y, 20)
    Range("A" & iRow + 1).Offset(0, 20).value = Sheets(55).Cells(Y, 21)
    Range("A" & iRow + 1).Offset(0, 21).value = Sheets(55).Cells(Y, 22)
    Range("A" & iRow + 1).Offset(0, 22).value = Sheets(55).Cells(Y, 23)
 Range("A" & iRow + 1).Offset(0, 23).value = Sheets(55).Cells(Y, 24)
    Range("A" & iRow + 1).Offset(0, 24).value = Sheets(55).Cells(Y, 25)
    Range("A" & iRow + 1).Offset(0, 25).value = Sheets(55).Cells(Y, 26)
    Range("A" & iRow + 1).Offset(0, 26).value = Sheets(55).Cells(Y, 27)
   Range("A" & iRow + 1).Offset(0, 27).value = Sheets(55).Cells(Y, 28)
  Range("A" & iRow + 1).Offset(0, 28).value = Sheets(55).Cells(Y, 29)
  Range("A" & iRow + 1).Offset(0, 29).value = Sheets(55).Cells(Y, 3)
   Range("A" & iRow + 1).Offset(0, 30).value = Sheets(55).Cells(Y, 31)
  Range("A" & iRow + 1).Offset(0, 31).value = Sheets(55).Cells(Y, 32)
   Range("A" & iRow + 1).Offset(0, 32).value = Sheets(55).Cells(Y, 33)
    Range("A" & iRow + 1).Offset(0, 33).value = Sheets(55).Cells(Y, 34)
    Range("A" & iRow + 1).Offset(0, 34).value = Sheets(55).Cells(Y, 35)
    Range("A" & iRow + 1).Offset(0, 35).value = Sheets(55).Cells(Y, 36)
    Range("A" & iRow + 1).Offset(0, 36).value = Sheets(55).Cells(Y, 37)
    Range("A" & iRow + 1).Offset(0, 37).value = Sheets(55).Cells(Y, 38)
    Range("A" & iRow + 1).Offset(0, 38).value = Sheets(55).Cells(Y, 39)
    Range("A" & iRow + 1).Offset(0, 39).value = Sheets(55).Cells(Y, 40)
    Range("A" & iRow + 1).Offset(0, 40).value = Sheets(55).Cells(Y, 41)
    Range("A" & iRow + 1).Offset(0, 41).value = Sheets(55).Cells(Y, 42)
    Range("A" & iRow + 1).Offset(0, 42).value = Sheets(55).Cells(Y, 43)
    Range("A" & iRow + 1).Offset(0, 43).value = Sheets(55).Cells(Y, 44)
    Range("A" & iRow + 1).Offset(0, 44).value = Sheets(55).Cells(Y, 45)
    Range("A" & iRow + 1).Offset(0, 45).value = Sheets(55).Cells(Y, 46)
   Range("A" & iRow + 1).Offset(0, 46).value = Sheets(55).Cells(Y, 47)
   Range("A" & iRow + 1).Offset(0, 47).value = Sheets(55).Cells(Y, 48)
   Range("A" & iRow + 1).Offset(0, 48).value = Sheets(55).Cells(Y, 49)
    Range("A" & iRow + 1).Offset(0, 49).value = Sheets(55).Cells(Y, 50)
    Range("A" & iRow + 1).Offset(0, 50).value = Sheets(55).Cells(Y, 51)
    Range("A" & iRow + 1).Offset(0, 51).value = Sheets(55).Cells(Y, 52)
   Range("A" & iRow + 1).Offset(0, 52).value = Sheets(55).Cells(Y, 53)
   Range("A" & iRow + 1).Offset(0, 53).value = Sheets(55).Cells(Y, 54)
    Range("A" & iRow + 1).Offset(0, 54).value = Sheets(55).Cells(Y, 55)
    Range("A" & iRow + 1).Offset(0, 55).value = Sheets(55).Cells(Y, 56)
    Range("A" & iRow + 1).Offset(0, 56).value = Sheets(55).Cells(Y, 57)
   Range("A" & iRow + 1).Offset(0, 57).value = Sheets(55).Cells(Y, 58)
   Range("A" & iRow + 1).Offset(0, 58).value = Sheets(55).Cells(Y, 59)
    Range("A" & iRow + 1).Offset(0, 59).value = Sheets(55).Cells(Y, 60)
   Range("A" & iRow + 1).Offset(0, 60).value = Sheets(55).Cells(Y, 61)
   Range("A" & iRow + 1).Offset(0, 61).value = Sheets(55).Cells(Y, 62)
  Range("A" & iRow + 1).Offset(0, 62).value = Sheets(55).Cells(Y, 63)
   Range("A" & iRow + 1).Offset(0, 63).value = Sheets(55).Cells(Y, 64)
    Range("A" & iRow + 1).Offset(0, 64).value = Sheets(55).Cells(Y, 65)
    Range("A" & iRow + 1).Offset(0, 65).value = Sheets(55).Cells(Y, 66)
    Range("A" & iRow + 1).Offset(0, 66).value = Sheets(55).Cells(Y, 67)
    Range("A" & iRow + 1).Offset(0, 67).value = Sheets(55).Cells(Y, 68)
    Range("A" & iRow + 1).Offset(0, 68).value = Sheets(55).Cells(Y, 69)
       Range("A" & iRow + 1).Offset(0, 69).value = Sheets(55).Cells(Y, 70)
    Range("A" & iRow + 1).Offset(0, 70).value = Sheets(55).Cells(Y, 71)
    Range("A" & iRow + 1).Offset(0, 71).value = Sheets(55).Cells(Y, 72)
   ' Range("A" & iRow + 1).Offset(0, 72).Value = .userform1.textbox85.Value
    Range("A" & iRow + 1).Offset(0, 73).value = Sheets(55).Cells(Y, 74)
   ' Range("A" & iRow + 1).Offset(0, 74).Value = UserForm1.TextBox8.Value
    '  Range("A" & iRow + 1).Offset(0, 75).Value = UserForm1.TextBox9.Value
   ' Range("A" & iRow + 1).Offset(0, 76).Value = UserForm1.TextBox10.Value
   ' Range("A" & iRow + 1).Offset(0, 77).Value = UserForm1.TextBox31.Text
   ' Range("A" & iRow + 1).Offset(0, 78).Value = UserForm1.TextBox48.Text
   'Range("A" & iRow + 1).Offset(0, 79).Value = UserForm1.TextBox49.Text
'Sheets(35).Cells(y, 82) = "«Œ—  ÕœÌÀ"
'Range("A" & iRow + 1).Offset(0, 81).Value = UserForm27.TextBox1.Text
 ' Range("A" & iRow + 1).Offset(0, 82).Value = UserForm27.TextBox7.Text
End If
Next
'Range("A" & iRow + 1).Offset(0, 81).Value = UserForm27.TextBox1.Text

'=========================================================
    
   Range("cc9").Formula2 = "=BZ9-AE9"
         Range("cc9:cc1000").FillDown
      
      
       
'============================================================================








End Sub
