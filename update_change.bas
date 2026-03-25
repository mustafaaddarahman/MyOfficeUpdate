Attribute VB_Name = "update_change"
Sub update_change_salary()
On Error Resume Next
Dim lastR As Integer
lastR = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(35).Cells(Y, 2) = UserForm1.TextBox2.Text And Sheets(35).Cells(Y, 5) = UserForm1.TextBox5.Text And Sheets(35).Cells(Y, 84) = UserForm27.TextBox9.Text Then
   
   
    Exit For
    End If
    Next Y
  ' Sheets(35).Cells(y, 1) = userform1.TextBox1.Text
   ' Sheets(35).Cells(y, 2) = userform1.TextBox2.Text
   ' Sheets(35).Cells(y, 3) = userform1.TextBox3.Text
  '  Sheets(35).Cells(y, 4) = userform1.ComboBox4.Value
   ' Sheets(35).Cells(y, 5) = userform1.TextBox5.Text
    Sheets(35).Cells(Y, 6) = UserForm1.ComboBox5.value
    Sheets(35).Cells(Y, 7) = UserForm1.TextBox7.Text
   ' Sheets(35).Cells(y, 8) = userform1.TextBox8.Text
   ' Sheets(35).Cells(y, 9) = userform1.TextBox9.Text
   ' Sheets(35).Cells(y, 10) = userform1.TextBox10.Text
    Sheets(35).Cells(Y, 11) = UserForm1.TextBox11.Text
    Sheets(35).Cells(Y, 12) = UserForm1.TextBox12.Text
     Sheets(35).Cells(Y, 13) = UserForm1.ComboBox6.value
    Sheets(35).Cells(Y, 14) = UserForm1.TextBox14.Text
    Sheets(35).Cells(Y, 15) = UserForm1.ComboBox7.value
    Sheets(35).Cells(Y, 16) = UserForm1.TextBox16.Text
    Sheets(35).Cells(Y, 17) = UserForm1.ComboBox8.value
    Sheets(35).Cells(Y, 18) = UserForm1.TextBox18.Text
    Sheets(35).Cells(Y, 19) = UserForm1.TextBox19.Text
    Sheets(35).Cells(Y, 20) = UserForm1.TextBox20.Text
    Sheets(35).Cells(Y, 21) = UserForm1.TextBox21.Text
    Sheets(35).Cells(Y, 22) = UserForm1.TextBox22.Text
    Sheets(35).Cells(Y, 23) = UserForm1.TextBox23.Text
    Sheets(35).Cells(Y, 24) = UserForm1.TextBox24.Text
    Sheets(35).Cells(Y, 25) = UserForm1.TextBox25.Text
    Sheets(35).Cells(Y, 26) = UserForm1.ComboBox10.value
    Sheets(35).Cells(Y, 27) = UserForm1.TextBox27.Text
    Sheets(35).Cells(Y, 28) = UserForm1.ComboBox11.value
  Sheets(35).Cells(Y, 29) = UserForm1.TextBox29.Text
    Sheets(35).Cells(Y, 30) = UserForm1.TextBox30.Text
   ' Sheets(35).Cells(y, 31) = userform1.TextBox31.Text
  Sheets(35).Cells(Y, 32) = UserForm1.TextBox32.Text
   Sheets(35).Cells(Y, 33) = UserForm1.TextBox33.Text
    Sheets(35).Cells(Y, 34) = UserForm1.TextBox34.Text
    Sheets(35).Cells(Y, 35) = UserForm1.TextBox35.Text
    Sheets(35).Cells(Y, 36) = UserForm1.TextBox36.Text
    Sheets(35).Cells(Y, 37) = UserForm1.TextBox37.Text
    Sheets(35).Cells(Y, 38) = UserForm1.TextBox38.Text
    Sheets(35).Cells(Y, 39) = UserForm1.TextBox39.Text
    Sheets(35).Cells(Y, 40) = UserForm1.TextBox40.Text
    Sheets(35).Cells(Y, 41) = UserForm1.TextBox41.Text
    Sheets(35).Cells(Y, 42) = UserForm1.TextBox42.Text
    Sheets(35).Cells(Y, 43) = UserForm1.TextBox43.Text
    Sheets(35).Cells(Y, 44) = UserForm1.TextBox44.Text
    Sheets(35).Cells(Y, 45) = UserForm1.TextBox45.Text
    Sheets(35).Cells(Y, 46) = UserForm1.TextBox46.Text
    Sheets(35).Cells(Y, 47) = UserForm1.TextBox47.Text
 ' Sheets(35).Cells(y, 48) = userform1.TextBox48.Text
 '  Sheets(35).Cells(y, 49) = userform1.TextBox49.Text
   Sheets(35).Cells(Y, 50) = UserForm1.TextBox50.Text
    Sheets(35).Cells(Y, 51) = UserForm1.ComboBox9.value
 '   Sheets(35).Cells(y, 52) = UserForm1.TextBox52.Text
'  Sheets(35).Cells(y, 53) = UserForm1.TextBox53.Text
 '  Sheets(35).Cells(y, 54) = UserForm1.TextBox54.Text
  '  Sheets(35).Cells(y, 55) = UserForm1.TextBox55.Text
 '   Sheets(35).Cells(y, 56) = UserForm1.TextBox56.Text
  '  Sheets(35).Cells(y, 57) = UserForm1.TextBox57.Text
 ' Sheets(35).Cells(y, 58) = UserForm1.TextBox58.Text
  ' Sheets(35).Cells(y, 59) = UserForm1.TextBox59.Text
    Sheets(35).Cells(Y, 60) = UserForm1.TextBox60.Text
 '  Sheets(35).Cells(y, 61) = UserForm1.TextBox61.Text
  Sheets(35).Cells(Y, 62) = UserForm1.TextBox62.Text
 '  Sheets(35).Cells(y, 63) = UserForm1.TextBox63.Text
 ' Sheets(35).Cells(y, 64) = UserForm1.TextBox64.Text
    
    
    Sheets(35).Cells(Y, 66) = Format(UserForm1.TextBox82.Text, "yyyy/mm/dd")
  Sheets(35).Cells(Y, 67) = UserForm1.TextBox83.Text
   
   Sheets(35).Cells(Y, 68) = UserForm1.TextBox84.Text
 Sheets(35).Cells(Y, 69) = UserForm1.TextBox85.Text
  Sheets(35).Cells(Y, 70) = UserForm1.TextBox86.Text
   Sheets(35).Cells(Y, 71) = UserForm1.TextBox87.Text
   Sheets(35).Cells(Y, 72) = UserForm1.TextBox88.Text
   Sheets(35).Cells(Y, 73) = UserForm1.TextBox64.Text
   Sheets(35).Cells(Y, 74) = UserForm1.TextBox91.Text
    
     Sheets(35).Cells(Y, 75) = UserForm1.TextBox8.Text
 Sheets(35).Cells(Y, 76) = UserForm1.TextBox9.Text
  Sheets(35).Cells(Y, 77) = UserForm1.TextBox10.Text
   Sheets(35).Cells(Y, 78) = UserForm1.TextBox31.Text
   Sheets(35).Cells(Y, 79) = UserForm1.TextBox48.Text
   Sheets(35).Cells(Y, 80) = UserForm1.TextBox49.Text
   
   Sheets(35).Cells(Y, 82) = UserForm27.TextBox1.Text
   Sheets(35).Cells(Y, 83) = UserForm27.TextBox7.Text
   
   
      
     MsgBox " „  ⁄„·Ì…  ÕœÌÀ «·»Ì«‰«  »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"



emad:







End Sub
