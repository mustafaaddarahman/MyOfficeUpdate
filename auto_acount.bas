Attribute VB_Name = "auto_acount"
Sub auto_acount_1()

Dim lr, nrow As Integer

lr = Sheets(1).Cells(Rows.count, 2).End(xlUp).row
For nrow = 9 To lr
If Sheets(1).Cells(nrow, "b").value = UserForm48.TextBox2.Text And Sheets(1).Cells(nrow, "e").value = UserForm48.TextBox5.Text Then

Sheets(1).Cells(nrow, "by") = Sheets(1).Cells(nrow, "aw")
Sheets(1).Cells(nrow, "bi") = UserForm48.TextBox61.Text


'================================
Select Case UserForm48.TextBox100.value
'==================================================„” „— »«·Œœ„…
Case 1
Dim lr1, nrow1 As Integer
lr1 = Sheets(1).Cells(Rows.count, 2).End(xlUp).row
For nrow1 = 9 To lr1
If UserForm48.TextBox8.value >= 8 And UserForm48.CheckBox4.value = True _
And Sheets(1).Cells(nrow1, "b").value = UserForm48.TextBox2.Text And Sheets(1).Cells(nrow1, "e").value = UserForm48.TextBox5.Text Then
Sheets(1).Cells(nrow1, "cd") = "true"
Sheets(9).Cells(18, "g") = UserForm48.TextBox8.Text
Sheets(9).Cells(18, "f") = UserForm48.TextBox9.Text
UserForm48.TextBox10.Text = Sheets(9).Cells(18, "h")
UserForm48.TextBox105.value = UserForm48.TextBox10.value + UserForm48.TextBox10.value * 0.5 - UserForm48.TextBox10.value



UserForm48.ComboBox10.value = 0
UserForm48.ComboBox11.value = 0
 UserForm48.TextBox27.value = 0
 UserForm48.TextBox29.value = 0
  Call mastmer_5dma_1
ElseIf UserForm48.CheckBox4.value = False _
And Sheets(1).Cells(nrow1, "b").value = UserForm48.TextBox2.Text And Sheets(1).Cells(nrow1, "e").value = UserForm48.TextBox5.Text Then
Sheets(1).Cells(nrow1, "cd") = "false"
Sheets(9).Cells(18, "g") = UserForm48.TextBox8.Text
Sheets(9).Cells(18, "f") = UserForm48.TextBox9.Text
UserForm48.TextBox10.Text = Sheets(9).Cells(18, "h")
Call mastmer_5dma_1
ElseIf UserForm48.CheckBox4.value = True And UserForm48.TextBox8.value < 8 _
And Sheets(1).Cells(nrow1, "b").value = UserForm48.TextBox2.Text And Sheets(1).Cells(nrow1, "e").value = UserForm48.TextBox5.Text Then
Sheets(1).Cells(nrow1, "cd") = "false"
Sheets(9).Cells(18, "g") = UserForm48.TextBox8.Text
Sheets(9).Cells(18, "f") = UserForm48.TextBox9.Text
UserForm48.TextBox10.Text = Sheets(9).Cells(18, "h")
Call mastmer_5dma_1
End If
Next nrow1
'======================================== «Ã«“… Œ„” ”‰Ê« 
Case 2
Call five_snwat_1

'==================================================«Ã«“… »œÊ‰ —« »
Case 3
Call bdoon_ratb1_1
'===========================================«Ã«“… „⁄Ì·
Case 4
Call mo3el_1_1
'=================================—›⁄ Ìœ «Ê «Ìﬁ«› ’—›
Case 5
Call raf3_yead_1_1
'=================================«„Ê„… «Ê· ”  «‘Â—
Case 6
Call mother_hood_1_6_1_1
'=================================«„Ê„… À«‰Ì ”  «‘Â—
Case 7
Call mother_hood_2_6_1_1
'================================= ‰”Ì»
Case 8
Call tanseeb_1_1
'=================================⁄ﬁœ
Case 9
Call contract_1_1
'=================================«Ã«“… ·Ã«‰ ÿ»Ì…

Case 10
Call lejan_tabea_1_1

'=================================„” „— »«·Œœ„Â „⁄  ÷„Ì‰…

Case 12
Call mastmer_5dma_1




'=================================„” „— »«·Œœ„Â „⁄ ”·›…

Case 13
Call mastmer_5dma_1


End Select

'==================================================«·’«›Ì





   '=========================================
Call update_data_acount_1


End If
Next
UserForm48.TextBox102.value = Format(Sheets(26).Cells(3, "s"), "00.0")
If UserForm48.TextBox102.value = 0 Then
 Label122.Caption = "«·ﬁÌ„ „ ÿ«»ﬁ…"
 Else
 UserForm48.Label122.Caption = "«·ﬁÌ„ €Ì— „ ÿ«»ﬁ…"
End If
MsgBox " „ «·«Õ ”«» »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"


UserForm48.TextBox76.Text = Sheets(1).Cells(2, "cc")
UserForm48.TextBox73.Text = Sheets(1).Cells(2, "cd")
UserForm48.TextBox74.Text = Sheets(1).Cells(2, "ce")
UserForm48.TextBox79.Text = Sheets(1).Cells(2, "cf")
UserForm48.TextBox80.Text = Sheets(1).Cells(2, "cg")
UserForm48.TextBox82.Text = Sheets(1).Cells(1, "bn")
UserForm48.TextBox66.Text = Sheets(1).Cells(2, "ch")


'UserForm48.TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'UserForm48.TextBox75.Text = ListView1.ListItems.Count
UserForm48.TextBox73.value = Format(UserForm48.TextBox73.value, "#,## IQD")
UserForm48.TextBox74.value = Format(UserForm48.TextBox74.value, "#,## IQD")
UserForm48.TextBox79.value = Format(UserForm48.TextBox79.value, "#,## IQD")
UserForm48.TextBox80.value = Format(UserForm48.TextBox80.value, "#,## IQD")
UserForm48.TextBox66.value = Format(UserForm48.TextBox66.value, "#,## IQD")
UserForm48.TextBox76.value = Format(UserForm48.TextBox76.value, "#,## IQD")
'=====================================‰ﬁ· «·»Ì«‰«  «·Ï ‘Ì  46
Sheets(46).Activate
Dim iRow3 As Long
Dim i3 As Long
Dim lastr3 As Long
Dim y3 As Long
If UserForm48.TextBox11.value > 0 And UserForm48.TextBox12 < 30 Then
       

 lastr3 = Sheets(46).Cells(Rows.count, 1).End(xlUp).row
    For y3 = 2 To lastr3
    
    If Sheets(46).Cells(y3, 1) = UserForm48.TextBox2.Text And Sheets(46).Cells(y3, 2) = UserForm48.TextBox5.Text Then
   Cells(y3, 1).Select
ActiveCell.EntireRow.Delete
    Exit For
    End If
    Next y3
     
    lastr3 = Sheets(46).Cells(Rows.count, 1).End(xlUp).row
    For y3 = 2 To lastr3
      If Sheets(46).Cells(y3, 1) <> UserForm48.TextBox2.Text And Sheets(46).Cells(y3, 2) <> UserForm48.TextBox5.Text Then
    
    Exit For
    End If
     Next
    iRow = Range("A" & Rows.count).End(xlUp).row
    Range("A" & iRow + 1).Offset(0, 0).value = UserForm48.TextBox2.value
    Range("A" & iRow + 1).Offset(0, 1).value = UserForm48.TextBox5.value
    Range("A" & iRow + 1).Offset(0, 2).value = UserForm48.TextBox10.value
    Range("A" & iRow + 1).Offset(0, 3).value = UserForm48.TextBox49.value
    Range("A" & iRow + 1).Offset(0, 4).value = UserForm48.TextBox12.value & "·œÌÂ €Ì«» Êﬁœ—Â " & UserForm48.TextBox11.value
    Range("A" & iRow + 1).Offset(0, 5).value = UserForm48.TextBox31.value
    Range("A" & iRow + 1).Offset(0, 6).value = UserForm48.TextBox58.value
    
Dim month_date As Date
Dim name_month As Variant
month_date = Month(Sheets(46).Cells(1, "m"))
name_month = monthName(month_date)
 Range("A" & iRow + 1).Offset(0, 7).value = name_month
   
End If




End Sub


