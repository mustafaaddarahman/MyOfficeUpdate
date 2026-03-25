VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm21 
   Caption         =   "UserForm21"
   ClientHeight    =   6690
   ClientLeft      =   6912
   ClientTop       =   468
   ClientWidth     =   8928.001
   OleObjectBlob   =   "UserForm21.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CheckBox10_Click()

End Sub

Private Sub CheckBox4_Click()

End Sub

Private Sub CheckBox5_Click()

End Sub

Private Sub CheckBox6_Click()

End Sub

Private Sub CommandButton1_Click()
'Sheets(1).Cells(9, "bo").Value = TextBox1.Text
'Ê—ﬁ…1.Range("bo9").Formula = Sheets(1).Cells(9, "bo").Value
'Ê—ﬁ…1.Range("bo9:bo3000").FillDown
'Sheets(1).Cells(9, "bs").Value = TextBox2.Text
'Ê—ﬁ…1.Range("bs9:bs3000").FillDown
'=============================«·«Õ ”«» «·„«·Ì »⁄œ «·«” ﬁÿ«⁄

'Ê—ﬁ…1.Range("aw9").Formula = "=IF(OR(BF9=""„” „— »«·Œœ„…"",BF9=""«„Ê„… «Ê· 6 «‘Â—"",BF9=""«„Ê„… À«‰Ì 6 «‘Â—""),(AE9+bp9+br9)-(AV9+BO9+bq9),AE9-AV9)"
'Ê—ﬁ…1.Range("aw9:aw3000").FillDown


'MsgBox "·ﬁœ  „  «” ﬁÿ«⁄ «·„»«·€ „‰ ﬂ«›… «·„ÊŸ›Ì‰", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
Dim lastR As Integer
Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11 As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row




For n1 = 9 To lastR
If CheckBox1.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" Then
Sheets(1).Cells(n1, "bo") = TextBox1.value
Sheets(1).Cells(n1, "bs") = TextBox2.value

End If
Next


For n2 = 9 To lastR
If CheckBox2.value = True And Sheets(1).Cells(n2, "bf") = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Then
Sheets(1).Cells(n2, "bo") = TextBox1.value
Sheets(1).Cells(n2, "bs") = TextBox2.value
End If
Next


For n3 = 9 To lastR
If CheckBox3.value = True And Sheets(1).Cells(n3, "bf") = " «Ã«“… »œÊ‰ —« »" Then
Sheets(1).Cells(n3, "bo") = TextBox1.value
Sheets(1).Cells(n3, "bs") = TextBox2.value
End If
Next



For n4 = 9 To lastR
If CheckBox4.value = True And Sheets(1).Cells(n4, "bf") = "≈Ã«“… „⁄Ì· «·„ ›—€" Then
Sheets(1).Cells(n4, "bo") = TextBox1.value
Sheets(1).Cells(n4, "bs") = TextBox2.value
End If
Next





For n5 = 9 To lastR
If CheckBox5.value = True And Sheets(1).Cells(n5, "bf") = "—›⁄ Ìœ" Then
Sheets(1).Cells(n5, "bo") = TextBox1.value
Sheets(1).Cells(n5, "bs") = TextBox2.value
End If
Next





For n6 = 9 To lastR
If CheckBox6.value = True And Sheets(1).Cells(n6, "bf") = "«„Ê„Â «Ê· 6 «‘Â—" Then
Sheets(1).Cells(n6, "bo") = TextBox1.value
Sheets(1).Cells(n6, "bs") = TextBox2.value
End If
Next





For n7 = 9 To lastR
If CheckBox7.value = True And Sheets(1).Cells(n7, "bf") = "«„Ê„Â À«‰Ì 6 «‘Â—" Then
Sheets(1).Cells(n7, "bo") = TextBox1.value
Sheets(1).Cells(n7, "bs") = TextBox2.value
End If
Next






For n8 = 9 To lastR
If CheckBox8.value = True And Sheets(1).Cells(n8, "bf") = " ‰”Ì»" Then
Sheets(1).Cells(n8, "bo") = TextBox1.value
Sheets(1).Cells(n8, "bs") = TextBox2.value
End If
Next





For n9 = 9 To lastR
If CheckBox9.value = True And Sheets(1).Cells(n9, "bf") = "⁄ﬁœ" Then
Sheets(1).Cells(n9, "bo") = TextBox1.value
Sheets(1).Cells(n9, "bs") = TextBox2.value
End If
Next








For n10 = 9 To lastR
If CheckBox10.value = True And Sheets(1).Cells(n10, "bf") = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" Then
Sheets(1).Cells(n10, "bo") = TextBox1.value
Sheets(1).Cells(n10, "bs") = TextBox2.value
End If
Next






For n11 = 9 To lastR
If CheckBox3.value = True And Sheets(1).Cells(n11, "bf") = "—« » Ã“∆Ì" Then
Sheets(1).Cells(n11, "bo") = TextBox1.value
Sheets(1).Cells(n11, "bs") = TextBox2.value
End If
Next

Call count_all

MsgBox "·ﬁœ  „  «” ﬁÿ«⁄ «·„»«·€ „‰ ﬂ«›… «·„ÊŸ›Ì‰", vbDefaultButton1, "—”«·…  ÊÃÌÂ"








End Sub


Private Sub CommandButton2_Click()
TextBox1.value = 0
TextBox2.value = ""

Sheets(1).Cells(9, "bo").value = TextBox1.Text
Ê—ﬁ…1.Range("bo9").Formula = Sheets(1).Cells(9, "bo").value
Ê—ﬁ…1.Range("bo9:bo3000").FillDown
Sheets(1).Cells(9, "bs").value = TextBox2.Text
Ê—ﬁ…1.Range("bs9:bs3000").FillDown
Ê—ﬁ…1.Range("aw9").Formula = "=IF(OR(BF9=""„” „— »«·Œœ„…"",BF9=""Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« "",BF9="" «Ã«“… »œÊ‰ —« »"",BF9=""≈Ã«“… „⁄Ì· «·„ ›—€"",BF9=""—›⁄ Ìœ"",BF9="" ‰”Ì»"",BF9=""⁄ﬁœ"",BF9=""≈Ã«“… ·Ã«‰ ÿ»ÌÂ"",BF9=""—« » Ã“∆Ì"",BF9=""·œÌÂ  ÷„Ì‰"",BF9=""·œÌÂ ”·›"",BF9=""«„Ê„Â «Ê· 6 «‘Â—"",BF9=""«„Ê„… À«‰Ì 6 «‘Â—""),(AE9+bp9+br9)-(AV9+BO9+bq9),AE9-AV9)"
Ê—ﬁ…1.Range("aw9:aw3000").FillDown
Call count_all
MsgBox "·ﬁœ  „   ’›Ì— Ã„Ì⁄ «·„»«·€ «·„” ﬁÿ⁄Â „‰ ﬂ«›… «·„ÊŸ›Ì‰", vbDefaultButton1, "—”«·…  ÊÃÌÂ"

End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()
TextBox3.value = Sheets(1).Cells(9, "bo").value
End Sub


