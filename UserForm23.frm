VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm23 
   Caption         =   "UserForm23"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9900.001
   OleObjectBlob   =   "UserForm23.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If CheckBox1.value = True Then

End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.value = True Then
CheckBox1.value = False
End If
End Sub

Private Sub CommandButton1_Click()


Dim lastR As Integer
Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12 As Integer
lastR = Sheets(1).Cells(Rows.count, 2).End(xlUp).row
'=============================================================

For n1 = 9 To lastR
If CheckBox3.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" And CheckBox1.value = True Then

Sheets(1).Cells(n1, "BQ").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n1, "j")
Sheets(1).Cells(n1, "BS").value = TextBox2.Text
End If
Next

'===========================================================
For n2 = 9 To lastR
If CheckBox4.value = True And Sheets(1).Cells(n2, "bf") = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " And CheckBox1.value = True Then
Sheets(1).Cells(n2, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n2, "j")
Sheets(1).Cells(n2, "bs").value = TextBox2.Text
End If
Next








'=================================================

For n3 = 9 To lastR
If CheckBox5.value = True And Sheets(1).Cells(n3, "bf") = "«Ã«“… »œÊ‰ —« »" And CheckBox1.value = True Then
Sheets(1).Cells(n3, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n3, "j")
Sheets(1).Cells(n3, "bs").value = TextBox2.Text
End If
Next








'========================================================


For n4 = 9 To lastR
If CheckBox6.value = True And Sheets(1).Cells(n4, "bf") = "≈Ã«“… „⁄Ì· «·„ ›—€" And CheckBox1.value = True Then
Sheets(1).Cells(n4, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n4, "j")
Sheets(1).Cells(n4, "bs").value = TextBox2.Text
End If
Next












'===================================================




For n5 = 9 To lastR
If CheckBox7.value = True And Sheets(1).Cells(n5, "bf") = "—›⁄ Ìœ" And CheckBox1.value = True Then
Sheets(1).Cells(n5, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n5, "j")
Sheets(1).Cells(n5, "bs").value = TextBox2.Text
End If
Next








'=======================================================




For n6 = 9 To lastR

If CheckBox8.value = True And Sheets(1).Cells(n6, "bf") = "«„Ê„Â «Ê· 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n6, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n6, "j")
Sheets(1).Cells(n6, "bs").value = TextBox2.Text
End If
Next








'=====================================================




For n7 = 9 To lastR
If CheckBox9.value = True And Sheets(1).Cells(n7, "bf") = "«„Ê„Â À«‰Ì 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n7, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n7, "j")
Sheets(1).Cells(n7, "bs").value = TextBox2.Text
End If
Next






'=====================================================





For n8 = 9 To lastR
If CheckBox10.value = True And Sheets(1).Cells(n8, "bf") = " ‰”Ì»" And CheckBox1.value = True Then
Sheets(1).Cells(n8, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n8, "j")
Sheets(1).Cells(n8, "bs").value = TextBox2.Text
End If
Next





'==================================================




For n9 = 9 To lastR
If CheckBox11.value = True And Sheets(1).Cells(n9, "bf") = "⁄ﬁœ" And CheckBox1.value = True Then
Sheets(1).Cells(n9, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n9, "j")
Sheets(1).Cells(n9, "bs").value = TextBox2.Text
End If
Next











'=======================================================







For n10 = 9 To lastR

If CheckBox12.value = True And Sheets(1).Cells(n10, "bf") = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" And CheckBox1.value = True Then
Sheets(1).Cells(n10, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n10, "j")
Sheets(1).Cells(n10, "bs").value = TextBox2.Text
End If
Next

















'===========================================





For n11 = 9 To lastR
If CheckBox13.value = True And Sheets(1).Cells(n11, "bf") = "—« » Ã“∆Ì" And CheckBox1.value = True Then
Sheets(1).Cells(n11, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n11, "j")
Sheets(1).Cells(n11, "bs").value = TextBox2.Text
End If
Next














'=======================================================
For n12 = 9 To lastR
If Sheets(1).Cells(n12, "h").value = 8 Or Sheets(1).Cells(n12, "h").value = 9 Or Sheets(1).Cells(n12, "h").value = 10 _
And CheckBox14.value = True And CheckBox1.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" Then

Sheets(1).Cells(n12, "bq").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n12, "j")
 Sheets(1).Cells(n12, "bs").value = TextBox2.Text

End If
Next
Call count_all
MsgBox " „  «÷«›… «·‰”»Â"














End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
'==============================
If TextBox1.value = "" Then
TextBox1.value = 0
TextBox3.value = 0
End If
Dim lastR As Integer
Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12 As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
'=============================================================

For n1 = 9 To lastR
If CheckBox3.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" And CheckBox1.value = True Then
Sheets(1).Cells(n1, "bq").value = 0
Sheets(1).Cells(n1, "bs").value = TextBox2.Text
End If

Next
TextBox3.value = 0

'===========================================================
For n2 = 9 To lastR
If CheckBox4.value = True And Sheets(1).Cells(n2, "bf") = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " And CheckBox1.value = True Then
Sheets(1).Cells(n2, "bq").value = 0
Sheets(1).Cells(n2, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0

'=================================================
For n3 = 9 To lastR
If CheckBox5.value = True And Sheets(1).Cells(n3, "bf") = "«Ã«“… »œÊ‰ —« »" And CheckBox1.value = True Then
Sheets(1).Cells(n3, "bq").value = 0
Sheets(1).Cells(n3, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0

'========================================================
For n4 = 9 To lastR
If CheckBox6.value = True And Sheets(1).Cells(n4, "bf") = "≈Ã«“… „⁄Ì· «·„ ›—€" And CheckBox1.value = True Then
Sheets(1).Cells(n4, "bq").value = 0
Sheets(1).Cells(n4, "bs").value = TextBox2.Text
TextBox3.value = 0
End If
Next
TextBox3.value = 0

'===================================================
For n5 = 9 To lastR
If CheckBox7.value = True And Sheets(1).Cells(n5, "bf") = "—›⁄ Ìœ" And CheckBox1.value = True Then
Sheets(1).Cells(n5, "bq").value = 0
Sheets(1).Cells(n5, "bs").value = TextBox2.Text
TextBox3.value = 0
End If
Next
TextBox3.value = 0

'=======================================================
For n6 = 9 To lastR
If CheckBox8.value = True And Sheets(1).Cells(n6, "bf") = "«„Ê„Â «Ê· 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n6, "bq").value = 0
Sheets(1).Cells(n6, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0


'=====================================================
For n7 = 9 To lastR
If CheckBox9.value = True And Sheets(1).Cells(n7, "bf") = "«„Ê„Â À«‰Ì 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n7, "bq").value = 0
Sheets(1).Cells(n7, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0

'=====================================================
For n8 = 9 To lastR
If CheckBox10.value = True And Sheets(1).Cells(n8, "bf") = " ‰”Ì»" And CheckBox1.value = True Then
Sheets(1).Cells(n8, "bq").value = 0
Sheets(1).Cells(n8, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0

'==================================================
For n9 = 9 To lastR
If CheckBox11.value = True And Sheets(1).Cells(n9, "bf") = "⁄ﬁœ" And CheckBox1.value = True Then
Sheets(1).Cells(n9, "bq").value = 0
Sheets(1).Cells(n9, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0


'=======================================================
For n10 = 9 To lastR
If CheckBox12.value = True And Sheets(1).Cells(n10, "bf") = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" And CheckBox1.value = True Then
Sheets(1).Cells(n10, "bq").value = 0
Sheets(1).Cells(n10, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0
 
'===========================================
For n11 = 9 To lastR
If CheckBox13.value = True And Sheets(1).Cells(n11, "bf") = "—« » Ã“∆Ì" And CheckBox1.value = True Then
Sheets(1).Cells(n11, "bq").value = 0
Sheets(1).Cells(n12, "bs").value = TextBox2.Text
End If
Next
TextBox3.value = 0

'=======================================================
For n12 = 9 To lastR
If Sheets(1).Cells(n12, "h").value = 8 Or Sheets(1).Cells(n12, "h").value = 9 Or Sheets(1).Cells(n12, "h").value = 10 _
And CheckBox14.value = True And CheckBox1.value = True And Sheets(1).Cells(n12, "bf") = "„” „— »«·Œœ„…" And TextBox1.value = "" Then
TextBox1.value = 0
Sheets(1).Cells(n12, "bq").value = 0
Sheets(1).Cells(n12, "bs").value = TextBox2.Text

End If
Next
 TextBox3.value = 0

Call count_all
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub TextBox1_Change()
Label4.Caption = "Â–Â «·ﬁÌ„Â  „À· „ﬁœ«— ‰”»… «·„∆ÊÌÂ „‰ „Œ’’«  «·—« » " & " " & Sheets(1).Cells(9, "e") & " " & _
"«·«”„ «·«Ê· „‰ »Ì«‰«  «·‰Ÿ«„"

End Sub

Private Sub TextBox3_AfterUpdate()
End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub
