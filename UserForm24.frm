VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm24 
   Caption         =   "UserForm24"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9672.001
   OleObjectBlob   =   "UserForm24.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If CheckBox1.value = True Then
CheckBox2.value = False
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.value = True Then
CheckBox1.value = False
End If
End Sub

Private Sub CommandButton1_Click()


'==============================
Dim lastR As Integer
Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12 As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
'=============================================================

For n1 = 9 To lastR
If CheckBox3.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" And CheckBox1.value = True Then
Sheets(1).Cells(n1, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n1, "j")
Sheets(1).Cells(n1, "bt").value = TextBox2.Text
End If
Next

'===========================================================
For n2 = 9 To lastR
If CheckBox4.value = True And Sheets(1).Cells(n2, "bf") = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " And CheckBox1.value = True Then
Sheets(1).Cells(n2, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n2, "j")
Sheets(1).Cells(n2, "bt").value = TextBox2.Text
End If
Next








'=================================================

For n3 = 9 To lastR
If CheckBox5.value = True And Sheets(1).Cells(n3, "bf") = "«Ã«“… »œÊ‰ —« »" And CheckBox1.value = True Then
Sheets(1).Cells(n3, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n3, "j")
Sheets(1).Cells(n3, "bt").value = TextBox2.Text
End If
Next








'========================================================


For n4 = 9 To lastR
If CheckBox6.value = True And Sheets(1).Cells(n4, "bf") = "≈Ã«“… „⁄Ì· «·„ ›—€" And CheckBox1.value = True Then
Sheets(1).Cells(n4, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n4, "j")
Sheets(1).Cells(n4, "bt").value = TextBox2.Text
End If
Next












'===================================================




For n5 = 9 To lastR
If CheckBox7.value = True And Sheets(1).Cells(n5, "bf") = "—›⁄ Ìœ" And CheckBox1.value = True Then
Sheets(1).Cells(n5, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n5, "j")
Sheets(1).Cells(n5, "bt").value = TextBox2.Text
End If
Next








'=======================================================




For n6 = 9 To lastR

If CheckBox8.value = True And Sheets(1).Cells(n6, "bf") = "«„Ê„Â «Ê· 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n6, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n6, "j")
Sheets(1).Cells(n6, "bt").value = TextBox2.Text
End If
Next








'=====================================================




For n7 = 9 To lastR
If CheckBox9.value = True And Sheets(1).Cells(n7, "bf") = "«„Ê„Â À«‰Ì 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n7, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n7, "j")
Sheets(1).Cells(n7, "bt").value = TextBox2.Text
End If
Next






'=====================================================





For n8 = 9 To lastR
If CheckBox10.value = True And Sheets(1).Cells(n8, "bf") = " ‰”Ì»" And CheckBox1.value = True Then
Sheets(1).Cells(n8, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n8, "j")
Sheets(1).Cells(n8, "bt").value = TextBox2.Text
End If
Next





'==================================================




For n9 = 9 To lastR
If CheckBox11.value = True And Sheets(1).Cells(n9, "bf") = "⁄ﬁœ" And CheckBox1.value = True Then
Sheets(1).Cells(n9, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n9, "j")
Sheets(1).Cells(n9, "bt").value = TextBox2.Text
End If
Next











'=======================================================







For n10 = 9 To lastR

If CheckBox12.value = True And Sheets(1).Cells(n10, "bf") = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" And CheckBox1.value = True Then
Sheets(1).Cells(n10, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n10, "j")
Sheets(1).Cells(n10, "bt").value = TextBox2.Text
End If
Next

















'===========================================





For n11 = 9 To lastR
If CheckBox13.value = True And Sheets(1).Cells(n11, "bf") = "—« » Ã“∆Ì" And CheckBox1.value = True Then
Sheets(1).Cells(n11, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n11, "j")
Sheets(1).Cells(n11, "bt").value = TextBox2.Text
End If
Next














'=======================================================
For n12 = 9 To lastR
If Sheets(1).Cells(n12, "h").value = 8 Or Sheets(1).Cells(n12, "h").value = 9 Or Sheets(1).Cells(n12, "h").value = 10 _
And CheckBox14.value = True And CheckBox1.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" Then

Sheets(1).Cells(n12, "br").value = val(TextBox1) + val(TextBox3) * Sheets(1).Cells(n12, "j")
 Sheets(1).Cells(n12, "bt").value = TextBox2.Text

End If
Next
Call count_all
MsgBox " „  «÷«›… «·‰”»Â"









End Sub

Private Sub CommandButton2_Click()
'==============================
If TextBox1.value = "" Then
TextBox1.value = 0
End If
Dim lastR As Integer
Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12 As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
'=============================================================

For n1 = 9 To lastR
If CheckBox3.value = True And Sheets(1).Cells(n1, "bf") = "„” „— »«·Œœ„…" And CheckBox1.value = True Then
Sheets(1).Cells(n1, "br").value = 0

End If

Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox3.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'===========================================================
For n2 = 9 To lastR
If CheckBox4.value = True And Sheets(1).Cells(n2, "bf") = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " And CheckBox1.value = True Then
Sheets(1).Cells(n2, "br").value = 0
End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox4.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'=================================================
For n3 = 9 To lastR
If CheckBox5.value = True And Sheets(1).Cells(n3, "bf") = "«Ã«“… »œÊ‰ —« »" And CheckBox1.value = True Then
Sheets(1).Cells(n3, "br").value = 0
End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox5.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'========================================================
For n4 = 9 To lastR
If CheckBox6.value = True And Sheets(1).Cells(n4, "bf") = "≈Ã«“… „⁄Ì· «·„ ›—€" And CheckBox1.value = True Then
Sheets(1).Cells(n4, "br").value = 0
TextBox3.value = 0
End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox6.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'===================================================
For n5 = 9 To lastR
If CheckBox7.value = True And Sheets(1).Cells(n5, "bf") = "—›⁄ Ìœ" And CheckBox1.value = True Then
Sheets(1).Cells(n5, "br").value = 0
TextBox3.value = 0
End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox7.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'=======================================================
For n6 = 9 To lastR
If CheckBox8.value = True And Sheets(1).Cells(n6, "bf") = "«„Ê„Â «Ê· 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n6, "br").value = 0
End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox8.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"


'=====================================================
For n7 = 9 To lastR
If CheckBox9.value = True And Sheets(1).Cells(n7, "bf") = "«„Ê„Â À«‰Ì 6 «‘Â—" And CheckBox1.value = True Then
Sheets(1).Cells(n7, "br").value = 0

End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox9.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'=====================================================
For n8 = 9 To lastR
If CheckBox10.value = True And Sheets(1).Cells(n8, "bf") = " ‰”Ì»" And CheckBox1.value = True Then
Sheets(1).Cells(n8, "br").value = 0

End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox10.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'==================================================
For n9 = 9 To lastR
If CheckBox11.value = True And Sheets(1).Cells(n9, "bf") = "⁄ﬁœ" And CheckBox1.value = True Then
Sheets(1).Cells(n9, "br").value = 0

End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox11.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"


'=======================================================
For n10 = 9 To lastR
If CheckBox12.value = True And Sheets(1).Cells(n10, "bf") = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" And CheckBox1.value = True Then
Sheets(1).Cells(n10, "br").value = 0
End If
Next
TextBox3.value = 0
 MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox12.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'===========================================
For n11 = 9 To lastR
If CheckBox13.value = True And Sheets(1).Cells(n11, "bf") = "—« » Ã“∆Ì" And CheckBox1.value = True Then
Sheets(1).Cells(n11, "br").value = 0
End If
Next
TextBox3.value = 0
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox13.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

'=======================================================
For n12 = 9 To lastR
If Sheets(1).Cells(n12, "h").value = 8 Or Sheets(1).Cells(n12, "h").value = 9 Or Sheets(1).Cells(n12, "h").value = 10 _
And CheckBox14.value = True And CheckBox1.value = True And Sheets(1).Cells(n12, "bf") = "„” „— »«·Œœ„…" And TextBox1.value = "" Then
TextBox1.value = 0
Sheets(1).Cells(n12, "br").value = 0

End If
Next
 TextBox3.value = 0
 Call count_all
MsgBox " „ «” ﬁÿ«⁄ «·‰”»Â „‰ «·›∆…" & "" & CheckBox14.Caption, vbDefaultButton1, "—”«·…  ÊÃÌÂ"

End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub

Private Sub Label5_Click()

End Sub

Private Sub TextBox1_Change()
Label4.Caption = "Â–Â «·ﬁÌ„Â  „À· „ﬁœ«— ‰”»… «·„∆ÊÌÂ „‰ „Œ’’«  «·—« » " & " " & Sheets(1).Cells(9, "e") & " " & _
"«·«”„ «·«Ê· „‰ »Ì«‰«  «·‰Ÿ«„"

End Sub

Private Sub TextBox3_AfterUpdate()

End Sub

Private Sub TextBox3_Change()
On Error Resume Next


End Sub

Private Sub UserForm_Activate()
TextBox4.value = Sheets(1).Cells(9, "br").value
End Sub

Private Sub UserForm_Click()

End Sub


