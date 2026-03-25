VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm27 
   Caption         =   "UserForm27"
   ClientHeight    =   8916.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10692
   OleObjectBlob   =   "UserForm27.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Private Sub CheckBox1_Click()
'If CheckBox1.Value = True Then
'CheckBox2.Value = False
'CommandButton1.Enabled = True
'ComboBox1.DropDown
'ComboBox1.Value = "Õ–›"
'TextBox7.Text = "Õ–›"
'End If
End Sub

Private Sub CheckBox2_Click()
'If CheckBox2.Value = True Then
'CheckBox1.Value = False
'CommandButton1.Enabled = True
'ComboBox1.DropDown
'ComboBox1.Value = "«÷«›…"
'TextBox7.Text = "«÷«›…"
'End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.value = True Then
CommandButton4.Enabled = True
Else
CommandButton4.Enabled = False
End If
End Sub

Private Sub ComboBox1_AfterUpdate()
TextBox7.Text = ComboBox1.value
TextBox1.Text = ""
If TextBox7.Text = "«÷«›…" Then
ComboBox2.Enabled = True
ComboBox3.Enabled = False

Else
ComboBox2.Enabled = False
ComboBox3.Enabled = True
End If
End Sub

Private Sub ComboBox1_Change()
TextBox7.Text = ComboBox1.value
TextBox1.Text = ""
If TextBox7.Text = "«÷«›…" Then
ComboBox3.Enabled = False
Else
ComboBox2.Enabled = False
ComboBox3.Enabled = True
End If
End Sub

Private Sub ComboBox2_AfterUpdate()
 
 ' ComboBox3.Value = ""
  ' ComboBox4.Value = ""
 TextBox1.Text = ""
' TextBox1.Text = ComboBox2.Value & " " & ComboBox3.Value & " " & "·‘Â—" & " " & ComboBox4.Value
End Sub

Private Sub ComboBox2_Change()
TextBox1.Text = ""
 TextBox1.Text = ComboBox2.value
  
TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & ComboBox4.value
End Sub

Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ComboBox3_AfterUpdate()
 
' ComboBox2.Value = ""
'ComboBox4.Value = ""
TextBox1.Text = ""

'TextBox1.Text = ComboBox2.Value & " " & ComboBox3.Value & " " & "·‘Â—" & " " & " " & ComboBox4.Value
End Sub

Private Sub ComboBox3_Change()
TextBox1.Text = ""

TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & " " & ComboBox4.value
End Sub

Private Sub ComboBox4_AfterUpdate()

 TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & ComboBox4.value
End Sub

Private Sub ComboBox4_Change()
TextBox1.Text = ""
TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & ComboBox4.value
End Sub

Private Sub CommandButton1_Click()
Call Remove_Duplicates1
'Call delete_change_month1
Call change_salary
MsgBox " „  ‰›Ì– »Ì«‰«  «·ﬁÌœ «·ﬁœÌ„ ·—« » «·‘Â— «·„«÷Ì", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
CommandButton2.Enabled = True
CommandButton1.Enabled = False
'Dim LastRow As Long
       
  '  LastRow = Sheets(35).Cells(Rows.Count, 1).End(xlUp).Row

'TextBox9.Value = Sheets(35).Cells(LastRow, 1).Value
Dim lastr110, n As Integer
lastr110 = Sheets(35).Cells(Rows.count, "a").End(xlUp).row

For n = 9 To lastr110
TextBox9.Text = Sheets(35).Cells(n, 84).value

Next

End Sub

Private Sub CommandButton2_Click()
Call update_change_salary
CommandButton2.Enabled = False
End Sub



Private Sub CommandButton3_Click()
'=VLOOKUP(AL8:AM8;A8:AI13;35;0)
'Ê—ﬁ…33.Range("am8").Formula = "= MAXIFS(AC9:AC13,A9:A13,AL8,B9:B13,AK8)"
'Ê—ﬁ…33.Range("an8").Formula2R1C1 = "= VLOOKUP(RC[-2]:RC[-1],RC[-39]:R[5]C[-5],35,0)"
'Ê—ﬁ…33.Cells(8, "ak") = UserForm1.TextBox5.Text         ' «·«”„
'Ê—ﬁ…33.Cells(8, "al") = UserForm1.TextBox2.Text  '«·—ﬁ„ «·ÊŸÌ›Ì
Unload Me
End Sub

Private Sub CommandButton4_Click()
Sheets(35).Activate
Call delete_change_month1

End Sub

Private Sub CommandButton5_Click()
 Dim lastRow As Long
    Dim i As Long
    
    
    lastRow = Sheets(35).Cells(Rows.count, 6).End(xlUp).row
    
   
    For i = 9 To lastRow
     If Sheets(35).Cells(i, 2).value = "" And Sheets(35).Cells(i, 5).value = "" And Sheets(35).Cells(i, 6).value = UserForm1.ComboBox5.value And Sheets(35).Cells(i, 7).value = UserForm1.TextBox7.value Then
           Sheets(35).Rows(i).Delete
        End If
    Next i
    MsgBox "ﬁœ  „ Õ–› «·ﬁÌœ »‰Ã«Õ Õ«Ê· «‰ ·«  ‰”Ï «·«Ã—«¡ «·’ÕÌÕ ÂÊ ﬁ„ »«· ‰›Ì– À„ «· ÕœÌÀ ‘ﬂ—« ·ﬂ", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
End Sub







Private Sub CommandButton6_Click()
Unload Me
UserForm71.Show
End Sub

Private Sub TextBox1_Change()
Call arb

End Sub

Private Sub TextBox7_Change()
ComboBox2.value = ""
ComboBox3.value = ""
If TextBox7.Text = "«÷«›…" Then
ComboBox2.Enabled = True
ComboBox3.Enabled = False
TextBox1.Text = ""
Else
ComboBox2.Enabled = False
ComboBox3.Enabled = True
TextBox1.Text = ""
End If
End Sub

Private Sub UserForm_Activate()

Call arb
CheckBox3.value = False
CommandButton2.Enabled = False





'================================================
Dim lastr100, i1 As Integer
lastr100 = Sheets(35).Cells(Rows.count, "f").End(xlUp).row
    For i1 = 9 To lastr100
If Sheets(35).Cells(i1, 2).value = "" And Sheets(35).Cells(i1, 5).value = "" And Sheets(35).Cells(i1, 6).value = UserForm1.ComboBox5.value And Sheets(35).Cells(i1, 7).value = UserForm1.TextBox7.value Then
TextBox8.Text = "Â‰«ﬂ Œÿ√ ›Ì «· ‰›Ì– ﬁœ ﬁ„  » ÕœÌÀ «·»Ì«‰«   ﬁ»· «‰  ﬁÊ„ » ‰›Ì– »Ì«‰«  «·‘Â— «·„«÷Ì ﬁ„ »Õ–› «·ﬁÌœ Ê«⁄«œ…  —ÕÌ·Â „‰ ÃœÌœ"
CommandButton5.Visible = True
ElseIf Sheets(35).Cells(i1, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(i1, 5).value = UserForm1.TextBox5.Text Then

TextBox8.Text = "«·ﬁÌœ „ÕœÀ Ê „  —ÕÌ·Â »‰Ã«Õ"
CommandButton5.Visible = False
End If
Next











'===========================================================

ComboBox2.Enabled = False
ComboBox3.Enabled = False

ComboBox1.value = ""
ComboBox2.value = ""
ComboBox3.value = ""
ComboBox4.value = ""

ComboBox1.List = Array("«÷«›…", "Õ–›")
ComboBox2.List = Array("«÷«›… ⁄·«ÊÂ ÃœÌœÂ", "«÷«›… œ—Ã… Ê —›Ì⁄", " €ÌÌ— ‰”»… ‘Â«œ…", " €Ì— ‰”»… Â‰œ”Ì…", "«÷«›… „‰’»", "«÷«›… „Êﬁ⁄ Ã€—«›Ì", "«÷«›… „Œ’’«  «ÿ›«·", _
"«÷«›… „Œ’’«  “ÊÃÌÂ ", "«÷«›… „Œ’’«  “ÊÃÌÂ „⁄ «ÿ›«·", "«÷«›… „Œ’’«  „Â‰Ì…", "«÷«›… „Œ’’«  ŒÿÊ—Â ··”Ê«ﬁ", "«÷«›… „Œ’’«  ŒÿÊ—… Õ«”»« ")
ComboBox3.List = Array("Õ–› „Œ’’«  „‰’»", "Œ–› „Œ’’«  „Êﬁ⁄ Ã€—«›Ì", "Õ–› „Œ’’«  «ÿ›«·", "Õ–› „Œ’’«  “ÊÃÌ…", "Õ–› „Œ’’«  «ÿ›«· „⁄ “ÊÃÌ…", "ÕÃ» „Œ’’«  „Â‰ÌÂ", _
"ÕÃ» „Œ’’«  ŒÿÊ—Â ··”Ê«ﬁ", "ÕÃ» „Œ’’«  ŒÿÊ—Â Õ«”»«  ‰”»…")
Dim nammonth1 As Date
Dim i As Date



For i = 1 To 12
nammonth = monthName(i, abbreviate:=False)
ComboBox4.AddItem nammonth
Next
Label1.Caption = "Õœœ ”»» «· ⁄œÌ· «·–Ì „‰ «Ã·Â  „ «· €ÌÌ— »Ì«‰«  «·„ÊŸ›" & ":" & UserForm1.TextBox5.Text
'==========================================

On Error Resume Next

    
   
  '==========================
  
  '= VLOOKUP(AL8:AM8,$A$8:$AI$300000,35,0)
'Ê—ﬁ…33.Range("am7").Formula = "= MAXIFS($BN$9:$BN$300000,$B$9:$B$300000,AL7;$E$9:$E$300000,AK7)"
'Ê—ﬁ…33.Range("an7").Formula2R1C1 = "= VLOOKUP(RC[-2]:RC[-1],R[1]C[-39]:R[4000]C[43],82,1)"
'Ê—ﬁ…33.Cells(7, "ak") = UserForm1.TextBox5.Text         ' «·«”„
'Ê—ﬁ…33.Cells(7, "al") = UserForm1.TextBox2.Text  '«·—ﬁ„ «·ÊŸÌ›Ì
'======================================================
Dim lastr001 As Integer
Dim y2, y3, y4 As Integer
lastr001 = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
    For y2 = 9 To lastr001

 If Sheets(35).Cells(y2, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(y2, 5).value = UserForm1.TextBox5.Text And Sheets(35).Cells(y2, "ce").value = "Õ–›" Then
TextBox2.Text = "Â–« «Œ—  «—ÌŒ  „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y2, "bn").value & " " & "Ê „  ÕœÌÀ »Ì«‰« Â" & "   " & Sheets(35).Cells(y2, 82) & "  " & "Ê«·„ €Ì—«  ›Ì Õ«·…" & " " & Sheets(35).Cells(y2, "ce")
 TextBox3.value = Sheets(35).Cells(y2, 31).value
  TextBox4.value = Sheets(35).Cells(y2, 78).value
 TextBox5.value = Sheets(35).Cells(y2, 81).value
TextBox6.value = ""
TextBox7.value = Sheets(35).Cells(y2, 83).value

End If
Next
'===================================================
 For y3 = 9 To lastr001
If Sheets(35).Cells(y3, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(y3, 5).value = UserForm1.TextBox5.Text And Sheets(35).Cells(y3, "ce").value = "«÷«›…" Then
TextBox2.Text = "Â–« «Œ—  «—ÌŒ  „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y3, "bn").value & "Ê „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y3, 82) & "Ê«·„ €Ì—«  ›Ì Õ«·…" & Sheets(35).Cells(y3, "ce")
TextBox3.value = Sheets(35).Cells(y3, 31).value
  TextBox4.value = Sheets(35).Cells(y3, 78).value
 TextBox5.value = ""
TextBox6.value = Sheets(35).Cells(y3, 81).value
TextBox7.value = Sheets(35).Cells(y3, 83).value

End If
Next
'===============================
For y4 = 9 To lastr2
If Sheets(35).Cells(y4, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(y4, 5).value = UserForm1.TextBox5.Text _
And Sheets(35).Cells(y4, "ce").value <> "«÷«›…" And Sheets(35).Cells(y4, "ce").value <> "Õ–›" Then
UserForm27.TextBox2.Text = "·«   Ê›— ·œÌÂ «Ì „ €Ì—«  :"
    End If
Next
  
  
  '=================================



If TextBox3.value = "" And TextBox4.value = "" Then

CommandButton4.Enabled = False
CommandButton1.Enabled = True
Else
CommandButton4.Enabled = True
CommandButton1.Enabled = False
End If

If TextBox3.value > 1 And TextBox4.value = "" Then

CommandButton4.Enabled = False
End If
'===================================

If TextBox7.Text = "«÷«›…" Then
CheckBox2.value = True
CheckBox1.value = False
CheckBox1.Enabled = False

If CheckBox2.value = True Then
CheckBox1.value = False
CommandButton1.Enabled = True
ComboBox1.DropDown
ComboBox1.value = "Õ–›"
TextBox7.Text = "Õ–›"
End If





ElseIf TextBox7.Text = "Õ–›" Then
CheckBox1.value = True
CheckBox2.value = False
CheckBox2.Enabled = False
If CheckBox1.value = True Then
CheckBox2.value = False
CommandButton1.Enabled = True
ComboBox1.DropDown
ComboBox1.value = "«÷«›…"
TextBox7.Text = "«÷«›…"
End If




ElseIf TextBox7.Text = "" Then
CheckBox1.value = False
CheckBox2.value = False
CheckBox1.Enabled = False
CheckBox2.Enabled = False
End If
End Sub










Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()


End Sub


