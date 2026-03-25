VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm35 
   Caption         =   "UserForm35"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9492.001
   OleObjectBlob   =   "UserForm35.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call Remove_Duplicates1
Call delete_change_month_form33
Call change_salary
Unload Me
End Sub

Private Sub CommandButton2_Click()
Call update_change_salary
End Sub



Private Sub CommandButton3_Click()
'=VLOOKUP(AL8:AM8;A8:AI13;35;0)
'Ê—ﬁ…33.Range("am8").Formula = "= MAXIFS(AC9:AC13,A9:A13,AL8,B9:B13,AK8)"
'Ê—ﬁ…33.Range("an8").Formula2R1C1 = "= VLOOKUP(RC[-2]:RC[-1],RC[-39]:R[5]C[-5],35,0)"
'Ê—ﬁ…33.Cells(8, "ak") = UserForm33.TextBox5.Text         ' «·«”„
'Ê—ﬁ…33.Cells(8, "al") = UserForm33.TextBox2.Text  '«·—ﬁ„ «·ÊŸÌ›Ì
Unload Me
End Sub

Private Sub CommandButton4_Click()
Call delete_change_month_form33
End Sub

Private Sub TextBox1_Change()
If TextBox1.Text = "" Then
CommandButton1.Enabled = True
Else
CommandButton1.Enabled = False
End If
End Sub

Private Sub UserForm_Activate()



Label1.Caption = "ÕÕœ ”»» «· ⁄œÌ· «·–Ì „‰ «Ã·Â  „ «· €ÌÌ— »Ì«‰«  «·„ÊŸ›" & ":" & UserForm33.TextBox5.Text
'==========================================

On Error Resume Next

    
   
  '==========================
  
  '= VLOOKUP(AL8:AM8,$A$8:$AI$300000,35,0)
'Ê—ﬁ…33.Range("am7").Formula = "= MAXIFS($BN$9:$BN$300000,$B$9:$B$300000,AL7;$E$9:$E$300000,AK7)"
'Ê—ﬁ…33.Range("an7").Formula2R1C1 = "= VLOOKUP(RC[-2]:RC[-1],R[1]C[-39]:R[4000]C[43],82,1)"
'Ê—ﬁ…33.Cells(7, "ak") = UserForm33.TextBox5.Text         ' «·«”„
'Ê—ﬁ…33.Cells(7, "al") = UserForm33.TextBox2.Text  '«·—ﬁ„ «·ÊŸÌ›Ì
'======================================================
Dim lastr2 As Integer
Dim y2 As Double
lastr2 = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
    For y2 = 9 To lastr2

 If Sheets(35).Cells(y2, 2).value = UserForm33.TextBox2.Text And Sheets(35).Cells(y2, 5).value = UserForm33.TextBox5.Text And Sheets(35).Cells(y2, "ce").value = "«Œ—  ÕœÌÀ" Then
UserForm35.TextBox2.Text = "Â–« «Œ—  «—ÌŒ  „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y2, "bn").value & "Ê „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y2, 82)
ElseIf Sheets(35).Cells(y2, 2).value <> UserForm33.TextBox2.Text And Sheets(35).Cells(y2, 5).value <> UserForm33.TextBox5.Text And Sheets(35).Cells(y2, "ce").value <> "«Œ—  ÕœÌÀ" Then
UserForm35.TextBox2.Text = "·«   Ê›— ·œÌÂ «Ì „ €Ì—«  :"
     
  End If
Next
  
  
  '=================================
'Sheets(35).Cells(8, "ak") = UserForm33.TextBox5.Text
'Sheets(35).Cells(8, "al") = UserForm33.TextBox2.Text
'Dim max_date As Date
'max_date = Sheets(35).Cells(8, "am")
'UserForm35.TextBox2.Text = Sheets(35).Cells(8, "an") & "  " & "Ê«Œ—  ÕœÌÀ ﬂ«‰ » «—ÌŒ" & "  " & Sheets(35).Cells(8, "am")

Dim lastr1 As Integer
Dim y1 As Double
lastr1 = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
    For y1 = 9 To lastr1
    
    If Sheets(35).Cells(y1, 2) = UserForm33.TextBox2.Text And Sheets(35).Cells(y1, 5) = UserForm33.TextBox5.Text And Sheets(35).Cells(y1, "ce") = "«Œ—  ÕœÌÀ" Then
 TextBox3.value = Sheets(35).Cells(y1, 31).value
  TextBox4.value = Sheets(35).Cells(y1, 78).value
  
    If Sheets(35).Cells(y1, 81).value > 0 Then

 TextBox5.value = Sheets(35).Cells(y1, 81).value
 Else
 TextBox6.value = Sheets(35).Cells(y1, 81).value
     
     End If
    End If
    Next y1

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


End Sub










Private Sub UserForm_Initialize()


End Sub

