Attribute VB_Name = "five_year"
Sub five_snwat_22()

On Error Resume Next

MsgBox "·œÌÂ «Ã«“Â Œ„” ”‰Ê« ", vbDefaultButton2, "—”«·…  ÊÃÌÂ"



' ’›Ì— ﬂ· «·„Œ’’« 
  UserForm1.TextBox32.value = val(UserForm1.TextBox10) * 0.1
UserForm1.TextBox33.value = val(UserForm1.TextBox10.value) * 0.15

  UserForm1.TextBox11.Text = 0
 UserForm1.TextBox12.Text = 0
 
 UserForm1.TextBox14.Text = 0
 
 UserForm1.TextBox16.Text = 0
 
 UserForm1.TextBox18.Text = 0
 
  UserForm1.TextBox19.value = 0
 UserForm1.TextBox20.Text = 0
  UserForm1.TextBox21.Text = 0
 
 ' userform1.textbox24.Text = 0
  UserForm1.TextBox25.Text = 0
  
 UserForm1.TextBox47.Text = 0
UserForm1.TextBox27.Text = 0
 UserForm1.TextBox29.Text = 0
 UserForm1.TextBox30.Text = 0
 UserForm1.TextBox35.value = 0
 UserForm1.TextBox38.Text = 0
'userform1.textbox34.Text = 0

 'userform1.textbox42.Text = 0
 
 

'==========================„Ã„Ê⁄ «·„Œ’’« 
UserForm1.TextBox30.value = val(UserForm1.TextBox14.value) + val(UserForm1.TextBox16.value) + val(UserForm1.TextBox18.value) + val(UserForm1.TextBox19.value) + val(UserForm1.TextBox21.value) + val(UserForm1.TextBox20.value) + val(UserForm1.TextBox25.value) + val(UserForm1.TextBox27.value) + val(UserForm1.TextBox29.value)
'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 

UserForm1.TextBox31.value = val(UserForm1.TextBox10.value) + val(UserForm1.TextBox14.value) + val(UserForm1.TextBox16.value) + val(UserForm1.TextBox18.value) + val(UserForm1.TextBox19.value) _
+ val(UserForm1.TextBox21.value) + val(UserForm1.TextBox20.value) + val(UserForm1.TextBox25.value) + val(UserForm1.TextBox27.value) + val(UserForm1.TextBox29.value)
UserForm1.TextBox32.value = CDbl(UserForm1.TextBox10.value * 0.1)
UserForm1.TextBox33.value = CDbl(UserForm1.TextBox10.value * 0.15)
'============================«Õ ”«» «·÷—Ì»…
Call state_tax


'===============«·«” ﬁÿ«⁄« 
If UserForm1.TextBox10.value > 500000 Then

 UserForm1.TextBox42.value = 1500
 Else
 UserForm1.TextBox42.value = 1000
End If
UserForm1.TextBox46.value = val(UserForm1.TextBox10) * 0.0025

UserForm1.TextBox48.value = val(UserForm1.TextBox24) + val(UserForm1.TextBox32) + val(UserForm1.TextBox33) + val(UserForm1.TextBox34) + val(UserForm1.TextBox35) + val(UserForm1.TextBox36) _
+ val(UserForm1.TextBox38) + val(UserForm1.TextBox40) + val(UserForm1.TextBox42) + val(UserForm1.TextBox44) + val(UserForm1.TextBox46) _
  + val(UserForm1.TextBox50) + val(UserForm1.TextBox62) + val(UserForm1.TextBox83) + val(UserForm1.TextBox85)






'=======================«·—« » «·’«›Ì
UserForm1.TextBox49.Text = CDbl(UserForm1.TextBox31.Text) - CDbl(UserForm1.TextBox48.Text)
'=======================

'==================== „»·€ «·€Ì«»






'========================================

Call update_data_acount_22

End Sub

