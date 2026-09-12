Attribute VB_Name = "five_year_1"
Sub five_snwat_1()

On Error Resume Next

MsgBox "·œÌÂ «Ã«“Â Œ„” ”‰Ê« ", vbDefaultButton2, "—”«·…  ÊÃÌÂ"

With UserForm48

' ’›Ì— ﬂ· «·„Œ’’« 
  .TextBox32.value = val(.TextBox10) * 0.1
.TextBox33.value = val(.TextBox10.value) * 0.15

  .TextBox11.Text = 0
 .TextBox12.Text = 0
 
 .TextBox14.Text = 0
 
 .TextBox16.Text = 0
 
 .TextBox18.Text = 0
 
  .TextBox19.value = 0
 .TextBox20.Text = 0
  .TextBox21.Text = 0
 
 ' .textbox24.Text = 0
  .TextBox25.Text = 0
  
 .TextBox47.Text = 0
.TextBox27.Text = 0
 .TextBox29.Text = 0
 .TextBox30.Text = 0
 .TextBox35.value = 0
 .TextBox38.Text = 0
'.textbox34.Text = 0

 '.textbox42.Text = 0
 
 

'==========================„Ã„Ê⁄ «·„Œ’’« 
.TextBox30.value = val(.TextBox14.value) + val(.TextBox16.value) + val(.TextBox18.value) + val(.TextBox19.value) + val(.TextBox21.value) + val(.TextBox20.value) + val(.TextBox25.value) + val(.TextBox27.value) + val(.TextBox29.value)
'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 

.TextBox31.value = val(.TextBox10.value) + val(.TextBox14.value) + val(.TextBox16.value) + val(.TextBox18.value) + val(.TextBox19.value) _
+ val(.TextBox21.value) + val(.TextBox20.value) + val(.TextBox25.value) + val(.TextBox27.value) + val(.TextBox29.value)
.TextBox32.value = CDbl(.TextBox10.value * 0.1)
.TextBox33.value = CDbl(.TextBox10.value * 0.15)
'============================«Õ ”«» «·÷—Ì»…
Call state_tax_22


'===============«·«” ﬁÿ«⁄« 
If .TextBox10.value > 500000 Then

 .TextBox42.value = 1500
 Else
 .TextBox42.value = 1000
End If
.TextBox46.value = val(.TextBox10) * 0.0025

.TextBox48.value = val(.TextBox24) + val(.TextBox32) + val(.TextBox33) + val(.TextBox34) + val(.TextBox35) + val(.TextBox36) _
+ val(.TextBox38) + val(.TextBox40) + val(.TextBox42) + val(.TextBox44) + val(.TextBox46) _
  + val(.TextBox50) + val(.TextBox62) + val(.TextBox83) + val(.TextBox85)






'=======================«·—« » «·’«›Ì
.TextBox49.Text = CDbl(.TextBox31.Text) - CDbl(.TextBox48.Text)
'=======================

'==================== „»·€ «·€Ì«»






'========================================

'.textbox49.Text = .textbox31.Text - .textbox48.Text

'Call update_data_acount
End With
Call update_data_acount_1

End Sub



