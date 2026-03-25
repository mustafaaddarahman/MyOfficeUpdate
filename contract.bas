Attribute VB_Name = "contract"
Sub contract_1_22()
On Error Resume Next

'Dim vlookup1 As Long
'Dim vlookup2 As Long
'Dim vlookup3 As Long

'Dim textvlue1 As Long
'Dim textvlue2 As Long
'Dim combvlue As Long
'Dim rang1 As Range
'Dim rang2 As Range



'„»·€ «·€Ì«»

'userform1.textbox49.Text = userform1.textbox49.Text - userform1.textbox48.Text
'=============================== «·«” ﬁÿ«⁄« 





'=================================================
If UserForm1.TextBox101.value = "" Then
MsgBox "ﬁ„ «Ê·« »«œŒ«· ﬁÌ„… —« » «·⁄ﬁÊœ ·Ì ”‰Ï «Õ ”«»  ›«’Ì· «·—« »", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Exit Sub
Else
UserForm1.TextBox8.value = 101
UserForm1.TextBox10.value = 0
UserForm1.TextBox11.value = 0
UserForm1.TextBox14.value = 0
UserForm1.TextBox16.value = 0
UserForm1.TextBox18.value = 0
UserForm1.TextBox19.value = 0
UserForm1.TextBox21.value = 0
UserForm1.TextBox20.value = 0
UserForm1.TextBox24.value = 0
UserForm1.TextBox25.value = 0
UserForm1.TextBox27.value = 0
UserForm1.TextBox29.value = 0
UserForm1.TextBox32.Text = 0
UserForm1.TextBox33.Text = 0
UserForm1.TextBox34.value = 0
UserForm1.TextBox35.value = 0
UserForm1.TextBox36.value = 0
UserForm1.TextBox38.value = 0
UserForm1.TextBox39.value = 0
UserForm1.TextBox40.value = 0
UserForm1.TextBox42.value = 0
UserForm1.TextBox44.value = 0
UserForm1.TextBox46.value = 0
UserForm1.TextBox47.value = 0
UserForm1.TextBox60.value = 0
UserForm1.TextBox62.value = 0
UserForm1.TextBox83.value = 0
UserForm1.TextBox84.value = 0
UserForm1.TextBox85.value = 0
UserForm1.TextBox86.value = 0
UserForm1.TextBox89.value = 0





UserForm1.TextBox31.value = 0
UserForm1.TextBox30.value = 0
UserForm1.ComboBox6.value = 0
UserForm1.ComboBox7.value = 0
UserForm1.ComboBox8.value = 0
UserForm1.ComboBox10.value = 0
UserForm1.ComboBox11.value = 0





'==========================„Ã„Ê⁄ «·„Œ’’« 
'Dim numlong As Long
'Dim numlong1 As Long
'numlong = 10000
'UserForm1.TextBox20.Value = numlong1
'If UserForm1.TextBox22.Value >= 5 Then
'numlong1 = 40000
'UserForm1.TextBox20.Value = numlong1
'ElseIf UserForm1.TextBox22.Value <= 4 Then
'UserForm1.TextBox20.Value = Val(UserForm1.TextBox22) * numlong
'End If



UserForm1.TextBox30.value = 0

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 


UserForm1.TextBox31.value = val(UserForm1.TextBox101) + val(UserForm1.TextBox30)






'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ
UserForm1.TextBox32.value = val(UserForm1.TextBox101) * UserForm1.TextBox104.value
'UserForm1.TextBox33.Value = Val(UserForm1.TextBox101) * 0.12


'==================== „»·€ «·€Ì«»

'Call state_tax_akood

'===============«·«” ﬁÿ«⁄« 
'If UserForm1.TextBox101.Value > 500000 Then

' UserForm1.TextBox42.Value = 1500
' Else
' UserForm1.TextBox42.Value = 1000
'End If

Dim sumastktaa As Double
sumastktaa = val(UserForm1.TextBox24) + val(UserForm1.TextBox32) + val(UserForm1.TextBox34) + val(UserForm1.TextBox35) + val(UserForm1.TextBox36) _
+ val(UserForm1.TextBox38) + val(UserForm1.TextBox40) + val(UserForm1.TextBox42) + val(UserForm1.TextBox44) + val(UserForm1.TextBox46) _
 + val(UserForm1.TextBox50) + val(UserForm1.TextBox62) + val(UserForm1.TextBox83) + val(UserForm1.TextBox85)
 
 
 UserForm1.TextBox48.value = sumastktaa
UserForm1.TextBox49.value = val(UserForm1.TextBox31) - val(UserForm1.TextBox48)
UserForm1.TextBox48.value = val(UserForm1.TextBox47) + sumastktaa
'=======================«·—« » «·’«›Ì
UserForm1.TextBox49.value = val(UserForm1.TextBox31) - val(UserForm1.TextBox48)









'=======================================================

   
   End If
  ' Call emad_pro
        
Call update_data_acount_22
End Sub
