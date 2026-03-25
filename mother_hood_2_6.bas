Attribute VB_Name = "mother_hood_2_6"
Sub mother_hood_2_6_1_22()

'Dim vlookup1 As Long
'Dim vlookup2 As Long
'Dim vlookup3 As Long

'Dim textvlue1 As Long
'Dim textvlue2 As Long
'Dim combvlue As Long
'Dim rang1 As Range
'Dim rang2 As Range

'=================================================

UserForm1.TextBox10.value = UserForm1.TextBox10.value / 2
UserForm1.TextBox14.value = val(UserForm1.ComboBox6) * val(UserForm1.TextBox10)
UserForm1.TextBox16.value = 0
UserForm1.TextBox18.value = 0
UserForm1.TextBox19.value = 0
UserForm1.TextBox20.value = val(UserForm1.TextBox22) * 10000
UserForm1.TextBox25.value = 0
UserForm1.ComboBox10.value = 0
UserForm1.TextBox27.value = 0
UserForm1.ComboBox11.value = 0
UserForm1.TextBox29.value = 0
UserForm1.TextBox101.value = 0  '⁄ﬁÊœ
UserForm1.TextBox33.value = 0
UserForm1.TextBox35.value = 0
UserForm1.TextBox36.value = 0
UserForm1.TextBox47.value = 0
'==========================„Ã„Ê⁄ «·„Œ’’« 
Dim numlong As Long
Dim numlong1 As Long
numlong = 10000
UserForm1.TextBox20.value = numlong1
If UserForm1.TextBox22.value >= 5 Then
numlong1 = 40000
UserForm1.TextBox20.value = numlong1
ElseIf UserForm1.TextBox22.value <= 4 Then
UserForm1.TextBox20.value = val(UserForm1.TextBox22) * numlong
End If



UserForm1.TextBox30.value = val(UserForm1.TextBox14) + val(UserForm1.TextBox21) + val(UserForm1.TextBox20) _
 + val(UserForm1.TextBox84) + val(UserForm1.TextBox86) + val(UserForm1.TextBox19)

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 


UserForm1.TextBox31.value = val(UserForm1.TextBox10) + val(UserForm1.TextBox30)






'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ
UserForm1.TextBox32.value = val(UserForm1.TextBox10) * 0.1
'UserForm1.TextBox33.Value = Val(UserForm1.TextBox10) * 0.15


'==================== „»·€ «·€Ì«»

Call state_tax


'===============«·«” ﬁÿ«⁄« 
If UserForm1.TextBox10.value > 500000 Then

 UserForm1.TextBox42.value = 1500
 Else
 UserForm1.TextBox42.value = 1000
End If
UserForm1.TextBox46.value = val(UserForm1.TextBox10) * 0.0025
Dim sumastktaa As Double
sumastktaa = val(UserForm1.TextBox24) + val(UserForm1.TextBox32) + val(UserForm1.TextBox33) + val(UserForm1.TextBox34) + val(UserForm1.TextBox35) + val(UserForm1.TextBox36) _
+ val(UserForm1.TextBox38) + val(UserForm1.TextBox40) + val(UserForm1.TextBox42) + val(UserForm1.TextBox44) + val(UserForm1.TextBox46) _
 + val(UserForm1.TextBox50) + val(UserForm1.TextBox62) + val(UserForm1.TextBox83) + val(UserForm1.TextBox85)
 UserForm1.TextBox48.value = sumastktaa

'=======================«·—« » «·’«›Ì
UserForm1.TextBox49.value = val(UserForm1.TextBox31) - val(UserForm1.TextBox48)










'=======================================================
Call update_data_acount_22
End Sub
