Attribute VB_Name = "mother_hood_2_6_1"
Sub mother_hood_2_6_1_1()

'Dim vlookup1 As Long
'Dim vlookup2 As Long
'Dim vlookup3 As Long

'Dim textvlue1 As Long
'Dim textvlue2 As Long
'Dim combvlue As Long
'Dim rang1 As Range
'Dim rang2 As Range

'=================================================
With UserForm48
.TextBox10.value = .TextBox10.value / 2
.TextBox14.value = val(.ComboBox6) * val(.TextBox10)
.TextBox16.value = 0
.TextBox18.value = 0
.TextBox19.value = 0
.TextBox20.value = val(.TextBox22) * 10000
.TextBox25.value = 0
.ComboBox10.value = 0
.TextBox27.value = 0
.ComboBox11.value = 0
.TextBox29.value = 0
.TextBox101.value = 0  '⁄ﬁÊœ
.TextBox33.value = 0
.TextBox35.value = 0
.TextBox36.value = 0
.TextBox47.value = 0
'==========================„Ã„Ê⁄ «·„Œ’’« 
Dim numlong As Long
Dim numlong1 As Long
numlong = 10000
.TextBox20.value = numlong1
If .TextBox22.value >= 5 Then
numlong1 = 40000
.TextBox20.value = numlong1
ElseIf .TextBox22.value <= 4 Then
.TextBox20.value = val(.TextBox22) * numlong
End If



.TextBox30.value = val(.TextBox14) + val(.TextBox21) + val(.TextBox20) _
 + val(.TextBox84) + val(.TextBox86) + val(.TextBox19)

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 


.TextBox31.value = val(.TextBox10) + val(.TextBox30)






'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ
.TextBox32.value = val(.TextBox10) * 0.1
'.TextBox33.Value = Val(.TextBox10) * 0.15


'==================== „»·€ «·€Ì«»

Call state_tax_22


'===============«·«” ﬁÿ«⁄« 
If .TextBox10.value > 500000 Then

 .TextBox42.value = 1500
 Else
 .TextBox42.value = 1000
End If
.TextBox46.value = val(.TextBox10) * 0.0025
Dim sumastktaa As Double
sumastktaa = val(.TextBox24) + val(.TextBox32) + val(.TextBox33) + val(.TextBox34) + val(.TextBox35) + val(.TextBox36) _
+ val(.TextBox38) + val(.TextBox40) + val(.TextBox42) + val(.TextBox44) + val(.TextBox46) _
 + val(.TextBox50) + val(.TextBox62) + val(.TextBox83) + val(.TextBox85)
 .TextBox48.value = sumastktaa

'=======================«·—« » «·’«›Ì
.TextBox49.value = val(.TextBox31) - val(.TextBox48)










'=======================================================

   
   End With
 Call update_data_acount_1
End Sub


