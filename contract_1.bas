Attribute VB_Name = "contract_1"
Sub contract_1_1()
On Error Resume Next
With UserForm48
'Dim vlookup1 As Long
'Dim vlookup2 As Long
'Dim vlookup3 As Long

'Dim textvlue1 As Long
'Dim textvlue2 As Long
'Dim combvlue As Long
'Dim rang1 As Range
'Dim rang2 As Range



'„»·€ «·€Ì«»

'.textbox49.Text = .textbox49.Text - .textbox48.Text
'=============================== «·«” ﬁÿ«⁄« 





'=================================================
If .TextBox101.value = "" Then
MsgBox "ﬁ„ «Ê·« »«œŒ«· ﬁÌ„… —« » «·⁄ﬁÊœ ·Ì ”‰Ï «Õ ”«»  ›«’Ì· «·—« »", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Exit Sub
Else
.TextBox8.value = 101
.TextBox10.value = 0
.TextBox11.value = 0
.TextBox14.value = 0
.TextBox16.value = 0
.TextBox18.value = 0
.TextBox19.value = 0
.TextBox21.value = 0
.TextBox20.value = 0
.TextBox24.value = 0
.TextBox25.value = 0
.TextBox27.value = 0
.TextBox29.value = 0
.TextBox32.Text = 0
.TextBox33.Text = 0
.TextBox34.value = 0
.TextBox35.value = 0
.TextBox36.value = 0
.TextBox38.value = 0
.TextBox39.value = 0
.TextBox40.value = 0
.TextBox42.value = 0
.TextBox44.value = 0
.TextBox46.value = 0
.TextBox47.value = 0
.TextBox60.value = 0
.TextBox62.value = 0
.TextBox83.value = 0
.TextBox84.value = 0
.TextBox85.value = 0
.TextBox86.value = 0
.TextBox89.value = 0





.TextBox31.value = 0
.TextBox30.value = 0
.ComboBox6.value = 0
.ComboBox7.value = 0
.ComboBox8.value = 0
.ComboBox10.value = 0
.ComboBox11.value = 0





'==========================„Ã„Ê⁄ «·„Œ’’« 
'Dim numlong As Long
'Dim numlong1 As Long
'numlong = 10000
'.TextBox20.Value = numlong1
'If .TextBox22.Value >= 5 Then
'numlong1 = 40000
'.TextBox20.Value = numlong1
'ElseIf .TextBox22.Value <= 4 Then
'.TextBox20.Value = Val(.TextBox22) * numlong
'End If



.TextBox30.value = 0

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 


.TextBox31.value = val(.TextBox101) + val(.TextBox30)






'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ
.TextBox32.value = val(.TextBox101) * .TextBox104.value
'.TextBox33.Value = Val(.TextBox101) * 0.12


'==================== „»·€ «·€Ì«»

'Call state_tax_akood

'===============«·«” ﬁÿ«⁄« 
'If .TextBox101.Value > 500000 Then

' .TextBox42.Value = 1500
' Else
' .TextBox42.Value = 1000
'End If

Dim sumastktaa As Double
sumastktaa = val(.TextBox24) + val(.TextBox32) + val(.TextBox34) + val(.TextBox35) + val(.TextBox36) _
+ val(.TextBox38) + val(.TextBox40) + val(.TextBox42) + val(.TextBox44) + val(.TextBox46) _
 + val(.TextBox50) + val(.TextBox62) + val(.TextBox83) + val(.TextBox85)
 
 
 .TextBox48.value = sumastktaa
.TextBox49.value = val(.TextBox31) - val(.TextBox48)
.TextBox48.value = val(.TextBox47) + sumastktaa
'=======================«·—« » «·’«›Ì
.TextBox49.value = val(.TextBox31) - val(.TextBox48)









'=======================================================

   
   End If
  ' Call emad_pro
    End With
Call update_data_acount_1

End Sub


