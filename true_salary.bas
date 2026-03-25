Attribute VB_Name = "true_salary"
Sub true_salary_1()
'On Error Resume Next
UserForm1.TextBox30.value = val(UserForm1.TextBox14.value) + val(UserForm1.TextBox16.value) + val(UserForm1.TextBox18.value) + val(UserForm1.TextBox19.value) + val(UserForm1.TextBox21.value) + val(UserForm1.TextBox20.value) + val(UserForm1.TextBox25.value) + val(UserForm1.TextBox27.value) + val(UserForm1.TextBox29.value)
'============================???????? ?????????

'============================????? ???????????
Dim ast78a8 As Double
ast78a8 = val(UserForm1.TextBox10.value) + val(UserForm1.TextBox14.value) + val(UserForm1.TextBox16.value) + val(UserForm1.TextBox18.value) + val(UserForm1.TextBox19.value) _
+ val(UserForm1.TextBox21.value) + val(UserForm1.TextBox20.value) + val(UserForm1.TextBox25.value) + val(UserForm1.TextBox27.value) + val(UserForm1.TextBox29.value)
'UserForm1.TextBox32.Value = CDbl(UserForm1.TextBox10.Value * 0.1)
'UserForm1.TextBox33.Value = CDbl(UserForm1.TextBox10.Value * 0.15)
'============================?????? ???????



'===============???????????
Dim ast8ta3 As Double

ast8ta3 = val(UserForm1.TextBox24) + val(UserForm1.TextBox32) + val(UserForm1.TextBox33) + val(UserForm1.TextBox34) _
+ val(UserForm1.TextBox35) + val(UserForm1.TextBox38) _
+ val(UserForm1.TextBox42) + val(UserForm1.TextBox46)

Dim al_safy As Double
al_safy = ast78a8 - ast8ta3
Dim Amount_of_absence As Double
Amount_of_absence = al_safy / 30

UserForm1.TextBox47 = Format(Amount_of_absence * val(UserForm1.TextBox11), "00.0")
UserForm1.TextBox12.value = 30 - val(UserForm1.TextBox11)

End Sub


