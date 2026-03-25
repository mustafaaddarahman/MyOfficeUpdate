VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm44 
   Caption         =   "UserForm44"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   28140
   OleObjectBlob   =   "UserForm44.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm44"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
If CheckBox1.value = True Then
Sheets(49).Cells(1, "a") = "true"
Else
Sheets(49).Cells(1, "a") = "false"
End If
End Sub

Private Sub CheckBox10_Click()
If CheckBox10.value = True Then
Sheets(49).Cells(1, "j") = "true"
Else
Sheets(49).Cells(1, "j") = "false"
End If
End Sub

Private Sub CheckBox11_Click()
If CheckBox11.value = True Then
Sheets(49).Cells(1, "k") = "true"
Else
Sheets(49).Cells(1, "k") = "false"
End If
End Sub

Private Sub CheckBox12_Click()
If CheckBox12.value = True Then
Sheets(49).Cells(1, "l") = "true"
Else
Sheets(49).Cells(1, "l") = "false"
End If
End Sub

Private Sub CheckBox13_Click()
If CheckBox13.value = True Then
Sheets(49).Cells(1, "m") = "true"
Else
Sheets(49).Cells(1, "m") = "false"
End If
End Sub

Private Sub CheckBox14_Click()
If CheckBox14.value = True Then
Sheets(49).Cells(1, "n") = "true"
Else
Sheets(49).Cells(1, "n") = "false"
End If
End Sub

Private Sub CheckBox15_Click()
If CheckBox15.value = True Then
Sheets(49).Cells(1, "o") = "true"
Else
Sheets(49).Cells(1, "o") = "false"
End If
End Sub

Private Sub CheckBox16_Click()
If CheckBox16.value = True Then
Sheets(49).Cells(1, "p") = "true"
Else
Sheets(49).Cells(1, "p") = "false"
End If
End Sub

Private Sub CheckBox17_Click()
If CheckBox17.value = True Then
Sheets(49).Cells(1, "q") = "true"
Else
Sheets(49).Cells(1, "q") = "false"
End If
End Sub

Private Sub CheckBox18_Click()
If CheckBox18.value = True Then
Sheets(49).Cells(1, "r") = "true"
Else
Sheets(49).Cells(1, "r") = "false"
End If
End Sub

Private Sub CheckBox19_Click()
If CheckBox19.value = True Then
Sheets(49).Cells(1, "s") = "true"
Else
Sheets(49).Cells(1, "s") = "false"
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.value = True Then
Sheets(49).Cells(1, "b") = "true"
Else
Sheets(49).Cells(1, "b") = "false"
End If
End Sub





Private Sub CheckBox20_Click()
If CheckBox20.value = True Then
Sheets(49).Cells(1, "t") = "true"
Else
Sheets(49).Cells(1, "t") = "false"
End If
End Sub

Private Sub CheckBox21_Click()
If CheckBox21.value = True Then
Sheets(49).Cells(1, "u") = "true"
Else
Sheets(49).Cells(1, "u") = "false"
End If
End Sub

Private Sub CheckBox22_Click()
If CheckBox22.value = True Then
Sheets(49).Cells(1, "v") = "true"
Else
Sheets(49).Cells(1, "v") = "false"
End If
End Sub

Private Sub CheckBox23_Click()
If CheckBox23.value = True Then
Sheets(49).Cells(1, "w") = "true"
Else
Sheets(49).Cells(1, "w") = "false"
End If
End Sub

Private Sub CheckBox24_Click()
If CheckBox24.value = True Then
Sheets(49).Cells(1, "x") = "true"
Else
Sheets(49).Cells(1, "x") = "false"
End If
End Sub

Private Sub CheckBox25_Click()
If CheckBox25.value = True Then
Sheets(49).Cells(1, "y") = "true"
Else
Sheets(49).Cells(1, "y") = "false"
End If
End Sub

Private Sub CheckBox26_Click()
If CheckBox26.value = True Then
Sheets(49).Cells(1, "z") = "true"
Else
Sheets(49).Cells(1, "z") = "false"
End If
End Sub

Private Sub CheckBox27_Click()
If CheckBox27.value = True Then
Sheets(49).Cells(1, "aa") = "true"
Else
Sheets(49).Cells(1, "aa") = "false"
End If
End Sub

Private Sub CheckBox28_Click()
If CheckBox28.value = True Then
Sheets(49).Cells(1, "ab") = "true"
Else
Sheets(49).Cells(1, "ab") = "false"
End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.value = True Then
Sheets(49).Cells(1, "c") = "true"
Else
Sheets(49).Cells(1, "c") = "false"
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.value = True Then
Sheets(49).Cells(1, "d") = "true"
Else
Sheets(49).Cells(1, "d") = "false"
End If
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.value = True Then
Sheets(49).Cells(1, "e") = "true"
Else
Sheets(49).Cells(1, "e") = "false"
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.value = True Then
Sheets(49).Cells(1, "f") = "true"
Else
Sheets(49).Cells(1, "f") = "false"
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox7.value = True Then
Sheets(49).Cells(1, "g") = "true"
Else
Sheets(49).Cells(1, "g") = "false"
End If
End Sub

Private Sub CheckBox8_Click()
If CheckBox8.value = True Then
Sheets(49).Cells(1, "h") = "true"
Else
Sheets(49).Cells(1, "h") = "false"
End If
End Sub

Private Sub CheckBox9_Click()
If CheckBox9.value = True Then
Sheets(49).Cells(1, "i") = "true"
Else
Sheets(49).Cells(1, "i") = "false"
End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Activate()
CheckBox1.value = Sheets(49).Cells(1, "a")
CheckBox2.value = Sheets(49).Cells(1, "b")
CheckBox3.value = Sheets(49).Cells(1, "c")
CheckBox4.value = Sheets(49).Cells(1, "d")
CheckBox5.value = Sheets(49).Cells(1, "e")
CheckBox6.value = Sheets(49).Cells(1, "f")
CheckBox7.value = Sheets(49).Cells(1, "g")
CheckBox8.value = Sheets(49).Cells(1, "h")
CheckBox9.value = Sheets(49).Cells(1, "i")
CheckBox10.value = Sheets(49).Cells(1, "g")
CheckBox11.value = Sheets(49).Cells(1, "k")
CheckBox12.value = Sheets(49).Cells(1, "l")
CheckBox13.value = Sheets(49).Cells(1, "m")


CheckBox14.value = Sheets(49).Cells(1, "n")
CheckBox15.value = Sheets(49).Cells(1, "o")
CheckBox16.value = Sheets(49).Cells(1, "p")
CheckBox17.value = Sheets(49).Cells(1, "q")
CheckBox18.value = Sheets(49).Cells(1, "r")
CheckBox19.value = Sheets(49).Cells(1, "s")
CheckBox20.value = Sheets(49).Cells(1, "t")
CheckBox21.value = Sheets(49).Cells(1, "u")
CheckBox22.value = Sheets(49).Cells(1, "v")
CheckBox23.value = Sheets(49).Cells(1, "w")
CheckBox24.value = Sheets(49).Cells(1, "x")
CheckBox25.value = Sheets(49).Cells(1, "y")
CheckBox26.value = Sheets(49).Cells(1, "z")

CheckBox27.value = Sheets(49).Cells(1, "aa")
CheckBox28.value = Sheets(49).Cells(1, "ab")








End Sub

Private Sub UserForm_Click()

End Sub
