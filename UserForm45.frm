VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm45 
   Caption         =   "UserForm45"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10692
   OleObjectBlob   =   "UserForm45.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm45"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton27_Click()
Unload Me
UserForm10.Show
End Sub

Private Sub CommandButton28_Click()



 Sheets(52).Cells(16, "a").value = TextBox1.Text
Sheets(52).Cells(16, "c").value = TextBox2.Text
Sheets(52).Cells(16, "e").value = TextBox3.Text
Sheets(52).Cells(16, "g").value = TextBox4.Text
Sheets(52).Cells(16, "i").value = TextBox5.Text
 Sheets(52).Cells(18, "a").value = TextBox6.Text
Sheets(52).Cells(18, "c").value = TextBox7.Text
Sheets(52).Cells(18, "e").value = TextBox8.Text
Sheets(52).Cells(18, "g").value = TextBox9.Text
Sheets(52).Cells(18, "i").value = TextBox10.Text
 Sheets(52).Cells(20, "a").value = TextBox11.Text
Sheets(52).Cells(20, "c").value = TextBox12.Text
Sheets(52).Cells(20, "e").value = TextBox13.Text
Sheets(52).Cells(20, "g").value = TextBox14.Text
Sheets(52).Cells(20, "i").value = TextBox15.Text
Sheets(52).Cells(22, "a").value = TextBox16.Text
Sheets(52).Cells(22, "e").value = TextBox17.Text







End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Activate()
TextBox1.Text = Sheets(1).Cells(3, "cj")
TextBox2.Text = Sheets(1).Cells(3, "ck")
TextBox3.Text = Sheets(1).Cells(3, "cl")
TextBox4.Text = Sheets(1).Cells(3, "cm")
TextBox5.Text = Sheets(1).Cells(3, "cn")
TextBox6.Text = Sheets(1).Cells(3, "co")
TextBox7.Text = Sheets(1).Cells(3, "cp")
TextBox8.Text = Sheets(1).Cells(3, "cq")
TextBox9.Text = Sheets(1).Cells(5, "cj")
TextBox10.Text = Sheets(1).Cells(5, "ck")
TextBox11.Text = Sheets(1).Cells(5, "cm")
TextBox12.Text = Sheets(1).Cells(5, "cn")
TextBox13.Text = Sheets(1).Cells(5, "co")
TextBox14.Text = Sheets(1).Cells(5, "cp")
TextBox15.Text = Sheets(1).Cells(5, "cq")
TextBox16.Text = Sheets(1).Cells(5, "cl")
TextBox17.Text = Sheets(1).Cells(2, "cg")







End Sub

