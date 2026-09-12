VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm45 
   Caption         =   "UserForm45"
   ClientHeight    =   7092
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9828.001
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



 Sheets(52).Cells(16, "a").value = Int(TextBox1.Text)
Sheets(52).Cells(16, "c").value = Int(TextBox2.Text)
Sheets(52).Cells(16, "e").value = Int(TextBox3.Text)
Sheets(52).Cells(16, "g").value = Int(TextBox4.Text)
Sheets(52).Cells(16, "i").value = Int(TextBox5.Text)
 Sheets(52).Cells(18, "a").value = Int(TextBox6.Text)
Sheets(52).Cells(18, "c").value = Int(TextBox7.Text)
Sheets(52).Cells(18, "e").value = Int(TextBox8.Text)
Sheets(52).Cells(18, "g").value = Int(TextBox9.Text)
Sheets(52).Cells(18, "i").value = Int(TextBox10.Text)
 Sheets(52).Cells(20, "a").value = Int(TextBox11.Text)
Sheets(52).Cells(20, "c").value = Int(TextBox12.Text)
Sheets(52).Cells(20, "e").value = Int(TextBox13.Text)
Sheets(52).Cells(20, "g").value = Int(TextBox14.Text)
Sheets(52).Cells(20, "i").value = Int(TextBox15.Text)
Sheets(52).Cells(22, "a").value = Int(TextBox16.Text)
Sheets(52).Cells(22, "e").value = Int(TextBox17.Text)
Sheets(52).Cells(20, "f").value = Int(TextBox18.Text)






End Sub

Private Sub Label3_Click()

End Sub

Private Sub UserForm_Activate()
TextBox1.Text = Int(Sheets(1).Cells(3, "cj"))
TextBox2.Text = Int(Sheets(1).Cells(3, "ck"))
TextBox3.Text = Int(Sheets(1).Cells(3, "cl"))
TextBox4.Text = Int(Sheets(1).Cells(3, "cm"))
TextBox5.Text = Int(Sheets(1).Cells(3, "cn"))
TextBox6.Text = Int(Sheets(1).Cells(3, "co"))
TextBox7.Text = Int(Sheets(1).Cells(3, "cp"))
TextBox8.Text = Int(Sheets(1).Cells(3, "cq"))
TextBox9.Text = Int(Sheets(1).Cells(5, "cj"))
TextBox10.Text = Int(Sheets(1).Cells(5, "ck"))
TextBox11.Text = Int(Sheets(1).Cells(5, "cm"))
TextBox12.Text = Int(Sheets(1).Cells(5, "cn"))
TextBox13.Text = Int(Sheets(1).Cells(5, "co"))
TextBox14.Text = Int(Sheets(1).Cells(5, "cp"))
TextBox15.Text = Int(Sheets(1).Cells(5, "cq"))
TextBox16.Text = Int(Sheets(1).Cells(5, "cl"))
TextBox17.Text = Int(Sheets(1).Cells(2, "cg"))
TextBox18.Text = Int(Sheets(1).Cells(3, "cr"))






End Sub

   

Private Sub UserForm_Initialize()

End Sub
