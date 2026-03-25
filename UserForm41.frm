VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm41 
   Caption         =   "UserForm41"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8148
   OleObjectBlob   =   "UserForm41.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label32_Click()

End Sub

Private Sub UserForm_Activate()
Call englsh
 TextBox1.Text = UserForm1.TextBox10.Text
TextBox2.Text = UserForm1.TextBox30.Text
TextBox3.Text = UserForm1.TextBox31.Text
TextBox4.Text = UserForm1.TextBox49.Text
Label36.Caption = UserForm1.TextBox5.Text
End Sub

Private Sub UserForm_Click()

End Sub
