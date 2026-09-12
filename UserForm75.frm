VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm75 
   Caption         =   "UserForm73"
   ClientHeight    =   7596
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6048
   OleObjectBlob   =   "UserForm75.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm75"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton17_Click()
Unload Me
End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
'On Error Resume Next
TextBox1.Text = Sheets(18).Cells(32, "m")
TextBox2.Text = Sheets(26).Cells(5, "t")

TextBox3.value = Sheets(18).Cells(32, "m") - Sheets(26).Cells(5, "t")
TextBox4.value = Sheets(26).Cells(9, "t")
TextBox5.value = Sheets(26).Cells(35, "p")
TextBox6.value = Sheets(26).Cells(35, "p") - Sheets(1).Cells(2, "cg")
TextBox7.value = Sheets(26).Cells(9, "t") - Sheets(1).Cells(2, "cf")
End Sub
