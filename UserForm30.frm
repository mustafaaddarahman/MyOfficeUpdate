VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm30 
   Caption         =   "UserForm30"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   8340.001
   OleObjectBlob   =   "UserForm30.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ButtonBar1_OnClick(ByVal ButtonId As Long)

End Sub

Private Sub CommandButton1_Click()
Frame1.Visible = True
End Sub

Private Sub ListBox1_Click()
If ListBox1.Selected(ListBox1.ListIndex) Then
Select Case ListBox1.ListIndex

Case 1

UserForm1.Show
Case 2

ListBox1.ListIndex = 2
UserForm2.Show
Case 3
ListBox1.ListIndex = 3
UserForm3.Show
End Select
End If
End Sub

Private Sub UserForm_Click()
Frame1.Visible = False
End Sub
