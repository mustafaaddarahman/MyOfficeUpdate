VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm42 
   Caption         =   "UserForm42"
   ClientHeight    =   2190
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7692
   OleObjectBlob   =   "UserForm42.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm42"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Call whatsapp_trans_one_tape_1
End Sub



Private Sub TextBox1_Change()
Call englsh
TextBox3.value = ""
End Sub



Private Sub TextBox2_Change()

On Error Resume Next
TextBox3.value = ""

Call arb
Dim last2, frw2 As Integer
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "cc").value = TextBox1.Text And Sheets(1).Cells(frw2, "e").value Like "*" & TextBox2.Text & "*" Then
'TextBox1.Text = "+964" & TextBox4.Text
TextBox3.value = Sheets(1).Cells(frw2, "a").value + 8

End If
Next
If TextBox2.value = "" Then
TextBox3.value = ""
End If
End Sub

Private Sub TextBox4_Change()

End Sub



Private Sub TextBox2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call arb
End Sub

Private Sub UserForm_Click()

End Sub
