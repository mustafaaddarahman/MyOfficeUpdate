VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm43 
   Caption         =   "UserForm43"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6816
   OleObjectBlob   =   "UserForm43.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm43"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim lastR, Y As Integer
lastR = Sheets(1).Cells(Rows.count, "b").End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, "e") = TextBox4.Text And Sheets(1).Cells(Y, "b") = TextBox1.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(1).Cells(Y, "cc").value = TextBox3.value
    MsgBox "·ﬁœ  „  «÷«›… —ﬁ„ «·„Ê»«Ì· »‰Ã«Õ", vbInformation + vbDefaultButton1, "—”«·…  «ﬂÌœ"
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox2_Change()
Call arb
Dim last2, frw2 As Integer


last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "e").value Like "*" & TextBox2.Text & "*" Then
TextBox1.value = Sheets(1).Cells(frw2, "b")
TextBox4.value = Sheets(1).Cells(frw2, "e")
End If
Next
End Sub

Private Sub TextBox3_Change()
Call englsh
End Sub

Private Sub TextBox4_Change()

End Sub

Private Sub UserForm_Click()

End Sub
