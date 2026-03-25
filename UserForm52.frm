VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm52 
   Caption         =   "‘—Ìÿ «· ﬁœ„ ·ÕÌ‰  Õ„Ì· «·»Ì«‰« "
   ClientHeight    =   810
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10908
   OleObjectBlob   =   "UserForm52.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm52"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
On Error Resume Next
Dim i As Integer
    Dim TotalSteps As Integer
    TotalSteps = 100 ' ????? ??????? ??? ??? ???????
    
    ' ????? ??????
    Me.Frame2.Width = 0
    
    For i = 1 To TotalSteps
       



UserForm1.ComboBox1.List = Sheets(4).Range("f1:f11").value
UserForm1.ComboBox2.List = Sheets(5).Range("a1:a373").value
UserForm1.ComboBox3.List = Sheets(4).Range("o1:p13").value
UserForm1.ComboBox15.List = Sheets(4).Range("o1:p13").value
UserForm1.ComboBox4.List = Array("–ﬂ—", "«‰ÀÏ")
UserForm1.ComboBox14.List = Array("ŒÿÊ—… ”«∆ﬁ", "·« Ì” Õﬁ")

UserForm1.ComboBox5.List = Sheets(5).Range("a1:a373").value
UserForm1.ComboBox6.List = Array("0.15", "0.25", "0.35", "0.45", "0.55", "0.75", "1.25", "1.50", "0")
UserForm1.ComboBox7.List = Array("0.35", "0.5", "0")
UserForm1.ComboBox8.List = Array("0.5", "0.3", "0.25", "0.20", "0.15", "0")
UserForm1.ComboBox9.List = Array("„«” —", "ﬂ«‘")
UserForm1.ComboBox10.List = Array("0.3", "0.25", "0.20", "0.15", "0")
UserForm1.ComboBox11.List = Array("1", "0.80", "0.30", "0.25", "0.20", "0.15", "0")
UserForm1.ComboBox12.List = Array("ﬂ«‰Ê‰ «·À«‰Ì", "‘»«ÿ", "¬–«—", "‰Ì”«‰", "√Ì«—", "Õ“Ì—«‰", " „Ê“", "¬»", "√Ì·Ê·", " ‘—Ì‰†«·√Ê·", " ‘—Ì‰†«·À«‰Ì", "ﬂ«‰Ê‰†«·√Ê·")
UserForm1.ComboBox13.List = Array("ﬂ«‰Ê‰ «·À«‰Ì", "‘»«ÿ", "¬–«—", "‰Ì”«‰", "√Ì«—", "Õ“Ì—«‰", " „Ê“", "¬»", "√Ì·Ê·", " ‘—Ì‰†«·√Ê·", " ‘—Ì‰†«·À«‰Ì", "ﬂ«‰Ê‰†«·√Ê·")


















Dim last1 As Integer
last1 = Sheets(1).Cells(Rows.count, "e").End(xlUp).row
UserForm1.TextBox75.Text = last1 - 8

        Me.Frame2.Width = (i / TotalSteps) * Me.Frame1.Width
        Me.Label1.Caption = "Ã«—Ì  ÂÌ∆… «·‰Ÿ«„ ·⁄«„ 2026... " & i & "%"
        
        ' ??? ???? ???? ?????? ????? ???????? ????? ???????
        DoEvents
        
        ' ????? ???? ??????? ?????? (???????)
        Application.Wait (Now + TimeValue("00:00:00.03"))
    Next i
    
    ' ????? ???? ?????? ???? ?????? ???????
    Unload Me
    UserForm1.Show ' ?????? ????? ???? ???? ????? ????? ??

End Sub

Private Sub UserForm_Click()

End Sub
