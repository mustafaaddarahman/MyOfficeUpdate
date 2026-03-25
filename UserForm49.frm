VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm49 
   Caption         =   "UserForm49"
   ClientHeight    =   4200
   ClientLeft      =   10128
   ClientTop       =   9468.001
   ClientWidth     =   10512
   OleObjectBlob   =   "UserForm49.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm49"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Label37_Click()
UserForm48.TextBox72.SetFocus
Label37.Caption = "«Ìﬁ«› «·«Ã—«¡"
Unload Me
End Sub

Private Sub Label38_Click()
On Error Resume Next
    Dim i As Long
    
    ' 1.  Õ—Ìﬂ «· ÕœÌœ ›Ì ListBox1 «·„ÊÃÊœ ›Ì UserForm48
    With UserForm48.ListBox1
        If .ListCount = 0 Then Exit Sub
        
        i = .ListIndex
        If i < .ListCount - 1 Then
            ' «·«‰ ﬁ«· ··⁄‰’— «· «·Ì (”ÌﬁÊ„ » ‘€Ì· ﬂÊœ «·‹ Click ›Ì 48  ·ﬁ«∆Ì«)
            .ListIndex = i + 1
            
            ' ÷„«‰  Ÿ·Ì· «·”ÿ— »«··Ê‰ «·√“—ﬁ »—„ÃÌ«
            .Selected(i + 1) = True
            
            ' 2.  ÕœÌÀ »Ì«‰«  UserForm49 «·Õ«·Ì… „»«‘—… „‰ ›Ê—„ 48 («·”—⁄… «·ﬁ’ÊÏ)
            ' »œ·« „‰ Unload Ê Show° ‰ﬁÊ„ » ÕœÌÀ «·ﬁÌ„ ›ﬁÿ · Ã‰» «· √ŒÌ—
           
            ' ‰ﬁ· «·’Ê—… »—„ÃÌ« „‰ «·–«ﬂ—… (√”—⁄ 100 „—… „‰  Õ„Ì·Â« „‰ «·ﬁ—’)
            
            
           '=====================================
Label36.Caption = UserForm48.TextBox5.Text
TextBox1.Text = UserForm48.TextBox10.Text
TextBox2.Text = UserForm48.TextBox30.Text
TextBox3.Text = UserForm48.TextBox31.Text
TextBox4.Text = UserForm48.TextBox49.Text
TextBox5.Text = UserForm48.TextBox24.Text
TextBox6.Text = UserForm48.TextBox48.Text
Set Me.Image1.Picture = UserForm48.Image1.Picture
            
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "e") = Label36.Caption Then
Image1.Picture = LoadPicture(Sheets(1).Cells(frw1, 64))

End If
Next frw1
            
            
            
            
            
            '====================================
            
            ' 3. ≈Ã»«— «·›Ê—„ ⁄·Ï ≈⁄«œ… «·—”„ ·≈ŸÂ«— «·»Ì«‰«  «·ÃœÌœ… ›Ê—«
            Me.Repaint
        Else
            MsgBox "Ê’·  ·‰Â«Ì… «·ﬁ«∆„…", vbInformation, " ‰»ÌÂ"
        End If
    End With
End Sub

Private Sub Label39_Click()

End Sub

Private Sub Label41_Click()
On Error Resume Next
    Dim i As Long
    
    ' 1. «·≈‘«—… ··ﬁ«∆„… ›Ì ›Ê—„ 48
    With UserForm48.ListBox1
        ' «· √ﬂœ „‰ ÊÃÊœ ”Ã·«  Ê√‰‰« ·”‰« ›Ì √Ê· ”ÿ—
        If .ListCount = 0 Then Exit Sub
        
        i = .ListIndex
        If i > 0 Then
            ' «·—ÃÊ⁄ ··”ÿ— «·”«»ﬁ (Â–« Ì‘€· ﬂÊœ «· ⁄»∆… ›Ì 48  ·ﬁ«∆Ì«)
            .ListIndex = i - 1
            
            '  Ÿ·Ì· «·”ÿ— «·”«»ﬁ »«··Ê‰ «·√“—ﬁ »—„ÃÌ« ( ÕœÌœ  ·ﬁ«∆Ì)
            .Selected(i - 1) = True
            
            ' 2.  ÕœÌÀ »Ì«‰«  ›Ê—„ 49 „‰ ›Ê—„ 48 ›Ê—« (œ›⁄ «·»Ì«‰« )
            ' »œ·« „‰ ≈€·«ﬁ Ê› Õ «·›Ê—„° ‰ÕœÀ «·ﬁÌ„ „»«‘—… ··”—⁄…
           Label36.Caption = UserForm48.TextBox5.Text
TextBox1.Text = UserForm48.TextBox10.Text
TextBox2.Text = UserForm48.TextBox30.Text
TextBox3.Text = UserForm48.TextBox31.Text
TextBox4.Text = UserForm48.TextBox49.Text
TextBox5.Text = UserForm48.TextBox24.Text
TextBox6.Text = UserForm48.TextBox48.Text
Set Me.Image1.Picture = UserForm48.Image1.Picture
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "e") = Label36.Caption Then
Image1.Picture = LoadPicture(Sheets(1).Cells(frw1, 64))

End If
Next frw1

            Me.Repaint
        Else
            MsgBox "√‰  »«·›⁄· ⁄‰œ √Ê· ”Ã· ›Ì «·ﬁ«∆„…", vbInformation, " ‰»ÌÂ"
        End If
    End With
End Sub

Private Sub UserForm_Activate()
Call englsh
Call auto_acount_1

 TextBox1.Text = UserForm48.TextBox10.Text
TextBox2.Text = UserForm48.TextBox30.Text
TextBox3.Text = UserForm48.TextBox31.Text
TextBox4.Text = UserForm48.TextBox49.Text
TextBox5.Text = UserForm48.TextBox24.Text
TextBox6.Text = UserForm48.TextBox48.Text





Label36.Caption = UserForm48.TextBox5.Text
End Sub

Private Sub UserForm_Click()

End Sub
