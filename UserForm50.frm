VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm50 
   Caption         =   "UserForm50"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6516
   OleObjectBlob   =   "UserForm50.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm50"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton1_Click()

Dim LicenseKey As String
    LicenseKey = "PRO-2026-XYZ" ' ÷⁄ ﬂÊœ «· ›⁄Ì· «·Œ«’ »ﬂ Â‰«

    If TextBox1.Text = LicenseKey Then
        ' ≈–« ﬂ«‰ «·ﬂÊœ ’ÕÌÕ«
        MsgBox " „  ›⁄Ì· «·‰Ÿ«„ »‰Ã«Õ ·„œ… ”‰….", vbInformation
        
        '  Œ“Ì‰  «—ÌŒ «· ›⁄Ì· Ê„”Õ √Ì ﬁ›· “„‰Ì
        SaveSetting "MySystem", "Security", "ActivationDate", Date
        SaveSetting "MySystem", "Security", "LockTime", ""
        
        ' «·«‰ ﬁ«· ··ÌÊ“—›Ê—„ 20
        Unload Me
        UserForm20.Show
    Else
        ' ≈–« ﬂ«‰ «·ﬂÊœ Œ«ÿ∆«
        MsgBox "ﬂÊœ «··Ì”‰“ €Ì— ’ÕÌÕ! ”Ì „ €·ﬁ «· ÿ»Ìﬁ Ê·‰ Ì› Õ ≈·« »⁄œ 5 œﬁ«∆ﬁ.", vbCritical
        
        '  ”ÃÌ· Êﬁ  «·Œÿ√ ··ﬁ›·
        SaveSetting "MySystem", "Security", "LockTime", Now
        
        ' Õ›Ÿ Ê≈€·«ﬁ «· ÿ»Ìﬁ ›Ê—«
        ThisWorkbook.Close SaveChanges:=True
    End If
    
     Unload Me ' ≈€·«ﬁ Ê«ÃÂ… «· ›⁄Ì· √Ê·« · ›—Ì€ «·–«ﬂ—…
       UserForm20.Show ' «” œ⁄«¡ «·„«ﬂ—Ê ·› Õ Ê«ÃÂ… «·»—‰«„Ã »”·«”…
End Sub









Private Sub CommandButton2_Click()
Application.Quit
End Sub


Private Sub TextBox1_Change()
On Error Resume Next
Dim LicenseKey As String
    LicenseKey = "PRO-2026-XYZ" ' ÷⁄ ﬂÊœ «· ›⁄Ì· «·Œ«’ »ﬂ Â‰«
 If TextBox1.Text = LicenseKey Then
        SaveSetting "MySystem", "Security", "ActivationDate", Date
        MsgBox " „ «· ›⁄Ì· »‰Ã«Õ!", vbInformation
         CommandButton1.Visible = True
           End If
End Sub



Private Sub UserForm_Activate()
ThisWorkbook.Application.Visible = False
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
Application.Visible = False
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = 0 Then
Cancel = True
End If
End Sub
