VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm74 
   Caption         =   "«· Õﬁﬁ „‰ «·ÂÊÌ…"
   ClientHeight    =   3432
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5784
   OleObjectBlob   =   "UserForm74.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm74"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnEnter_Click()
    If txtPass.Text = "emadsaad" Then
        Unload Me
        On Error Resume Next
        UserForm73.Show
    Else
        MsgBox "ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ…!", vbCritical, " ‰»ÌÂ"
        txtPass.Text = ""
        txtPass.SetFocus
    End If
End Sub
