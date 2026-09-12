VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm84 
   Caption         =   "UserForm84"
   ClientHeight    =   2412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5424
   OleObjectBlob   =   "UserForm84.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm84"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim ws As Worksheet
    Dim rngFound As Range
    Dim searchVal As String
    Dim newValue As Double
    Dim lastR As Long
    
    Set ws = ThisWorkbook.Sheets(1)
    
    ' 1. Ã·» «·ﬁÌ„
    searchVal = UserForm1.TextBox2.Text
    newValue = val(Me.TextBox1.Text)
    
    ' 2. «·»ÕÀ Ê«· ÕœÌÀ
    lastR = ws.Cells(ws.rowS.count, 2).End(xlUp).row
    Set rngFound = ws.Range("B9:B" & lastR).Find(What:=searchVal, LookAt:=xlWhole)
    
    If rngFound Is Nothing Then
        MsgBox "·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·ﬁÌ„… «·„ÿ«»ﬁ… ›Ì «·⁄„Êœ B", vbCritical
        Exit Sub
    End If
    
    ' 3.  ÕœÌÀ «·ﬁÌ„
    ws.Cells(rngFound.row, "CH").value = newValue
    
    If UserForm1.CheckBox6.value = True Then
        ws.Cells(rngFound.row, 24).value = newValue '  ÕœÌÀ X ≈–« ﬂ«‰ «·‰Ÿ«„ ÌœÊÌ«
    End If
    
    ' ---  „ Õ–› ”ÿ— UserForm1.RefreshData „‰ Â‰« ---
    
    MsgBox " „  ÕœÌÀ «·ﬁÌ„… »‰Ã«Õ œÊ‰ «·Õ«Ã… · ÕœÌÀ «·Ê«ÃÂ…", vbInformation
End Sub

Private Sub CommandButton2_Click()
If MsgBox("Â·  Êœ Õ›Ÿ «· €ÌÌ—«  Ê«·Œ—ÊÃø", vbQuestion + vbYesNo, " √ﬂÌœ «·Œ—ÊÃ") = vbYes Then
        ThisWorkbook.Save
        Unload Me
    Else
        ' ≈–« «Œ «— ·«° Ì„ﬂ‰ ≈€·«ﬁ «·›Ê—„ ›ﬁÿ œÊ‰ Õ›Ÿ
        Unload Me
    End If
End Sub
