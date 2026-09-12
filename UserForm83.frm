VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm83 
   Caption         =   "UserForm83"
   ClientHeight    =   12000
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   19752
   OleObjectBlob   =   "UserForm83.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm83"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Initialize()
    '  ÂÌ∆… «·‹ ListView »—„ÃÌ« ·÷„«‰ ÊÃÊœ √⁄„œ… (·„‰⁄ Œÿ√ Invalid Property)
    Dim i As Integer
    Me.ListView1.View = 3 ' ⁄—÷ ﬂÃœÊ·
    Me.ListView1.ColumnHeaders.Clear
    For i = 1 To 10
        Me.ListView1.ColumnHeaders.Add , , "⁄„Êœ " & i
    Next i
    
    Call RefreshData
End Sub

Public Sub RefreshData()
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Integer
    Dim activeCol As Integer
    Dim ctl As Control
    Dim itm As Object
    
    Set ws = ThisWorkbook.Sheets("BlackBox")
    
    '  ÕœÌœ —ﬁ„ «·⁄„Êœ »‰«¡ ⁄·Ï «·‹ CheckBox «·„›⁄·
    activeCol = 1
    For Each ctl In Me.Controls
        If TypeName(ctl) = "CheckBox" Then
            If ctl.value = True Then
                ' «” Œ—«Ã «·—ﬁ„ „‰ «·«”„ (CheckBox1 -> 1)
                activeCol = val(Replace(ctl.Name, "CheckBox", ""))
                Exit For
            End If
        End If
    Next ctl
    
    '  ‰ŸÌ› «·»Ì«‰« 
    Me.ListView1.ListItems.Clear
    
    ' Ã·» «·»Ì«‰« 
    For r = 2 To ws.Cells(ws.rowS.count, 1).End(xlUp).row
        ' «· Õﬁﬁ „‰ «·‰’
        If Me.TextBox1.Text = "" Or InStr(1, CStr(ws.Cells(r, activeCol).value), Me.TextBox1.Text, vbTextCompare) > 0 Then
            
            ' ≈÷«›… «·⁄‰’—
            Set itm = Me.ListView1.ListItems.Add(, , CStr(ws.Cells(r, 1).value))
            
            '  ⁄»∆… «·√⁄„œ… «· ”⁄…
            For i = 1 To 9
                itm.SubItems(i) = CStr(ws.Cells(r, i + 1).value)
            Next i
        End If
    Next r
End Sub

' «·√Õœ«À (·«  ‰”Û —»ÿÂ«)
Private Sub TextBox1_Change(): Call RefreshData: End Sub

Private Sub CheckBox1_Click(): Call RefreshData: End Sub
Private Sub CheckBox2_Click(): Call RefreshData: End Sub
Private Sub CheckBox3_Click(): Call RefreshData: End Sub
Private Sub CheckBox4_Click(): Call RefreshData: End Sub
Private Sub CheckBox5_Click(): Call RefreshData: End Sub
Private Sub CheckBox6_Click(): Call RefreshData: End Sub
Private Sub CheckBox7_Click(): Call RefreshData: End Sub
Private Sub CheckBox8_Click(): Call RefreshData: End Sub
Private Sub CheckBox9_Click(): Call RefreshData: End Sub
Private Sub CheckBox10_Click(): Call RefreshData: End Sub
