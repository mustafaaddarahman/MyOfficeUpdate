VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "UserForm7"
   ClientHeight    =   3285
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5676
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME As Long = &H40000
Const MIN_BOX As Long = &H20000
Const MAX_BOX As Long = &H10000
Private Declare PtrSafe Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As Long

Public Sub AddToForm(ByVal Box_Type As Long)
Dim BisMask As Long
Dim Window_Handle As Long
Dim WindowStyle As Long
Dim Ret As Long
If Box_Type = MIN_BOX Or Box_Type = MAX_BOX Then
    Window_Handle = GetForegroundWindow()
    WindowStyle = GetWindowLong(Window_Handle, GWL_STYLE)
    BisMask = WindowStyle Or Box_Type
    Ret = SetWindowLong(Window_Handle, GWL_STYLE, BisMask)
    Ret = DrawMenuBar(Window_Handle)
End If
End Sub
Private Sub ComboBox1_Change()

Dim luser As Integer
On Error Resume Next

luser = Sheets(12).Range("e1").value
Sheets(12).Range("d1").value = ComboBox1.value
CommandButton3.Enabled = False
If Not Intersect(Target, TextBox2.value) Is Nothing Then
Call englsh

End If
End Sub

Private Sub CommandButton1_Click()
On Error GoTo emad

Sheets(53).Activate
Call copy_from_to
Dim luser As Integer
luser = Sheets(12).Range("e1").value
Sheets(12).Range("d1").value = ComboBox1.value
Dim user_na, pasword As Variant


user_na = ComboBox1.Text
pasword = TextBox2.Text


If user_na = Sheets(12).Cells(luser, "a").Text And pasword = Sheets(12).Cells(luser, "b").Text And Sheets(12).Cells(luser, "c").Text = "admin" Then
'Application.Visible = False
Unload Me


UserForm9.Show


ElseIf user_na = Sheets(12).Cells(luser, "a").Text And pasword = Sheets(12).Cells(luser, "b").Text And Sheets(12).Cells(luser, "c").Text = "user" Then
'Application.Visible = False
Unload Me


UserForm9.Show
Else
MsgBox "·ﬁœ ﬁ„  »«œŒ«· ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕÂ Õ«Ê· ﬂ «» Â« «œŒ«·Â« „—… «Œ—Ï", vbCritical + vb, "—”«·…  ÊÃÌÂ"

TextBox2.Text = ""
 Exit Sub
 End If
 
emad:
End Sub

Private Sub CommandButton2_Click()

Application.Quit

End Sub



Private Sub CommandButton3_Click()
On Error GoTo emad
UserForm8.Show
emad:
End Sub

Private Sub CommandButton4_Click()

End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub TextBox2_AfterUpdate()
On Error GoTo emad


Dim luser As Integer
luser = Sheets(12).Range("e1").value
Sheets(12).Range("d1").value = ComboBox1.value
Dim user_na, pasword As Variant


user_na = ComboBox1.Text
pasword = TextBox2.Text

If user_na = Sheets(12).Cells(luser, "a").Text And pasword = Sheets(12).Cells(luser, "b").Text And Sheets(12).Cells(luser, "c").Text = "admin" Then
'Application.Visible = False

CommandButton3.Enabled = True
Else
CommandButton3.Enabled = False
Exit Sub
End If
emad:
End Sub

Private Sub TextBox2_Change()
On Error Resume Next
If Not Intersect(Target, TextBox2.value) Is Nothing Then
Call englsh

End If
Dim luser As Integer
luser = Sheets(12).Range("e1").value
Sheets(12).Range("d1").value = ComboBox1.value
Dim user_na, pasword As Variant


user_na = ComboBox1.Text
pasword = TextBox2.Text


If user_na = Sheets(12).Cells(luser, "a").Text And pasword = Sheets(12).Cells(luser, "b").Text And Sheets(12).Cells(luser, "c").Text = "admin" Then
'Application.Visible = False


CommandButton3.Enabled = True
Else
CommandButton3.Enabled = False
Exit Sub
End If

End Sub
Private Sub OptionButton1_Click()
If OptionButton1.value = True Then
OptionButton2.value = False
Call arb
End If
End Sub

Private Sub OptionButton2_Click()
On Error Resume Next
If OptionButton2.value = True Then
OptionButton1.value = False
If Not Intersect(Target, UserForm1.ActiveControl) Is Nothing Then
Call englsh
End If
End If
End Sub



Private Sub UserForm_Activate()
On Error GoTo emad
ComboBox1.SetFocus
ComboBox1.DropDown
Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

Dim lrow As Integer

lrow = Sheets(12).Range("a" & Rows.count).End(xlUp).row

ComboBox1.List = Sheets(12).Range("a2:a" & lrow).value

CommandButton3.Enabled = False
emad:
End Sub



Private Sub UserForm_Initialize()

TextBox2.Text = ""

End Sub
