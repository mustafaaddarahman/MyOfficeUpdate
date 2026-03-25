VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm8 
   Caption         =   "UserForm8"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7332
   OleObjectBlob   =   "UserForm8.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm8"
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
Private Sub CommandButton1_Click()
'On Error Resume Next
Dim new_user, new_passwor, permation As String
Dim rowCount As Integer
Dim lastRow As Integer
rowCount = Sheets(12).Cells(Rows.count, "a").End(xlUp).row + 1

new_user = TextBox1.Text
new_passwor = TextBox2.Text
permation = TextBox3.Text

Sheets(12).Cells(rowCount, 1) = new_user
Sheets(12).Cells(rowCount, 2) = new_passwor
Sheets(12).Cells(rowCount, 3) = permation

End Sub

Private Sub CommandButton3_Click()
Unload Me
UserForm9.Show
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
On Error Resume Next
TextBox1.SetFocus
Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

End Sub

