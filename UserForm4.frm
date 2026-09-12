VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "«Õ’«∆Ì«  Õ”» «·œ—Ã« "
   ClientHeight    =   13620
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   23184
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm4"
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
Private Sub CommandButton3_Click()
On Error Resume Next
Dim vlookup1 As Long
Dim vlookup2 As Long
Dim vlookup3 As Long

Dim textvlue1 As Long
Dim textvlue2 As Long
Dim combvlue As Long
Dim rang1 As Range
Dim rang2 As Range

'=IF(AND(VLOOKUP(W9,“ÊÃÌ…4,2,0)<4,V9>0),"Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—",VLOOKUP(J9,«” ﬁÿ«⁄« ,VLOOKUP(W9,“ÊÃÌ…4,2,0)+V9,1))

Set rang1 = Sheets(1).Range("“ÊÃÌ…4")
Set rang2 = Sheets(2).Range("«” ﬁÿ«⁄« ")
textvlue1 = TextBox2.Text
textvlue2 = TextBox1.Text

vlookup1 = Application.WorksheetFunction.VLookup(ComboBox2.value, rang1, 2, 0) < 4 And TextBox2.value > 0
'TextBox27.Value = vlookup1

'TextBox26.Value = Application.WorksheetFunction.VLookup(textvlue2, rang2, Application.WorksheetFunction.VLookup(ComboBox2.Value, rang1, 2, 0) + textvlue1, 1)


If vlookup1 = -1 Then
TextBox2.value = 0
MsgBox "Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—", vbOKCancel, "—”«·…  ÊÃÌÂ"
Exit Sub
Else
TextBox26.value = Application.WorksheetFunction.VLookup(textvlue2, rang2, Application.WorksheetFunction.VLookup(ComboBox2.value, rang1, 2, 0) + textvlue1, 1)


End If

End Sub

Private Sub UserForm_Activate()
Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
ComboBox2.AddItem
ComboBox2.List = Sheets(1).Range("bg1:bg5").value

End Sub
