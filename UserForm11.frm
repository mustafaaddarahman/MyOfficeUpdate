VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm11 
   Caption         =   "UserForm11"
   ClientHeight    =   2952
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17280
   OleObjectBlob   =   "UserForm11.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm11"
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
On Error Resume Next
Dim cell As Range
Dim lr As Integer
If Me.ComboBox1.value = "" Then Me.ComboBox1.Clear: Exit Sub
lr = Sheets(13).Cells(Rows.count, 1).End(xlUp).row
For Each cell In Sheets(13).Range("e9:e" & lr)
  If Right(cell.value, 1) = Me.ComboBox1.Text Then
    Me.ComboBox1.AddItem cell.value
     End If
    Next
   '=======================================
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   Sheets(17).Activate
[b1] = Me.ComboBox1.Text


 TextBox1.Text = [e1]
TextBox3.Text = [e2]
 TextBox4.Text = [h2]
 TextBox5.Text = [j2]
 TextBox6.Text = [l2]
 TextBox7.Text = [b3]
 TextBox8.Text = [d3]
 TextBox9.Text = [f3]
TextBox10.Text = [h3]
 TextBox11.Text = [j3]
 TextBox12.Text = [l3]
TextBox19.Text = [b5]
 TextBox41.Text = [d5]
 TextBox21.Text = [f5]
 TextBox22.Text = [h5]
 TextBox23.Text = [j5]
 TextBox24.Text = [l5]
 TextBox25.Text = [b5]
 TextBox26.Text = [d6]
 TextBox27.Text = [f6]
 TextBox28.Text = [h6]
 TextBox29.Text = [j6]
 TextBox30.Text = [l6]
 TextBox31.Text = [b7]
 TextBox32.Text = [d7]
 TextBox33.Text = [f7]
 TextBox13.Text = [b4]
 TextBox14.Text = [d4]
 TextBox15.Text = [f4]
 TextBox16.Text = [h4]
 TextBox17.Text = [j4]
 TextBox18.Text = [l4]
 TextBox42.Text = [g1]
 TextBox34.Text = [H7]
  TextBox43.Text = [j7]
  TextBox44.Text = [h1]
  


    
    
    
    
    
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
Private Sub CommandButton1_Click()
Unload Me
Sheets(17).Range("a1:l7").PrintPreview
UserForm11.Show
End Sub

Private Sub TextBox2_Change()
On Error Resume Next
Sheets(17).Activate
[b1] = TextBox2.Text

If TextBox2.Text = [b1] Then
 TextBox1.Text = [e1]
TextBox3.Text = [e2]
 TextBox4.Text = [h2]
 TextBox5.Text = [j2]
 TextBox6.Text = [l2]
 TextBox7.Text = [b3]
 TextBox8.Text = [d3]
 TextBox9.Text = [f3]
TextBox10.Text = [h3]
 TextBox11.Text = [j3]
 TextBox12.Text = [l3]
TextBox19.Text = [b5]
 TextBox41.Text = [d5]
 TextBox21.Text = [f5]
 TextBox22.Text = [h5]
 TextBox23.Text = [j5]
 TextBox24.Text = [l5]
 TextBox25.Text = [b5]
 TextBox26.Text = [d6]
 TextBox27.Text = [f6]
 TextBox28.Text = [h6]
 TextBox29.Text = [j6]
 TextBox30.Text = [l6]
 TextBox31.Text = [b7]
 TextBox32.Text = [d7]
 TextBox33.Text = [f7]
 TextBox13.Text = [b4]
 TextBox14.Text = [d4]
 TextBox15.Text = [f4]
 TextBox16.Text = [h4]
 TextBox17.Text = [j4]
 TextBox18.Text = [l4]
 TextBox42.Text = [g1]
 TextBox34.Text = [H7]
  TextBox43.Text = [n1]
  TextBox44.Text = [n2]
End If











End Sub

Private Sub TextBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ComboBox1.Visible = True
End Sub



Private Sub TextBox42_Change()

End Sub

Private Sub UserForm_Activate()

Dim lastR As Integer
lastR = Sheets(13).Cells(Rows.count, 2).End(xlUp).row
ComboBox1.List = Sheets(13).Range("e9:e" & lastR).value
Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo emad
Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Sheets(17).Range("a1:l7").PrintPreview
ThisWorkbook.Application.Visible = False
UserForm10.Show
emad:
End Sub
