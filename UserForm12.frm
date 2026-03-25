VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm12 
   Caption         =   "UserForm12"
   ClientHeight    =   11685
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7512
   OleObjectBlob   =   "UserForm12.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm12"
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

Private Sub CommandButton15_Click()
 Dim iRow, iRow1 As Long, i As Long

    Sheets(6).Activate

    iRow = Range("A" & Rows.count).End(xlUp).row
    
'=======================================================
  

'=========================================================
    Range("A" & iRow + 1).Offset(0, 0).value = iRow
    Range("A" & iRow + 1).Offset(0, 1).value = TextBox2.value
    Range("A" & iRow + 1).Offset(0, 2).value = TextBox3.value
    Range("A" & iRow + 1).Offset(0, 3).value = TextBox4.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox5.value
    Range("A" & iRow + 1).Offset(0, 5).value = TextBox6.value
    Range("A" & iRow + 1).Offset(0, 6).value = TextBox7.value
    Range("A" & iRow + 1).Offset(0, 7).value = TextBox8.value
    Range("A" & iRow + 1).Offset(0, 8).value = TextBox9.value
    Range("A" & iRow + 1).Offset(0, 9).value = TextBox10.value
    Range("A" & iRow + 1).Offset(0, 10).value = Format(TextBox54.value, "yyyy/mm/dd")
    Range("A" & iRow + 1).Offset(0, 11).value = Format(TextBox59.value, "yyyy/mm/dd")
    Range("A" & iRow + 1).Offset(0, 13).value = Format(TextBox55.value, "yyyy/mm/dd")
    
    

    
    'Range("A" & iRow + 1).Offset(0, 13).Value = Format(TextBox55.Value, "mm/dd/yyyy")
    Range("A" & iRow + 1).Offset(0, 14).value = ""
    Range("A" & iRow + 1).Offset(0, 15).value = ""
   
   
   ' Range("A" & iRow + 1).Offset(0, 16).Value = TextBox17.Value
   ' Range("A" & iRow + 1).Offset(0, 17).Value = TextBox18.Value

        Sheets(6).Activate
'============================================================================


Sheets(1).Activate
For X = 2 To 1500
If Cells(X, 1) = TextBox1.Text And Cells(X, 2) = TextBox2.Text Then


    Cells(X, 1).Select
Exit For
End If
Next X

Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text And Sheets(1).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
 Sheets(1).Cells(Y, "bg").value = TextBox54.Text
 Sheets(1).Cells(Y, "bc").value = TextBox55.Text
 Sheets(1).Cells(Y, "bb").value = TextBox59.Text





  
End Sub

Private Sub CommandButton16_Click()
On Error GoTo emad
Dim i As Integer
 For i = 1 To 60

Controls("textbox" & i).Text = ""
TextBox60.Text = ""
TextBox1.Text = ""
Next
emad:
End Sub

Private Sub CommandButton17_Click()
Unload Me

End Sub

Private Sub TextBox60_Change()
'On Error GoTo emad









Dim r As Integer
Dim last_row As Integer

last_row = Sheets(1).Range("a10000").End(xlUp).row
For r = 2 To last_row
a = Len(TextBox60.Text)
If UCase(Left(Sheets(1).Cells(r, 5).value, a)) = UCase(TextBox60.Text) Then
 TextBox1.Text = Sheets(1).Cells(r, "a").value
TextBox2.Text = Sheets(1).Cells(r, "b").value
  TextBox3.Text = Sheets(1).Cells(r, "c").value
 TextBox4.Text = Sheets(1).Cells(r, "d").value
  TextBox5.Text = Sheets(1).Cells(r, "e").value
 TextBox6.Text = Sheets(1).Cells(r, "f").value
  TextBox7.Text = Sheets(1).Cells(r, "g").value
 TextBox8.Text = Sheets(1).Cells(r, "h").value
  TextBox9.Text = Sheets(1).Cells(r, "i").value
 TextBox10.Text = Sheets(1).Cells(r, "j").value
  TextBox54.Text = Sheets(1).Cells(r, "bg").value
 TextBox55.Text = Sheets(1).Cells(r, "bc").value
 TextBox59.Text = Sheets(1).Cells(r, "bb").value
 Sheets(1).Cells(r, "bg").value = TextBox54.Text
 Sheets(1).Cells(r, "bc").value = TextBox55.Text
 Sheets(1).Cells(r, "bb").value = TextBox59.Text
 
 



End If
Next

 
'emad:
End Sub

Private Sub TextBox60_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo emad
Dim i As Integer
 For i = 1 To 60

Controls("textbox" & i).Text = ""
TextBox60.Text = ""
TextBox1.Text = ""
Next
emad:
End Sub

Private Sub TextBox61_AfterUpdate()
If TextBox61.Text = "#N/A" Then

 Beep
 Beep
Application.Speech.Speak " al    rakum   gair     mau juod"
End If
End Sub

Private Sub TextBox61_Change()
On Error GoTo emad
Sheets(1).Cells(3, 64) = TextBox61.Text
TextBox60.Text = Sheets(1).Cells(3, 65).Text
If TextBox5.Text = Sheets(1).Cells(3, 65).Text Then
 Beep
Application.Speech.Speak " al    rakum       mau juod"
'If TextBox60.Text = "NA" Then
'Application.Speech.Speak " al    rakum   gaer    mau juod"
End If
emad:
End Sub

Private Sub TextBox61_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Sheets(1).Cells(3, 64).value = ""
TextBox61.value = ""
 TextBox1.Text = ""
TextBox2.Text = ""
  TextBox3.Text = ""
 TextBox4.Text = ""
  TextBox5.Text = ""
 TextBox6.Text = ""
  TextBox7.Text = ""
 TextBox8.Text = ""
  TextBox9.Text = ""
 TextBox10.Text = ""
  TextBox54.Text = ""
 TextBox55.Text = ""
 TextBox59.Text = ""
 
 
End Sub

Private Sub UserForm_Activate()

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

End Sub

