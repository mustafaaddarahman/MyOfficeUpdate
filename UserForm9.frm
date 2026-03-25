VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm9 
   Caption         =   "UserForm9"
   ClientHeight    =   13668
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   28560
   OleObjectBlob   =   "UserForm9.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If




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

Private Sub CheckBox1_Click()
If CheckBox1.value = True Then
CommandButton17.Enabled = True
Else
CommandButton17.Enabled = False
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.value = True Then
CommandButton20.Enabled = True
Else
CommandButton20.Enabled = False
End If
End Sub

Private Sub CommandButton1_Click()
Unload Me

UserForm52.Show

End Sub

Private Sub CommandButton10_Click()
 Dim timeStamp As String, fileName As String
    Dim fso As Object
    
    ' 1. ”—⁄… ﬁ’ÊÏ: ≈Ìﬁ«› ﬂ· „« Ì” Â·ﬂ «·„⁄«·Ã
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .DisplayAlerts = False
    End With

    ' 2. Õ›Ÿ «·„·› «·Õ«·Ì («·√’·Ì) »”—⁄…
    ThisWorkbook.Save
    
    ' 3.  ÃÂÌ“ «·„”«—« 
    timeStamp = Format(Now, "ddd HH mm")
    fileName = ThisWorkbook.Name
    
    ' 4. «” Œœ«„  ﬁ‰Ì… FSO (√”—⁄ »ﬂÀÌ— „‰ SaveCopyAs ·√‰Â«   ⁄«„· „⁄ «·„·› ﬂﬂ ·… ’„«¡)
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' «·‰”Œ Ì „ ›Ì –«ﬂ—… «·ÊÌ‰œÊ“ „»«‘—… (Œ·›Ì… «·‰Ÿ«„)
    fso.CopyFile ThisWorkbook.fullName, "D:\emad\" & timeStamp & " - " & fileName, True
    fso.CopyFile ThisWorkbook.fullName, "E:\emad\" & timeStamp & " - " & fileName, True
    On Error GoTo 0

    ' 5. ≈€·«ﬁ ›Ê—Ì œÊ‰ √Ì  √ŒÌ—
    If Workbooks.count <= 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

Private Sub CommandButton11_Click()
Dim PASS As Variant
Dim user_pass As Variant
PASS = "0001000"
 user_pass = InputBox("ﬁ„ »«œŒ«· ﬂ·„… «·”— ·Ì ”‰Ï ·ﬂ «·œŒÊ· «·Ï »Ì∆… «· ÿ»Ìﬁ", "—”«·…  ÊÃÌÂ", "")
If user_pass = PASS Then
Call Showit
Application.Visible = True
Unload Me
Else
Exit Sub
End If
End Sub

Private Sub CommandButton12_Click()
Unload Me
UserForm13.Show

End Sub



Private Sub CommandButton13_Click()
UserForm18.Show
End Sub

Private Sub CommandButton14_Click()
UserForm42.Show
End Sub

Private Sub CommandButton15_Click()
If MsgBox("Â· «‰  „ √ﬂœ „‰ «—”«· Ã„Ì⁄ «‘—ÿ… «·—« » «·Ï «·„‰ ”»Ì‰ø" & vbCrLf & _
          "Â·  ÿ»Ìﬁ «·Ê« ”«» Ã«Â“ ··⁄„·ø" & vbCrLf & _
          "Â· Ã„Ì⁄ «—ﬁ«„ «·„‰ ”»Ì‰ „ÊÃÊœ… ›Ì ﬁ«⁄œ… «·»Ì«‰« ø" & vbCrLf & vbCrLf & _
          "«÷€ÿ OK ··„ «»⁄… «Ê Cancel ·«·€«¡ «·⁄„·Ì….", _
          vbOKCancel + vbExclamation, "—”«·…  ‰»ÌÂ") = vbOK Then

    On Error GoTo ErrHandler
    Call whats_app_trans
    Exit Sub

Else
    Exit Sub
End If

ErrHandler:
MsgBox "ÕœÀ Œÿ√ «À‰«¡ ⁄„·Ì… «—”«· «·—Ê« ». Ì—ÃÏ «· √ﬂœ „‰ «·Ê« ”«» Ê„‰ ’Õ… «·«—ﬁ«„.", vbCritical, "Œÿ√"







End Sub

Private Sub CommandButton16_Click()
UserForm43.Show
End Sub

Private Sub CommandButton17_Click()
Dim PASS As Variant
Dim user_pass As Variant
PASS = "1110111"
 user_pass = InputBox("ﬁ„ »«œŒ«· ﬂ·„… «·„—Ê— ·Ì ”‰Ï ·ﬂ  ’›Ì— ﬁ«⁄œ… «·»Ì«‰« ", "ﬂ·„… «·„—Ê—")
If user_pass = PASS Then
Call emad1
Sheets(3).Range("a2:bw10000").ClearContents
Sheets(28).Range("a9:bz10000").ClearContents
Sheets(30).Range("a9:cc10000").ClearContents
Sheets(35).Range("a9:ce10000").ClearContents
Sheets(36).Range("a9:dz10000").ClearContents
Sheets(37).Range("a2:m10000").ClearContents
Sheets(38).Range("a2:f10000").ClearContents
Sheets(43).Range("b9:m10000").ClearContents
Sheets(44).Range("b9:m10000").UnMerge
Sheets(44).Range("b9:m10000").ClearContents
Sheets(46).Range("a2:h10000").ClearContents
MsgBox " „ ⁄„· ‰”Œ… «Õ Ì«ÿÌ… Ê ’›Ì— Ã„Ì⁄ «·»Ì«‰«  ··”‰… «·„«·Ì… «·ÃœÌœÂ ", , "—”«·…  ÊÃÌÂ"


Else
MsgBox "·ﬁœ ﬁ„  »«œŒ«· ﬂ·„… «·„—Ê— «·Œÿ√ Õ«Ê· „—Â À«‰ÌÂ «Ê « ’· »«·„»—„Ã", , "—”«·…  ‰»ÌÂ"

Exit Sub
End If
End Sub

Private Sub CommandButton18_Click()
Unload Me
UserForm47.Show

End Sub

Private Sub CommandButton19_Click()
UserForm48.Show
End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
ListView2.ListItems.Clear
TextBox1.Visible = True
TextBox2.Visible = True
TextBox3.Visible = True
TextBox4.Visible = True
TextBox9.Visible = True
TextBox1.Text = ""
TextBox2.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Visible = False
TextBox6.Visible = False
TextBox7.Visible = False
TextBox8.Visible = False
TextBox10.Visible = False

Label2.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
ListView1.Visible = False
ListView2.Visible = True
CommandButton9.Visible = True
CommandButton2.BackColor = RGB(84, 53, 110)
CommandButton2.ForeColor = &HFFFFC0
CommandButton8.BackColor = &HC00000
CommandButton8.ForeColor = &HFFFFC0

With Me.ListView2
.ListItems.Clear

Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(1).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(1).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(1).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(1).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(1).Cells(frw1, "bk")





Next frw1
Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next


UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox65.value = Format(UserForm1.TextBox65.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")

'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(60).Alignment = lvwColumnCenter
   ListView2.ColumnHeaders(61).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(62).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(63).Alignment = lvwColumnCenter
   














'=================================================




End With
End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"

End Sub

Private Sub CommandButton20_Click()
On Error GoTo emad

Dim PASS As Variant
Dim user_pass As Variant
PASS = "1110111"
 user_pass = InputBox("ﬁ„ »«œŒ«· ﬂ·„… «·„—Ê— ·Ì ”‰Ï ·ﬂ  ’›Ì— ﬁ«⁄œ… «·»Ì«‰« ", "ﬂ·„… «·„—Ê—")
If user_pass = PASS Then
Application.DisplayAlerts = False
Dim num_day As Integer
num_day = Day(Date)
If num_day = 26 Or num_day = 28 Or num_day = 28 Or num_day = 29 Or num_day = 30 Or num_day = 31 Or num_day Then
Sheets(55).Range("a9:bz8000").ClearContents
Application.ScreenUpdating = False
Sheets(55).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select
Selection.Copy Sheets(55).Range("A1:bz8000")

Application.ScreenUpdating = True
[a8].Select
Sheets(55).Select
Sheets(55).Activate
Application.ScreenUpdating = True
MsgBox " „ ‰ﬁ· «·»Ì«‰«  »‰Ã«Õ", vbDefaultButton1, "—”«·…  ‰»ÌÂ"

Else
MsgBox "Â‰«ﬂ Œ„”… «Ì«„ ›ﬁÿ ·ﬂ «·’·«ÕÌÂ ›Ì ‰ﬁ· »Ì«‰«  «·‘Â— «·„«÷Ì · ﬂÊ‰ „—Ã⁄« Õ”«»Ì« ÊÂÊ ÌÊ„ 26Ê27Ê28Ê29Ê30Ê31Ê1 „‰ ﬂ· ‘Â— ", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Exit Sub
End If

Else
MsgBox "·ﬁœ ﬁ„  »«œŒ«· ﬂ·„… «·„—Ê— «·Œÿ√ Õ«Ê· „—Â À«‰ÌÂ «Ê « ’· »«·„»—„Ã", , "—”«·…  ‰»ÌÂ"

Exit Sub
End If





'===============================================================

emad:
End Sub

Private Sub CommandButton21_Click()
Unload Me
UserForm58.Show
End Sub

Private Sub CommandButton22_Click()
Unload Me
UserForm62.Show
End Sub

Private Sub CommandButton23_Click()
If Sheets(4).Cells(1, "n") = 1 And Sheets(35).Cells(9, "a").value <> "" Then
MsgBox "·ﬁœ  „ Õ–› «·»Ì«‰«  «··„ €Ì—«  «·‘Â— «·„«÷Ì »‰Ã«Õ", vbOK, "—”«·…  ‰»ÌÂ"
Call copy_sheet35_tosheet50_1
Sheets(35).Range("a9:ce2000").Delete

Else
Exit Sub
End If
End Sub

Private Sub CommandButton24_Click()
Unload Me
UserForm66.Show
End Sub

Private Sub CommandButton25_Click()
Call Push_Update_To_GitHub
End Sub

Private Sub CommandButton3_Click()
Unload Me
UserForm69.Show

End Sub

Private Sub CommandButton4_Click()
Unload Me
UserForm68.Show

End Sub

Private Sub CommandButton5_Click()
Unload Me
UserForm45.Show

End Sub

Private Sub CommandButton6_Click()
Unload Me
UserForm46.Show

End Sub

Private Sub CommandButton7_Click()
On Error Resume Next
Unload Me
UserForm10.Show

End Sub

Private Sub CommandButton8_Click()
ListView1.ListItems.Clear
TextBox1.Visible = False
TextBox2.Visible = False
TextBox3.Visible = False
TextBox4.Visible = False
TextBox9.Visible = False

TextBox5.Visible = True
TextBox6.Visible = True
TextBox7.Visible = True
TextBox8.Visible = True
TextBox10.Visible = True

TextBox5.Text = ""
TextBox6.Text = ""
TextBox7.Text = ""
TextBox8.Text = ""
TextBox10.Text = ""
Label2.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
ListView1.Visible = True
ListView2.Visible = False
CommandButton9.Visible = False
On Error Resume Next

'=========================
CommandButton8.BackColor = RGB(84, 53, 110)
 CommandButton2.BackColor = &HC00000
 CommandButton8.ForeColor = &HFFFFC0
CommandButton2.ForeColor = &HFFFFC0


With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")

 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next






'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(60).Alignment = lvwColumnCenter
   ListView1.ColumnHeaders(61).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(62).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(63).Alignment = lvwColumnCenter
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"

End Sub

Private Sub CommandButton9_Click()
'On Error Resume Next
Dim item1 As ListItem
Dim last1, frw1 As Integer
Dim rngData As Range
Dim wksSource As Worksheet
 Dim fString As Variant
 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row



'=============================================

   Sheets(1).Activate
    
    
    For frw1 = 9 To last1
                
  If Sheets(1).Cells(frw1, "f") Like "*" & TextBox2.Text & "*" And Sheets(1).Cells(frw1, "g") Like "*" & TextBox3.Text & "*" _
  And Sheets(1).Cells(frw1, "j") >= TextBox4.Text And Sheets(1).Cells(frw1, "h") = TextBox9.Text Then
     
'==================================================

Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(1).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(1).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(1).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(1).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(1).Cells(frw1, "bk")



End If
Next frw1
End With

'============================================= «·Õ”«»«  «·ﬂ·ÌÂ









End Sub

Private Sub Frame2_Click()

End Sub

Private Sub TextBox1_AfterUpdate()
ListView2.ListItems.Clear
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"



Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"

End Sub

Private Sub TextBox1_Change()
On Error Resume Next
ListView2.ListItems.Clear
ListView2.Visible = True
ListView1.Visible = False
On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================





'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "e") Like "*" & TextBox1.Text & "*" Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub

Private Sub TextBox10_AfterUpdate()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "h") = TextBox10.Text Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox10_Change()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "h") = TextBox10.Text Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox2_AfterUpdate()
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "f") Like "*" & TextBox2.Text & "*" Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub



Private Sub TextBox2_Change()
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "f") Like "*" & TextBox2.Text & "*" Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub





Private Sub TextBox3_AfterUpdate()




ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "g") Like "*" & TextBox3.Text & "*" Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub



Private Sub TextBox3_Change()
On Error Resume Next
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "g") Like "*" & TextBox3.Text & "*" Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub



Private Sub TextBox4_AfterUpdate()
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "j") >= TextBox4.Text Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub

Private Sub TextBox4_Change()
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "j") >= TextBox4.Text Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub


Private Sub TextBox5_AfterUpdate()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "e") Like "*" & TextBox5.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox5_Change()
'If Sheets(6).Cells(frw1, "b").Value = Me.TextBox2.Text Or Sheets(6).Cells(frw1, "e") Like "*" & TextBox5.Text & "*" Then
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "e") Like "*" & TextBox5.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox6_AfterUpdate()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "f") Like "*" & TextBox6.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With

End Sub

Private Sub TextBox6_Change()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "f") Like "*" & TextBox6.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With

End Sub

Private Sub TextBox7_AfterUpdate()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "g") Like "*" & TextBox7.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox7_Change()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "g") Like "*" & TextBox7.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox8_AfterUpdate()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "j") >= TextBox8.Text Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox8_Change()
ListView1.Visible = True
ListView2.Visible = False
'On Error Resume Next
ListView1.ListItems.Clear
'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "j") >= TextBox8.Text Then
Set item1 = ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(3).Cells(frw1, "b")
item1.SubItems(2) = Sheets(3).Cells(frw1, "c")
item1.SubItems(3) = Sheets(3).Cells(frw1, "d")
item1.SubItems(4) = Sheets(3).Cells(frw1, "e")
item1.SubItems(5) = Sheets(3).Cells(frw1, "f")
item1.SubItems(6) = Sheets(3).Cells(frw1, "g")
item1.SubItems(7) = Sheets(3).Cells(frw1, "h")
item1.SubItems(8) = Sheets(3).Cells(frw1, "i")
item1.SubItems(9) = Sheets(3).Cells(frw1, "j")
item1.SubItems(10) = Sheets(3).Cells(frw1, "k")
item1.SubItems(11) = Sheets(3).Cells(frw1, "l")
item1.SubItems(12) = Sheets(3).Cells(frw1, "m")
item1.SubItems(13) = Sheets(3).Cells(frw1, "n")
item1.SubItems(14) = Sheets(3).Cells(frw1, "o")
item1.SubItems(15) = Sheets(3).Cells(frw1, "p")
item1.SubItems(16) = Sheets(3).Cells(frw1, "q")
item1.SubItems(17) = Sheets(3).Cells(frw1, "r")
item1.SubItems(18) = Sheets(3).Cells(frw1, "s")
item1.SubItems(19) = Sheets(3).Cells(frw1, "t")
item1.SubItems(20) = Sheets(3).Cells(frw1, "u")
item1.SubItems(21) = Sheets(3).Cells(frw1, "v")
item1.SubItems(22) = Sheets(3).Cells(frw1, "w")
item1.SubItems(23) = Sheets(3).Cells(frw1, "x")
item1.SubItems(24) = Sheets(3).Cells(frw1, "y")
item1.SubItems(25) = Sheets(3).Cells(frw1, "z")
item1.SubItems(26) = Sheets(3).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(3).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(3).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(3).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(3).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(3).Cells(frw1, "af")
item1.SubItems(32) = Sheets(3).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(3).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(3).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(3).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(3).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(3).Cells(frw1, "al")
item1.SubItems(38) = Sheets(3).Cells(frw1, "am")
item1.SubItems(39) = Sheets(3).Cells(frw1, "an")
item1.SubItems(40) = Sheets(3).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(3).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(3).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(3).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(3).Cells(frw1, "as")
item1.SubItems(45) = Sheets(3).Cells(frw1, "at")
item1.SubItems(46) = Sheets(3).Cells(frw1, "au")
item1.SubItems(47) = Sheets(3).Cells(frw1, "av")
item1.SubItems(48) = Sheets(3).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(3).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(3).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(3).Cells(frw1, "az")
item1.SubItems(52) = Sheets(3).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(3).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(3).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(3).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(3).Cells(frw1, "be")
item1.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub

Private Sub TextBox9_AfterUpdate()
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "h") = TextBox9.Text Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
End Sub

Private Sub TextBox9_Change()
ListView2.Visible = True
ListView1.Visible = False
'On Error Resume Next
ListView2.ListItems.Clear
'=========================

With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True




ListView2.Font = Bold
ListView2.Font.Size = 16
ListView2.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With Me.ListView2
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "h") = TextBox9.Text Then
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(2) = Sheets(1).Cells(frw1, "c")
item1.SubItems(3) = Sheets(1).Cells(frw1, "d")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(7) = Sheets(1).Cells(frw1, "h")
item1.SubItems(8) = Sheets(1).Cells(frw1, "i")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(10) = Sheets(1).Cells(frw1, "k")
item1.SubItems(11) = Sheets(1).Cells(frw1, "l")
item1.SubItems(12) = Sheets(1).Cells(frw1, "m")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(14) = Sheets(1).Cells(frw1, "o")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(16) = Sheets(1).Cells(frw1, "q")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(21) = Sheets(1).Cells(frw1, "v")
item1.SubItems(22) = Sheets(1).Cells(frw1, "w")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(25) = Sheets(1).Cells(frw1, "z")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(1).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(1).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(38) = Sheets(1).Cells(frw1, "am")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(40) = Sheets(1).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(1).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(1).Cells(frw1, "as")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(1).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(1).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(1).Cells(frw1, "az")
item1.SubItems(52) = Sheets(1).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(1).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(1).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(1).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(1).Cells(frw1, "be")
item1.SubItems(57) = Sheets(1).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView2.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
'ListView2.ForeColor = "blue"
End With
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

    Dim FinalW As Single, FinalH As Single
    Dim FinalL As Single, FinalT As Single
    Dim i As Double
    
    ' 1. Õ›Ÿ «·√»⁄«œ «·√’·Ì…
    FinalW = Me.Width: FinalH = Me.Height
    FinalL = Me.Left: FinalT = Me.Top

    ' 2. ‰ﬁÿ… «·»œ«Ì… (€Ì— „—∆Ì… ›Ì «·”‰ —)
    Me.Width = 0: Me.Height = 0
    
    ' 3. „—Õ·… "«·«‰»À«ﬁ «·Â«œ∆" (The Grand Reveal)
    ' Ã⁄·‰« «·ﬁ›“… 0.01 »œ·« „‰ 0.04 (Ì⁄‰Ì 100 ŒÿÊ… »œ·« „‰ 25) ·÷„«‰ «·‰⁄Ê„…
    For i = 0.01 To 1.1 Step 0.01
        Me.Width = FinalW * i
        Me.Height = FinalH * i
        
        ' „Ê«“‰… «·„Êﬁ⁄ »œﬁ… »ﬂ”· »»ﬂ”·
        Me.Left = FinalL + (FinalW - Me.Width) / 2
        Me.Top = FinalT + (FinalH - Me.Height) / 2
        
        ' —›⁄‰« «·«‰ Ÿ«— ·‹ 35 ·Ã⁄· «·⁄Ì‰  ” „ ⁄ »ﬂ· Õ—ﬂ…
        Sleep 35
        DoEvents
    Next i

    ' 4. „—Õ·… "«·«” ﬁ—«— «·„ÿ«ÿÌ" (Soft Bounce)
    ' «·⁄Êœ… „‰ 110% ≈·Ï 100% »»ÿ¡ ‘œÌœ
    For i = 1.1 To 1 Step -0.005
        Me.Width = FinalW * i
        Me.Height = FinalH * i
        Me.Left = FinalL + (FinalW - Me.Width) / 2
        Me.Top = FinalT + (FinalH - Me.Height) / 2
        
        Sleep 40
        DoEvents
    Next i
    
    ' «· √ﬂÌœ «·‰Â«∆Ì
    Me.Width = FinalW: Me.Height = FinalH
    Me.Left = FinalL: Me.Top = FinalT





Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

'================================== listview2
With ListView2
  .Gridlines = True
  .View = lvwReport
 .FullRowSelect = True

.ColumnHeaders.Add , , "«·—ﬁ„", 40
.ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
.ColumnHeaders.Add , , "—ﬁ„ «·„«” —", 120
.ColumnHeaders.Add , , "«·Ã‰”", 40
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 140
.ColumnHeaders.Add , , "«·⁄‰Ê«‰ «·ÊŸÌ›Ì", 100
.ColumnHeaders.Add , , " «·ﬁ”„", 100
.ColumnHeaders.Add , , " «·œ—Ã…", 50
.ColumnHeaders.Add , , "«·„—Õ·…", 50
.ColumnHeaders.Add , , "«·—« » «·«”„Ì ", 140
.ColumnHeaders.Add , , " «Ì«„ «·€Ì«»"
.ColumnHeaders.Add , , " «Ì«„ «·œÊ«„", 120
.ColumnHeaders.Add , , " ‰”»… «·‘Â«œ…", 120
.ColumnHeaders.Add , , "„Œ’’«  «·‘Â«œ…", 40
.ColumnHeaders.Add , , "‰”»… «·Â‰œ”ÌÂ ", 140
.ColumnHeaders.Add , , " „Œ’’«  «·Â‰œ”ÌÂ", 40
.ColumnHeaders.Add , , "‰”»… «·„‰’» ", 100
.ColumnHeaders.Add , , "„Œ’’«  «·„‰’» ", 50
.ColumnHeaders.Add , , "„Œ’’«  „Êﬁ⁄ Ã€—«›Ì", 50
.ColumnHeaders.Add , , "„Œ’’«  «ÿ›«·  ", 140
.ColumnHeaders.Add , , "„Œ’’«  “ÊÃÌÂ", 40
.ColumnHeaders.Add , , " ⁄œœ «·«ÿ›«·", 120
.ColumnHeaders.Add , , "«·Õ«·Â «·«Ã „«⁄ÌÂ ", 120
.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄ «·÷—Ì»Ì", 40
.ColumnHeaders.Add , , "„Œ’’«  ŒÿÊ—… „»·€ ", 140
.ColumnHeaders.Add , , "„Œ’’«  „Â‰ÌÂ ", 40
.ColumnHeaders.Add , , " „‰”»… „Œ’’«  „Â‰ÌÂ ", 50
.ColumnHeaders.Add , , "‰”»… «·ŒÿÊ—Â", 50
.ColumnHeaders.Add , , " „Œ’’«  «·ŒÿÊ—… ", 140
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·„Œ’’« ", 120
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·«” Õﬁ«ﬁ« ", 120
.ColumnHeaders.Add , , "  Êﬁ›«   ﬁ«⁄œÌÂ 10", 100
.ColumnHeaders.Add , , "« Êﬁ›«   ﬁ«⁄œÌÂ 15%", 40
.ColumnHeaders.Add , , "«· «„Ì‰ «·’ÕÌ ", 50
.ColumnHeaders.Add , , " «ÃÊ— «·‰ﬁ·", 40
.ColumnHeaders.Add , , " —”„ «·ÿ«»⁄", 100
.ColumnHeaders.Add , , " Ê“«—… «·„«·ÌÂ", 50
.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄1", 50
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄1  ", 140
.ColumnHeaders.Add , , "«” ﬁÿ«⁄ „ ›—ﬁ", 40
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄ ", 120
.ColumnHeaders.Add , , " ‰«œÌ «·ÂÌ√… «·’‰«⁄«  «·Õ—»Ì", 120
.ColumnHeaders.Add , , "ÃÂ… «·«” ﬁÿ«⁄ ", 40
.ColumnHeaders.Add , , "«” ﬁÿ«⁄ 2 ", 140
.ColumnHeaders.Add , , "ÃÂ… «·«” ﬁÿ«⁄ 2 ", 40
.ColumnHeaders.Add , , "«” ﬁÿ«⁄«  —⁄«ÌÂ", 100
.ColumnHeaders.Add , , " „»·€ «·€Ì«»", 50
.ColumnHeaders.Add , , " „Ã„Ê⁄ «·«” ﬁÿ«⁄« ", 50
.ColumnHeaders.Add , , "«·—« » «·’«›Ì", 150
.ColumnHeaders.Add , , " «·—« » «·’«›Ì „⁄ »Ì«‰ „Êﬁ› „Êﬁ› «·Œœ„… ", 140
.ColumnHeaders.Add , , " ÿ—Ìﬁ… «·œ›⁄ ", 140
.ColumnHeaders.Add , , " —ﬁ„ «·«Ì»«‰ ", 140
.ColumnHeaders.Add , , "  «—ÌŒ ‰ﬁ·Â Ê«÷«› Â «·Ï «·‰Ÿ«„  ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ’Ê·Â ⁄·Ï «·⁄·«Ê… ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ’Ê·Â ⁄·Ï «· —›Ì⁄ ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ–›Â Ê —ÕÌ·Â ", 100


.ColumnHeaders.Add , , " ÃÂ… «·œ›⁄ ", 140
.ColumnHeaders.Add , , " „Êﬁ› «·Œœ„… ", 140
.ColumnHeaders.Add , , " «—ÌŒ Õ’Ê·Â ⁄·Ï ﬂ «» ‘ﬂ— Ê ﬁœÌ— ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ–›Â Ê —ÕÌ·Â ", 0
.ColumnHeaders.Add , , " «·≈÷«›… «·„«·ÌÂ «·„ƒﬁ Â", 140
.ColumnHeaders.Add , , " «·„·«ÕŸ«  ", 140
.ColumnHeaders.Add , , " «·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ  ", 140
.ColumnHeaders.Add , , " «·„·«ÕŸ«  ", 140




  ListView2.Font = Bold
  ListView2.Font.Size = 16
 ListView2.Font.Name = "PT hrading"

End With

'==========================================================================ListView1

With ListView1
 .Gridlines = True
  .View = lvwReport
  .FullRowSelect = True

.ColumnHeaders.Add , , "«·—ﬁ„", 40
.ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
.ColumnHeaders.Add , , "—ﬁ„ «·„«” —", 120
.ColumnHeaders.Add , , "«·Ã‰”", 40
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 140
.ColumnHeaders.Add , , "«·⁄‰Ê«‰ «·ÊŸÌ›Ì", 100
.ColumnHeaders.Add , , " «·ﬁ”„", 100
.ColumnHeaders.Add , , " «·œ—Ã…", 50
.ColumnHeaders.Add , , "«·„—Õ·…", 50
.ColumnHeaders.Add , , "«·—« » «·«”„Ì ", 140
.ColumnHeaders.Add , , " «Ì«„ «·€Ì«»"
.ColumnHeaders.Add , , " «Ì«„ «·œÊ«„", 120
.ColumnHeaders.Add , , " ‰”»… «·‘Â«œ…", 120
.ColumnHeaders.Add , , "„Œ’’«  «·‘Â«œ…", 40
.ColumnHeaders.Add , , "‰”»… «·Â‰œ”ÌÂ ", 140
.ColumnHeaders.Add , , " „Œ’’«  «·Â‰œ”ÌÂ", 40
.ColumnHeaders.Add , , "‰”»… «·„‰’» ", 100
.ColumnHeaders.Add , , "„Œ’’«  «·„‰’» ", 50
.ColumnHeaders.Add , , "„Œ’’«  „Êﬁ⁄ Ã€—«›Ì", 50
.ColumnHeaders.Add , , "„Œ’’«  «ÿ›«·  ", 140
.ColumnHeaders.Add , , "„Œ’’«  “ÊÃÌÂ", 40
.ColumnHeaders.Add , , " ⁄œœ «·«ÿ›«·", 120
.ColumnHeaders.Add , , "«·Õ«·Â «·«Ã „«⁄ÌÂ ", 120
.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄ «·÷—Ì»Ì", 40
.ColumnHeaders.Add , , "„Œ’’«  ŒÿÊ—… „»·€ ", 140
.ColumnHeaders.Add , , "„Œ’’«  „Â‰ÌÂ ", 40
.ColumnHeaders.Add , , " „‰”»… „Œ’’«  „Â‰ÌÂ ", 50
.ColumnHeaders.Add , , "‰”»… «·ŒÿÊ—Â", 50
.ColumnHeaders.Add , , " „Œ’’«  «·ŒÿÊ—… ", 140
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·„Œ’’« ", 120
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·«” Õﬁ«ﬁ« ", 120
.ColumnHeaders.Add , , "  Êﬁ›«   ﬁ«⁄œÌÂ 10", 100
.ColumnHeaders.Add , , "« Êﬁ›«   ﬁ«⁄œÌÂ 15%", 40
.ColumnHeaders.Add , , "«· «„Ì‰ «·’ÕÌ ", 50
.ColumnHeaders.Add , , " «ÃÊ— «·‰ﬁ·", 40
.ColumnHeaders.Add , , " —”„ «·ÿ«»⁄", 100
.ColumnHeaders.Add , , " Ê“«—… «·„«·ÌÂ", 50
.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄1", 50
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄1  ", 140
.ColumnHeaders.Add , , "«” ﬁÿ«⁄ „ ›—ﬁ", 40
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄ ", 120
.ColumnHeaders.Add , , " ‰«œÌ «·ÂÌ√… «·’‰«⁄«  «·Õ—»Ì", 120
.ColumnHeaders.Add , , "ÃÂ… «·«” ﬁÿ«⁄ ", 40
.ColumnHeaders.Add , , "«” ﬁÿ«⁄ 2 ", 140
.ColumnHeaders.Add , , "ÃÂ… «·«” ﬁÿ«⁄ 2 ", 40
.ColumnHeaders.Add , , "«” ﬁÿ«⁄«  —⁄«ÌÂ", 100
.ColumnHeaders.Add , , " „»·€ «·€Ì«»", 50
.ColumnHeaders.Add , , " „Ã„Ê⁄ «·«” ﬁÿ«⁄« ", 50
.ColumnHeaders.Add , , "«·—« » «·’«›Ì", 150
.ColumnHeaders.Add , , " «·—« » «·’«›Ì „⁄ »Ì«‰ „Êﬁ› „Êﬁ› «·Œœ„… ", 140
.ColumnHeaders.Add , , " ÿ—Ìﬁ… «·œ›⁄ ", 140
.ColumnHeaders.Add , , " —ﬁ„ «·«Ì»«‰ ", 140
.ColumnHeaders.Add , , "  «—ÌŒ ‰ﬁ·Â Ê«÷«› Â «·Ï «·‰Ÿ«„  ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ’Ê·Â ⁄·Ï «·⁄·«Ê… ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ’Ê·Â ⁄·Ï «· —›Ì⁄ ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ–›Â Ê —ÕÌ·Â ", 100


.ColumnHeaders.Add , , " ÃÂ… «·œ›⁄ ", 140
.ColumnHeaders.Add , , " „Êﬁ› «·Œœ„… ", 140
.ColumnHeaders.Add , , " «—ÌŒ Õ’Ê·Â ⁄·Ï ﬂ «» ‘ﬂ— Ê ﬁœÌ— ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ–›Â Ê —ÕÌ·Â ", 0




 ListView1.Font = Bold
  ListView1.Font.Size = 16
  ListView1.Font.Name = "PT hrading"

End With

'==========================================================================listview1
With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

'=================================================
If Sheets(4).Cells(1, "n") > 1 Or Sheets(35).Cells(9, "a").value = "" Then
Exit Sub
End If

If Sheets(4).Cells(1, "n") = 1 And Sheets(35).Cells(9, "a").value <> "" _
And MsgBox("”Ê› Ì „ Õ–› «·»Ì«‰«  «·ﬁœÌ„… ··„ €Ì—«  «·‘ÂÊ— «·„«÷ÌÂ «÷€ÿ ﬂ‰”· «–« ·„  ﬂ‰  —Ìœ Õ–›Â« Õ«·Ì« ", vbOK, "—”«·…  ‰»Â") = vbOK Then
MsgBox "·ﬁœ  „ Õ–› «·»Ì«‰«  «··„ €Ì—«  «·‘Â— «·„«÷Ì »‰Ã«Õ", vbOK, "—”«·…  ‰»ÌÂ"
Call copy_sheet35_tosheet50_1
Sheets(35).Range("a9:ce2000").Delete

Else
Exit Sub
End If


End With
'======================================

End Sub

Private Sub UserForm_Initialize()

  Dim ctrl As Control
    Dim MaxW As Double, MaxH As Double, FinalZoom As Double
    
    ' 1.  ”—Ì⁄ «·√œ«¡ »≈Ìﬁ«› «· ÕœÌÀ «··ÕŸÌ
    Application.ScreenUpdating = False
    
    ' 2. Õ”«» «·√»⁄«œ ( „ œ„Ã «·„ €Ì—«  · ﬁ·Ì· «” Â·«ﬂ «·–«ﬂ—…)
    For Each ctrl In Me.Controls
        If ctrl.Visible Then
            If ctrl.Left + ctrl.Width > MaxW Then MaxW = ctrl.Left + ctrl.Width
            If ctrl.Top + ctrl.Height > MaxH Then MaxH = ctrl.Top + ctrl.Height
        End If
    Next ctrl
    
    ' 3.  ÿ»Ìﬁ «·√»⁄«œ „»«‘—…
    With Me
        .Top = 0
        .Left = 0
        .Width = Application.Width
        .Height = Application.Height
        
        ' 4. Õ”«» «·“ÊÊ„ »ŒÿÊ… —Ì«÷Ì… Ê«Õœ… (Shorthand Math)
        FinalZoom = (Application.Width / (MaxW + 15)) * 100
        Dim zH As Double: zH = (Application.Height / (MaxH + 30)) * 100
        If zH < FinalZoom Then FinalZoom = zH
        
        ' 5. ’„«„ «·√„«‰ Ê«· ÿ»Ìﬁ
        If FinalZoom < 10 Then FinalZoom = 10 Else If FinalZoom > 400 Then FinalZoom = 400
        .Zoom = FinalZoom
    End With

    Application.ScreenUpdating = True




'=======================================================================


TextBox1.Visible = False
TextBox2.Visible = False
TextBox3.Visible = False
TextBox4.Visible = False
TextBox5.Visible = False
TextBox6.Visible = False
TextBox7.Visible = False
TextBox8.Visible = False
TextBox9.Visible = False
TextBox10.Visible = False
Label2.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
ListView1.Visible = False
ListView1.Visible = False
CommandButton9.Visible = False


       Call sigen_connect.Sigen_Main_Engine

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then
Cancel = True
End If

    Dim i As Double
    Dim FW As Single, FH As Single, FL As Single, FT As Single
    
    ' Õ›Ÿ «·√»⁄«œ «·Õ«·Ì… ﬁ»· «·»œ¡ »«·«‰ﬂ„«‘
    FW = Me.Width: FH = Me.Height
    FL = Me.Left: FT = Me.Top
    
    ' Õ—ﬂ… "«· ·«‘Ì ··œ«Œ·" („‰ 100% ≈·Ï 0%)
    For i = 1 To 0 Step -0.05
        Me.Width = FW * i
        Me.Height = FH * i
        
        ' «·Õ›«Ÿ ⁄·Ï «· „—ﬂ“ √À‰«¡ «·«Œ ›«¡
        Me.Left = FL + (FW - Me.Width) / 2
        Me.Top = FT + (FH - Me.Height) / 2
        
        Sleep 15 ' ”—⁄… «·«‰ﬂ„«‘ («Ã⁄·Â« √”—⁄ ﬁ·Ì·« „‰ «·› Õ · »œÊ —‘Ìﬁ…)
        DoEvents
    Next i
     
    
    
   Application.OnTime Sigen_NextPulse, "sigen_connect.Sigen_Trigger", , False
     
End Sub


