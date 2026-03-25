VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "«·‰Ÿ«„ «·„«·Ì"
   ClientHeight    =   13812
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   20604
   OleObjectBlob   =   "UserForm1.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Public DictOld As Object
Dim IsLoading As Boolean
Dim OrigW As Single
Dim OrigH As Single
Dim OrigColor As Long

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

Function rightToLeftListView()
Dim Columns As Long
Dim i As Long
With ListView1
Columns = .ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next
End With
End Function







Private Sub CheckBox1_Change()

 
End Sub



Private Sub CheckBox2_Click()
If CheckBox2.value = True Then
CommandButton15.Enabled = True
Me.TextBox55.Text = ""
Me.TextBox54.Text = ""
Me.TextBox59.Text = ""
Else
CommandButton15.Enabled = False
End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.value = True Then
Label124.Enabled = True
Else
Label124.Enabled = False
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox4.value = False Then
TextBox105.value = ""
End If
End Sub

Private Sub CheckBox5_Change()
If CheckBox5.value = True Then
        ' ≈ŸÂ«— «··Ì»· √Ê·«
        Label127.Visible = True
        ' ≈⁄«œ… «··Ê‰ «·√’·Ì («·„Œ“‰ ›Ì «·„ €Ì— ⁄‰œ  ‘€Ì· «·›Ê—„)
        Label127.BackColor = OrigColor
        '  ‘€Ì· Õ—ﬂ… «· „œœ ··√»⁄«œ «·√’·Ì…
        AnimateLabel OrigW, OrigH
    End If
End Sub

Private Sub CheckBox5_Click()
    ' „‰⁄  ‰›Ì– «·ﬂÊœ ⁄‰œ «· ‰ﬁ· ›Ì «··”  »Êﬂ”
    If IsLoading = True Then Exit Sub
    
    Dim EmpID As String: EmpID = Me.TextBox2.value
    Dim PasswordEntry As String

    ' «·Õ«·… «·√Ê·Ï: ≈–« ﬁ«„ «·„” Œœ„ » ›⁄Ì· «·ÃÌﬂ »Êﬂ” (Ê÷⁄ ’Õ)
    If Me.CheckBox5.value = True Then
        PasswordEntry = InputBox("Ì—ÃÏ ≈œŒ«· «·»«”Ê—œ ·· ›⁄Ì· Ê«·«” À‰«¡ «·÷—Ì»Ì:", " Õﬁﬁ «·√„«‰")
        
        If PasswordEntry <> "" And PasswordEntry = EmpID Then
            '  ÕœÌÀ «·‘Ì  ··ﬁÌ„… True
            Call UpdateCGInSheet(EmpID, True)
            MsgBox " „  ›⁄Ì· «·«” À‰«¡ «·÷—Ì»Ì ›Ì «·‘Ì .", vbInformation
        Else
            ' ≈–« ﬂ«‰ «·»«”Ê—œ Œÿ√° ‰⁄Ìœ «·ÃÌﬂ »Êﬂ” ·Ê÷⁄Â «·”«»ﬁ »’„ 
            MsgBox "»«”Ê—œ Œÿ√° ·« Ì„ﬂ‰ «· ›⁄Ì·.", vbCritical
            IsLoading = True
            Me.CheckBox5.value = False
            IsLoading = False
        End If

    ' «·Õ«·… «·À«‰Ì…: ≈–« ﬁ«„ «·„” Œœ„ »≈·€«¡ «· ›⁄Ì· (≈“«·… «·’Õ)
    Else
        '  ÕÊÌ· «·ﬁÌ„… ›Ì «·‘Ì  ≈·Ï False ›Ê—« ··„ÊŸ› «·Õ«·Ì
        Call UpdateCGInSheet(EmpID, False)
        MsgBox " „ ≈·€«¡ «·«” À‰«¡ «·÷—Ì»Ì Ê ÕœÌÀ «·⁄„Êœ CG ≈·Ï False.", vbExclamation
    End If
    
   
    
    
End Sub

Private Sub CheckBox5_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub ComboBox1_AfterUpdate()
 ' «·«⁄“»
'  «·„ÿ·ﬁ Ê·Ì” ·œÌÂ Õ÷«‰… ··√Ê·«œ
'  «·„ÊŸ›Â «·„›’Ê· œŒ·Â« ⁄‰ œŒ· “ÊÃÂ«
If TextBox22.value = 0 Then
Exit Sub
End If
If ComboBox1.value = "«·«⁄“»" Or ComboBox1.value = "«·„ÿ·ﬁ Ê·Ì” ·œÌÂ Õ÷«‰… ··√Ê·«œ" Or ComboBox1.value = "«·„ÊŸ›Â «·„›’Ê· œŒ·Â« ⁄‰ œŒ· “ÊÃÂ«" Then
MsgBox "Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—", vbCritical, "—”«·…  ‰»ÌÂ"
TextBox22.value = 0
Exit Sub
End If
End Sub

Private Sub ComboBox1_Change()
TextBox23.value = ComboBox1.value
ComboBox1.Visible = False
End Sub

Private Sub ComboBox1_DropButtonClick()

TextBox23.value = ComboBox1.value

End Sub

Private Sub ComboBox10_AfterUpdate()

If val(UserForm1.TextBox27) > 0 And val(UserForm1.TextBox29) > 0 Then
UserForm1.TextBox29 = 0
ComboBox11.value = 0
'MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
UserForm1.TextBox29 = UserForm1.TextBox29
End If
'=================================

End Sub

Private Sub ComboBox10_Change()
On Error Resume Next

If TextBox10.value = "" Then
TextBox10.value = 0

Else
TextBox27.value = TextBox10.value * ComboBox10.value
End If

If val(UserForm1.TextBox27) > 0 And val(UserForm1.TextBox29) > 0 Then
UserForm1.TextBox29 = 0
ComboBox11.value = 0
'MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
UserForm1.TextBox29 = UserForm1.TextBox29
End If
'===============================================

End Sub

Private Sub ComboBox11_AfterUpdate()

If val(UserForm1.TextBox27) > 0 And val(UserForm1.TextBox29) > 0 Then
UserForm1.TextBox27 = 0
ComboBox10.value = 0
MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
UserForm1.TextBox27 = UserForm1.TextBox27
End If
'===========================


End Sub

Private Sub ComboBox11_Change()
On Error Resume Next

If TextBox10.value = "" Then
TextBox10.value = 0
Else
TextBox29.value = TextBox10.value * ComboBox11.value
End If

If val(UserForm1.TextBox27) > 0 And val(UserForm1.TextBox29) > 0 Then
UserForm1.TextBox27 = 0
ComboBox10.value = 0
MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
UserForm1.TextBox27 = UserForm1.TextBox27
End If


End Sub

Private Sub ComboBox12_AfterUpdate()
TextBox61.Text = ComboBox12.value
End Sub

Private Sub ComboBox12_Change()
TextBox61.Text = ComboBox12.value
End Sub

Private Sub ComboBox13_AfterUpdate()
TextBox53.Text = ComboBox13.value
End Sub

Private Sub ComboBox13_Change()
TextBox53.Text = ComboBox13.value
End Sub

Private Sub ComboBox14_Change()
If ComboBox14.value = "ŒÿÊ—… ”«∆ﬁ" Then
TextBox25.value = 87500
Else
TextBox25.value = 0
End If
End Sub

Private Sub ComboBox15_AfterUpdate()


Call UpdateSummaryFormula

ComboBox3.value = ""

Sheets(1).Cells(4, "cb").value = ComboBox15.value

If Sheets(1).Cells(4, "cb").value = ComboBox15.value Then

TextBox93.value = Sheets(1).Cells(4, "cd")
TextBox94.value = Sheets(1).Cells(4, "ce")
TextBox97.value = Sheets(1).Cells(4, "cf")
TextBox98.value = Sheets(1).Cells(4, "cg")
TextBox95.value = Sheets(1).Cells(4, "ch")
TextBox96.value = Sheets(1).Cells(4, "cc")
TextBox107.value = Sheets(1).Cells(4, "ci")
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox93.value = Format(UserForm1.TextBox93.value, "#,## IQD")
UserForm1.TextBox94.value = Format(UserForm1.TextBox94.value, "#,## IQD")
UserForm1.TextBox97.value = Format(UserForm1.TextBox97.value, "#,## IQD")
UserForm1.TextBox98.value = Format(UserForm1.TextBox98.value, "#,## IQD")
UserForm1.TextBox95.value = Format(UserForm1.TextBox95.value, "#,## IQD")
UserForm1.TextBox96.value = Format(UserForm1.TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & ComboBox15.value
End If
'===========================================
ListBox1.Clear

TextBox72.Text = ""
Sheets(1).Activate
Dim frow1 As Integer
     ss = Sheets(1).Cells(Rows.count, 5).End(xlUp).row
     
     For frow1 = 9 To ss
    If ComboBox15.value = Sheets(1).Cells(frow1, 58) Then
        ListBox1.AddItem
        ListBox1.List(ListBox1.ListCount - 1, 0) = Cells(frow1, 5).value
       
    End If
    
Next
Dim last2 As Integer
     last2 = Sheets(1).Cells(Rows.count, 5).End(xlUp).row
     
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("bf9:bf" & last2), Me.ComboBox15.Text)


'================================ Õ„Ì· «·»Ì«‰«  ›Ì «· listview

With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If ComboBox15.value = Sheets(1).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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

UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")

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
   End With

'===============================================================




End Sub

Private Sub ComboBox15_Change()

ComboBox3.value = ""

Sheets(1).Cells(4, "cb").value = ComboBox15.value

If Sheets(1).Cells(4, "cb").value = ComboBox15.value Then

TextBox93.value = Sheets(1).Cells(4, "cd")
TextBox94.value = Sheets(1).Cells(4, "ce")
TextBox97.value = Sheets(1).Cells(4, "cf")
TextBox98.value = Sheets(1).Cells(4, "cg")
TextBox95.value = Sheets(1).Cells(4, "ch")
TextBox96.value = Sheets(1).Cells(4, "cc")
  TextBox107.value = Sheets(1).Cells(4, "ci")
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox93.value = Format(UserForm1.TextBox93.value, "#,## IQD")
UserForm1.TextBox94.value = Format(UserForm1.TextBox94.value, "#,## IQD")
UserForm1.TextBox97.value = Format(UserForm1.TextBox97.value, "#,## IQD")
UserForm1.TextBox98.value = Format(UserForm1.TextBox98.value, "#,## IQD")
UserForm1.TextBox95.value = Format(UserForm1.TextBox95.value, "#,## IQD")
UserForm1.TextBox96.value = Format(UserForm1.TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & ComboBox15.value
End If

'===================================
ListBox1.Clear

Sheets(1).Activate
Dim frow1 As Integer
     ss = Sheets(1).Cells(Rows.count, 5).End(xlUp).row
     
     For frow1 = 9 To ss
    If ComboBox15.value = Sheets(1).Cells(frow1, 58) Then
        ListBox1.AddItem
        ListBox1.List(ListBox1.ListCount - 1, 0) = Cells(frow1, 5).value
       
    End If
    
Next
Dim last2 As Integer
     last2 = Sheets(1).Cells(Rows.count, 5).End(xlUp).row
     
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("bf9:bf" & last2), Me.ComboBox15.Text)


'================================ Õ„Ì· «·»Ì«‰«  ›Ì «· listview

With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If ComboBox15.value = Sheets(1).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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

UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")

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
   End With
'================================



End Sub

Private Sub ComboBox2_Change()
ListView1.ListItems.Clear
End Sub

Private Sub ComboBox3_AfterUpdate()
On Error Resume Next

Sheets(1).Cells(4, "cb").value = ComboBox3.value
TextBox58.Text = ComboBox3.value
TextBox100.Text = ComboBox3.Column(1)

If Sheets(1).Cells(4, "cb").value = ComboBox3.value Then

TextBox93.value = Sheets(1).Cells(4, "cd")
TextBox94.value = Sheets(1).Cells(4, "ce")
TextBox97.value = Sheets(1).Cells(4, "cf")
TextBox98.value = Sheets(1).Cells(4, "cg")
TextBox95.value = Sheets(1).Cells(4, "ch")
TextBox96.value = Sheets(1).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox93.value = Format(UserForm1.TextBox93.value, "#,## IQD")
UserForm1.TextBox94.value = Format(UserForm1.TextBox94.value, "#,## IQD")
UserForm1.TextBox97.value = Format(UserForm1.TextBox97.value, "#,## IQD")
UserForm1.TextBox98.value = Format(UserForm1.TextBox98.value, "#,## IQD")
UserForm1.TextBox95.value = Format(UserForm1.TextBox95.value, "#,## IQD")
UserForm1.TextBox96.value = Format(UserForm1.TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & TextBox58.Text
End If


End Sub

Private Sub ComboBox3_Change()
On Error Resume Next

TextBox58.Text = ComboBox3.value
TextBox100.Text = ComboBox3.Column(1)


Sheets(1).Cells(4, "cb").value = ComboBox3.value
If Sheets(1).Cells(4, "cb").value = ComboBox3.value Then

TextBox93.value = Sheets(1).Cells(4, "cd")
TextBox94.value = Sheets(1).Cells(4, "ce")
TextBox97.value = Sheets(1).Cells(4, "cf")
TextBox98.value = Sheets(1).Cells(4, "cg")
TextBox95.value = Sheets(1).Cells(4, "ch")
TextBox96.value = Sheets(1).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox93.value = Format(UserForm1.TextBox93.value, "#,## IQD")
UserForm1.TextBox94.value = Format(UserForm1.TextBox94.value, "#,## IQD")
UserForm1.TextBox97.value = Format(UserForm1.TextBox97.value, "#,## IQD")
UserForm1.TextBox98.value = Format(UserForm1.TextBox98.value, "#,## IQD")
UserForm1.TextBox95.value = Format(UserForm1.TextBox95.value, "#,## IQD")
UserForm1.TextBox96.value = Format(UserForm1.TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & TextBox58.Text
End If
If ComboBox3.value = "⁄ﬁœ" Then
TextBox101.Visible = True
TextBox101.value = ""
TextBox10.Visible = False
Label123.Visible = True
Else
TextBox101.Visible = False
TextBox10.Visible = True
End If
If ComboBox3.value = "„” „— »«·Œœ„…" Then
Label123.Visible = False
TextBox10.Visible = True
End If
End Sub

Private Sub ComboBox6_Change()
On Error GoTo emad
TextBox14.value = TextBox10.value * ComboBox6.value
emad:
End Sub

Private Sub ComboBox7_Change()
On Error Resume Next
If TextBox10.value = "" Then
TextBox10.value = 0
Else
TextBox16.value = TextBox10.value * ComboBox7.value
End If
End Sub

Private Sub ComboBox8_Change()
On Error Resume Next
If TextBox10.value = "" Then
Exit Sub
Else
TextBox18.value = TextBox10.value * ComboBox8.value
End If
End Sub

Private Sub CommandButton1_Click()
Unload Me
 Application.Quit




End Sub

Private Sub CommandButton10_Click()

End Sub

Private Sub CommandButton11_Click()



End Sub

Private Sub CommandButton12_Click()



End Sub

Private Sub CommandButton13_Click()


End Sub

Private Sub CommandButton14_Click()



End Sub

Private Sub CommandButton100_Click()
 
End Sub

Private Sub CommandButton15_Click()

 UserForm12.Show
 
End Sub

Private Sub CommandButton16_Click()
On Error Resume Next

If ComboBox3.value = "„” „— »«·Œœ„…" Then
TextBox100.value = 1
TextBox58.value = "„” „— »«·Œœ„…"
End If
If ComboBox3.value = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Then
TextBox100.value = 2
TextBox58.value = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« "
End If
If ComboBox3.value = " «Ã«“… »œÊ‰ —« »" Then
TextBox100.value = 3
TextBox58.value = "«Ã«“… »œÊ‰ —« »"
End If
If ComboBox3.value = "≈Ã«“… „⁄Ì· «·„ ›—€" Then
TextBox100.value = 4
TextBox58.value = "≈Ã«“… „⁄Ì· «·„ ›—€"
End If
If ComboBox3.value = "—›⁄ Ìœ «Ê ≈Ìﬁ«› ’—›" Then
TextBox100.value = 5
TextBox58.value = "—›⁄ Ìœ «Ê ≈Ìﬁ«› ’—›"
End If

If ComboBox3.value = "«„Ê„Â «Ê· 6 «‘Â—" Then
TextBox100.value = 6
TextBox58.value = "«„Ê„Â «Ê· 6 «‘Â—"
End If

If ComboBox3.value = "«„Ê„Â À«‰Ì 6 «‘Â—" Then
TextBox100.value = 7
TextBox58.value = "«„Ê„Â À«‰Ì 6 «‘Â—"
End If

If ComboBox3.value = " ‰”Ì»" Then
TextBox100.value = 8
TextBox58.value = " ‰”Ì»"
End If



If ComboBox3.value = "⁄ﬁœ" Then
TextBox100.value = 9
TextBox58.value = "⁄ﬁœ"
End If
If ComboBox3.value = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" Then
TextBox100.value = 10
TextBox58.value = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ"
End If
If ComboBox3.value = "—« » Ã“∆Ì" Then
TextBox100.value = 11
TextBox58.value = "—« » Ã“∆Ì"
End If
If ComboBox3.value = "·œÌÂ  ÷„Ì‰" Then
TextBox100.value = 12
TextBox58.value = "·œÌÂ  ÷„Ì‰"
End If
If ComboBox3.value = "·œÌÂ ”·›" Then
TextBox100.value = 13
TextBox58.value = "·œÌÂ ”·›"
End If















'=============================================== ‘—Êÿ «·‰Ÿ«„ «·÷—Ì»Ì
If ComboBox4.value = "«‰ÀÏ" And TextBox23.Text = "«·„ÊŸ› «·„ “ÊÃ Ê “ÊÃ Â —»… »Ì " Then
MsgBox " «·‰Ÿ«„ «·Ÿ—Ì»Ì Ì—›÷ Â–Â «·„⁄«ÌÌ— ·ﬂÊ‰Â« «‰ÀÏ Ê·« Ì„ﬂ‰ «‰  ‰ÿ»ﬁ ⁄·Ï “ÊÃÂ« —»  »Ì øøø ", vbCritical, "—”«·…  ‰»ÌÂ"
Exit Sub
ElseIf TextBox23.Text = "«·«⁄“»" And TextBox22.Text > 0 Then
MsgBox "Â‰«ﬂ ⁄œ„  Ê«›ﬁ „‰ÿﬁÌ ··‰Ÿ«„ «·÷—Ì»Ì ﬂ√‰ ÌﬂÊ‰ «⁄“» Ê·œÌÂ «ÿ›«· øøø ", vbCritical, "—”«·…  ‰»ÌÂ"
Exit Sub
Else







'===================================================
Dim lr, nrow As Integer

lr = Sheets(1).Cells(Rows.count, 2).End(xlUp).row
For nrow = 9 To lr
If Sheets(1).Cells(nrow, "b").value = TextBox2.Text And Sheets(1).Cells(nrow, "e").value = TextBox5.Text Then

Sheets(1).Cells(nrow, "by") = Sheets(1).Cells(nrow, "aw")
Sheets(1).Cells(nrow, "bi") = TextBox61.Text












'================================
Select Case TextBox100.value
'==================================================„” „— »«·Œœ„…
Case 1
Dim lr1, nrow1 As Integer
lr1 = Sheets(1).Cells(Rows.count, 2).End(xlUp).row
For nrow1 = 9 To lr1
If UserForm1.TextBox8.value >= 8 And CheckBox4.value = True _
And Sheets(1).Cells(nrow1, "b").value = TextBox2.Text And Sheets(1).Cells(nrow1, "e").value = TextBox5.Text Then
Sheets(1).Cells(nrow1, "cd") = "true"
Sheets(9).Cells(18, "g") = TextBox8.Text
Sheets(9).Cells(18, "f") = TextBox9.Text
TextBox10.Text = Sheets(9).Cells(18, "h")
UserForm1.TextBox105.value = UserForm1.TextBox10.value + UserForm1.TextBox10.value * 0.5 - UserForm1.TextBox10.value
ComboBox10.value = 0
ComboBox11.value = 0
 UserForm1.TextBox27.value = 0
 UserForm1.TextBox29.value = 0
 TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
  Call mastmer_5dma_22
ElseIf CheckBox4.value = False _
And Sheets(1).Cells(nrow1, "b").value = TextBox2.Text And Sheets(1).Cells(nrow1, "e").value = TextBox5.Text Then
Sheets(1).Cells(nrow1, "cd") = "false"
Sheets(9).Cells(18, "g") = TextBox8.Text
Sheets(9).Cells(18, "f") = TextBox9.Text
TextBox10.Text = Sheets(9).Cells(18, "h")
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mastmer_5dma_22
ElseIf CheckBox4.value = True And UserForm1.TextBox8.value < 8 _
And Sheets(1).Cells(nrow1, "b").value = TextBox2.Text And Sheets(1).Cells(nrow1, "e").value = TextBox5.Text Then
Sheets(1).Cells(nrow1, "cd") = "false"
Sheets(9).Cells(18, "g") = TextBox8.Text
Sheets(9).Cells(18, "f") = TextBox9.Text
TextBox10.Text = Sheets(9).Cells(18, "h")
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mastmer_5dma_22
End If
Next nrow1
'======================================== «Ã«“… Œ„” ”‰Ê« 
Case 2
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call five_snwat_22

'==================================================«Ã«“… »œÊ‰ —« »
Case 3
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call bdoon_ratb1_22
'===========================================«Ã«“… „⁄Ì·
Case 4
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mo3el_1_22
'=================================—›⁄ Ìœ «Ê «Ìﬁ«› ’—›
Case 5
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call raf3_yead_1_22
'=================================«„Ê„… «Ê· ”  «‘Â—
Case 6
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mother_hood_1_6_1_22
'=================================«„Ê„… À«‰Ì ”  «‘Â—
Case 7
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mother_hood_2_6_1_22
'================================= ‰”Ì»
Case 8
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call tanseeb_1_22
'=================================⁄ﬁœ
Case 9
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call contract_1_22
'=================================«Ã«“… ·Ã«‰ ÿ»Ì…

Case 10
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call lejan_tabea_1_22

'=================================„” „— »«·Œœ„Â „⁄  ÷„Ì‰…

Case 12
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mastmer_5dma_22




'=================================„” „— »«·Œœ„Â „⁄ ”·›…

Case 13
TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
Call mastmer_5dma_22


End Select

'==================================================«·’«›Ì





   '=========================================



End If
Next
TextBox102.value = Format(Sheets(26).Cells(3, "t"), "00.0")
If TextBox102.value = 0 Then
 Label122.Caption = "«·ﬁÌ„ „ ÿ«»ﬁ…"
 Else
 Label122.Caption = "«·ﬁÌ„ €Ì— „ ÿ«»ﬁ…"
End If
MsgBox " „ «·«Õ ”«» »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"

If CheckBox3.value = True Then

UserForm41.Label36 = TextBox5.Text
UserForm41.TextBox1.value = TextBox10.value
UserForm41.TextBox2.value = TextBox30.value
UserForm41.TextBox3.value = TextBox31.value
UserForm41.TextBox4.value = TextBox49.value
UserForm41.Show

End If
TextBox76.Text = Sheets(1).Cells(2, "cc")
TextBox73.Text = Sheets(1).Cells(2, "cd")
TextBox74.Text = Sheets(1).Cells(2, "ce")
TextBox79.Text = Sheets(1).Cells(2, "cf")
TextBox80.Text = Sheets(1).Cells(2, "cg")
TextBox82.Text = Sheets(1).Cells(1, "bn")
TextBox66.Text = Sheets(1).Cells(2, "ch")


'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")
'=====================================‰ﬁ· «·»Ì«‰«  «·Ï ‘Ì  46




End If






Call RecordToBlackBox(Me, DictOld)
End Sub

Private Sub CommandButton17_Click()
On Error GoTo emad
Application.ThisWorkbook.Save
ThisWorkbook.SaveCopyAs fileName:="d:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
ThisWorkbook.SaveCopyAs fileName:="e:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
MsgBox "·ﬁœ  „  ⁄„·Ì… «·‰”Œ «·«Œ Ì«ÿÌ »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
emad:
End Sub

Private Sub CommandButton18_Click()
On Error Resume Next

End Sub

Private Sub CommandButton19_Click()
On Error GoTo err_command30_click
Dim pathscaner As String
pathscaner = "c:\windows\system32\wiaacmgr.exe"
Call Shell(pathscaner, 1)
exit_command30_click:
Exit Sub
err_command30_click:
 MsgBox Err.Description
Resume exit_command30_click






End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
Sheets(1).Activate
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
For X = 9 To lastR
If Sheets(1).Cells(X, 1).value = TextBox1.value Or Sheets(1).Cells(X, 2) = TextBox2.Text Then
 MsgBox "«·—ﬁ„ „ﬂ—— ›Ì «·ﬁÌœ ·« Ì„ﬂ‰ «÷«› Â „—Â À«‰ÌÂ", vbOKOnly, "—”«·…  ÊÃÌÂ"
Exit Sub
End If
Next


'---------------------------------------
  TextBox14.Enabled = False
  TextBox16.Enabled = False
  TextBox18.Enabled = False
  TextBox20.Enabled = False
  TextBox24.Enabled = False
  TextBox27.Enabled = False
  TextBox29.Enabled = False
  TextBox30.Enabled = False
  TextBox31.Enabled = False
  TextBox32.Enabled = False
  TextBox33.Enabled = False
  TextBox47.Enabled = False
  TextBox48.Enabled = False
  TextBox49.Enabled = False
    
    
    Dim iRow As Long, i As Long

    Sheets(1).Activate

    iRow = Range("A" & Rows.count).End(xlUp).row
'=======================================================
  ' TextBox14.Text = 0                                              '
  ' TextBox16.Text = 0
  ' TextBox18.Text = 0                                                                       '
  '  TextBox20.Text = 0
  '  TextBox24.Text = 0
  '  TextBox27.Text = 0
  '  TextBox29.Text = 0
  '  TextBox30.Text = 0
  '  TextBox31.Text = 0
   ' TextBox32.Text = 0                  '
  '  TextBox33.Text = 0
   '  TextBox47.Text = 0
   '  TextBox48.Text = 0
  ' TextBox49.Text = 0
Sheets(1).Activate
'=========================================================
If TextBox53.Text = "" Or TextBox58.Text = "" Then
MsgBox "Â‰«ﬂ ‰ﬁ’ ›Ì «·»Ì«‰«  «·„œŒ·Â ﬁœ  ﬂÊ‰  «—ÌŒ «·«÷«›Â ·„  ﬁ„ »«Œ Ì«— «·‘Â— «·–Ì  „ «÷«› Â «Ê ·„  ﬁ„ »«Œ Ì«— «·„Êﬁ› «·ÊŸÌ›Ì «·„‰«”» ﬁ„ »«Œ Ì«—Â„« »‘ﬂ· ’ÕÌÕ ·Ì ”‰Ï ··‰Ÿ«„ «÷«›… «·ﬁÌœ ", vbOKOnly, "—”«·…  ÊÃÌÂ"

Exit Sub
Else
    Range("A" & iRow + 1).Offset(0, 0).value = TextBox1.value
    Range("A" & iRow + 1).Offset(0, 1).value = TextBox2.value
    Range("A" & iRow + 1).Offset(0, 2).value = TextBox3.value
    Range("A" & iRow + 1).Offset(0, 3).value = ComboBox4.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox5.value
    Range("A" & iRow + 1).Offset(0, 5).value = ComboBox5.value
    Range("A" & iRow + 1).Offset(0, 6).value = TextBox7.value
    Range("A" & iRow + 1).Offset(0, 7).value = TextBox8.value
    Range("A" & iRow + 1).Offset(0, 8).value = TextBox9.value
    Range("A" & iRow + 1).Offset(0, 9).value = TextBox10.value
    Range("A" & iRow + 1).Offset(0, 10).value = TextBox11.value
    Range("A" & iRow + 1).Offset(0, 11).value = TextBox12.value
    Range("A" & iRow + 1).Offset(0, 12).value = ComboBox6.value
   ' Range("A" & iRow + 1).Offset(0, 13).Value = TextBox14.Value
    Range("A" & iRow + 1).Offset(0, 14).value = ComboBox7.value
   ' Range("A" & iRow + 1).Offset(0, 15).Value = TextBox16.Value
    Range("A" & iRow + 1).Offset(0, 16).value = ComboBox8.value
   ' Range("A" & iRow + 1).Offset(0, 17).Value = TextBox18.Value
    Range("A" & iRow + 1).Offset(0, 18).value = TextBox19.value
  '  Range("A" & iRow + 1).Offset(0, 19).Value = TextBox20.Value
    Range("A" & iRow + 1).Offset(0, 20).value = TextBox21.value
    Range("A" & iRow + 1).Offset(0, 21).value = TextBox22.value
    Range("A" & iRow + 1).Offset(0, 22).value = TextBox23.value
  'Range("A" & iRow + 1).Offset(0, 23).Value = TextBox24.Value
    Range("A" & iRow + 1).Offset(0, 24).value = TextBox25.value
    Range("A" & iRow + 1).Offset(0, 25).value = ComboBox10.value
   ' Range("A" & iRow + 1).Offset(0, 26).Value = TextBox27.Value
   Range("A" & iRow + 1).Offset(0, 27).value = ComboBox11.value
  '  Range("A" & iRow + 1).Offset(0, 28).Value = TextBox29.Value
  '  Range("A" & iRow + 1).Offset(0, 29).Value = TextBox30.Value
   ' Range("A" & iRow + 1).Offset(0, 30).Value = TextBox31.Value
   ' Range("A" & iRow + 1).Offset(0, 31).Value = TextBox32.Value
   ' Range("A" & iRow + 1).Offset(0, 32).Value = TextBox33.Value
    Range("A" & iRow + 1).Offset(0, 33).value = TextBox34.value
    Range("A" & iRow + 1).Offset(0, 34).value = TextBox35.value
    Range("A" & iRow + 1).Offset(0, 35).value = TextBox36.value
    Range("A" & iRow + 1).Offset(0, 36).value = TextBox37.value
    Range("A" & iRow + 1).Offset(0, 37).value = TextBox38.value
    Range("A" & iRow + 1).Offset(0, 38).value = TextBox39.value
    Range("A" & iRow + 1).Offset(0, 39).value = TextBox40.value
    Range("A" & iRow + 1).Offset(0, 40).value = TextBox41.value
    Range("A" & iRow + 1).Offset(0, 41).value = TextBox42.value
    Range("A" & iRow + 1).Offset(0, 42).value = TextBox43.value
    Range("A" & iRow + 1).Offset(0, 43).value = TextBox44.value
    Range("A" & iRow + 1).Offset(0, 44).value = TextBox45.value
    Range("A" & iRow + 1).Offset(0, 45).value = TextBox46.value
   ' Range("A" & iRow + 1).Offset(0, 46).Value = TextBox47.Value
   ' Range("A" & iRow + 1).Offset(0, 47).Value = TextBox48.Value
   ' Range("A" & iRow + 1).Offset(0, 48).Value = TextBox49.Value
    Range("A" & iRow + 1).Offset(0, 49).value = TextBox50.value
    Range("A" & iRow + 1).Offset(0, 50).value = ComboBox9.value
  Range("A" & iRow + 1).Offset(0, 51).value = TextBox52.value
   Range("A" & iRow + 1).Offset(0, 52).value = TextBox53.value
' Range("A" & iRow + 1).Offset(0, 53).Value = TextBox54.Value
    Range("A" & iRow + 1).Offset(0, 54).value = TextBox55.value
    Range("A" & iRow + 1).Offset(0, 55).value = TextBox56.value
    Range("A" & iRow + 1).Offset(0, 56).value = TextBox57.value
   Range("A" & iRow + 1).Offset(0, 57).value = TextBox58.value
    Range("A" & iRow + 1).Offset(0, 58).value = TextBox59.value
    Range("A" & iRow + 1).Offset(0, 59).value = TextBox60.value
   ' Range("A" & iRow + 1).Offset(0, 64).Value = TextBox64.Value
   ' Range("A" & iRow + 1).Offset(0, 61).Value = TextBox62.Value
   ' Range("A" & iRow + 1).Offset(0, 62).Value = TextBox63.Value
   
   
        For i = 1 To 63
        Controls("TextBox" & i).value = ""
    Next i
MsgBox " „  «÷«›… «·»Ì«‰«  »‰Ã«Õ", vbOKOnly, "—”«·…  ÊÃÌÂ"
End If



End Sub



Private Sub CommandButton20_Click()
On Error GoTo emad








'=======================================================
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(1).Cells(Y, 1) = TextBox1.Text
    Sheets(1).Cells(Y, 2) = TextBox2.Text
    Sheets(1).Cells(Y, 3) = TextBox3.Text
    Sheets(1).Cells(Y, 4) = ComboBox4.value
    Sheets(1).Cells(Y, 5) = TextBox5.Text
    Sheets(1).Cells(Y, 6) = ComboBox5.value
    Sheets(1).Cells(Y, 7) = TextBox7.Text
    Sheets(1).Cells(Y, 8) = TextBox8.Text
    Sheets(1).Cells(Y, 9) = TextBox9.Text
    Sheets(1).Cells(Y, 10) = TextBox10.Text
    Sheets(1).Cells(Y, 11) = TextBox11.Text
    Sheets(1).Cells(Y, 12) = TextBox12.Text
     Sheets(1).Cells(Y, 13) = ComboBox6.value
    Sheets(1).Cells(Y, 14) = TextBox14.Text
    Sheets(1).Cells(Y, 15) = ComboBox7.value
    Sheets(1).Cells(Y, 16) = TextBox16.Text
    Sheets(1).Cells(Y, 17) = ComboBox8.value
    Sheets(1).Cells(Y, 18) = TextBox18.Text
    Sheets(1).Cells(Y, 19) = TextBox19.Text
    Sheets(1).Cells(Y, 20) = TextBox20.Text
    Sheets(1).Cells(Y, 21) = TextBox21.Text
    Sheets(1).Cells(Y, 22) = TextBox22.Text
    Sheets(1).Cells(Y, 23) = TextBox23.Text
    Sheets(1).Cells(Y, 24) = TextBox24.Text
    Sheets(1).Cells(Y, 25) = TextBox25.Text
    Sheets(1).Cells(Y, 26) = ComboBox10.value
    Sheets(1).Cells(Y, 27) = TextBox27.Text
    Sheets(1).Cells(Y, 28) = ComboBox11.value
  Sheets(1).Cells(Y, 29) = TextBox29.Text
    Sheets(1).Cells(Y, 30) = TextBox30.Text
    Sheets(1).Cells(Y, 31) = TextBox31.Text
  Sheets(1).Cells(Y, 32) = TextBox32.Text
   Sheets(1).Cells(Y, 33) = TextBox33.Text
    Sheets(1).Cells(Y, 34) = TextBox34.Text
    Sheets(1).Cells(Y, 35) = TextBox35.Text
    Sheets(1).Cells(Y, 36) = TextBox36.Text
    Sheets(1).Cells(Y, 37) = TextBox37.Text
    Sheets(1).Cells(Y, 38) = TextBox38.Text
    Sheets(1).Cells(Y, 39) = TextBox39.Text
    Sheets(1).Cells(Y, 40) = TextBox40.Text
    Sheets(1).Cells(Y, 41) = TextBox41.Text
    Sheets(1).Cells(Y, 42) = TextBox42.Text
    Sheets(1).Cells(Y, 43) = TextBox43.Text
    Sheets(1).Cells(Y, 44) = TextBox44.Text
    Sheets(1).Cells(Y, 45) = TextBox45.Text
    Sheets(1).Cells(Y, 46) = TextBox46.Text
    Sheets(1).Cells(Y, 47) = TextBox47.Text
  Sheets(1).Cells(Y, 48) = TextBox48.Text
   Sheets(1).Cells(Y, 49) = TextBox49.Text
   Sheets(1).Cells(Y, 50) = TextBox50.Text
    Sheets(1).Cells(Y, 51) = ComboBox9.value
    Sheets(1).Cells(Y, 52) = TextBox52.Text
  Sheets(1).Cells(Y, 53) = TextBox53.Text
   Sheets(1).Cells(Y, 54) = TextBox54.Text
    Sheets(1).Cells(Y, 55) = TextBox55.Text
    Sheets(1).Cells(Y, 56) = TextBox56.Text
    Sheets(1).Cells(Y, 57) = TextBox57.Text
  Sheets(1).Cells(Y, 58) = TextBox58.Text
   Sheets(1).Cells(Y, 59) = TextBox59.Text
    Sheets(1).Cells(Y, 60) = TextBox60.Text
    Sheets(1).Cells(Y, 61) = TextBox61.Text
  Sheets(1).Cells(Y, 62) = TextBox62.Text
   Sheets(1).Cells(Y, 63) = TextBox63.Text
  Sheets(1).Cells(Y, 64) = TextBox64.Text
    
    
    Sheets(1).Cells(Y, 66) = Format(TextBox82.Text, "yyyy/mm/dd")
  Sheets(1).Cells(Y, 67) = TextBox83.Text
   
   Sheets(1).Cells(Y, 68) = TextBox84.Text
 Sheets(1).Cells(Y, 69) = TextBox85.Text
  Sheets(1).Cells(Y, 70) = TextBox86.Text
   Sheets(1).Cells(Y, 71) = TextBox87.Text
   Sheets(1).Cells(Y, 72) = TextBox88.Text
   Sheets(1).Cells(Y, 73) = TextBox64.Text
   Sheets(1).Cells(Y, 74) = TextBox91.Text
    Sheets(1).Cells(Y, "bu") = TextBox99.Text
    Sheets(1).Cells(Y, "ca") = TextBox100.Text
    
   Call emad_pro
     MsgBox " „  ⁄„·Ì…  ÕœÌÀ «·»Ì«‰«  »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
emad:
End Sub

Private Sub CommandButton21_Click()
On Error GoTo emad
Dim file_dialog As Office.FileDialog
Dim file_path As String
Set file_dialog = Application.FileDialog(msoFileDialogFilePicker)
 file_dialog.title = "«Œ — ’Ê—…"
With file_dialog
.title = "«Œ — ’Ê—…"

.Filters.Clear
.Filters.Add "Image", "*.gif,*.jpg,*.*jpeg"
.AllowMultiSelect = False

If file_dialog.Show = 0 Then Exit Sub
 file_path = file_dialog.SelectedItems(1)
 Me.Image1.Picture = LoadPicture(file_path)
 
 TextBox64.Text = file_path
End With




emad:

End Sub

Private Sub CommandButton22_Click()
Sheets(1).Activate
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row

For X = 9 To lastR
If Cells(X, 5) = TextBox5.Text And Cells(X, 2) = TextBox2.Text Then
Cells(X, 64).Delete
Cells(X, 64).value = ""
MsgBox " „ Õ–› «·’Ê—… »‰Ã«Õ", vbCritical + vbOKOnly, "—”«·…  ‰»ÌÂ"
If Cells(X, 64).value = "" Then
Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
TextBox64.Text = "D:\employ_pic\shdow_pic.jpg"
Cells(X, 64).value = TextBox64.Text
SavePicture Image1.Picture, "Sheets.Cells(x, 64).Value"
End If
Exit For
End If
Next X




End Sub

Private Sub CommandButton23_Click()
Application.DisplayAlerts = False
ThisWorkbook.Save
MsgBox "·ﬁœ  „ Õ›Ÿ «·»Ì«‰«  «· Ì ﬁ„  »«œŒ«·Â«", vbOKOnly, "—”«·…  ‰»ÌÂ"

Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton24_Click()
On Error GoTo emad
Dim lastR As Integer
Dim file_path As String
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text And Sheets(1).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    
  
   
   'Call emad_pro

 '=========================================
 file_path = "D:\employ_pic\" & TextBox5.Text & ".jpg"
 
 TextBox64.Text = file_path
 SavePicture Image1.Picture, TextBox64.Text

Sheets(1).Cells(Y, 64) = TextBox64.Text
MsgBox " „ Õ›Ÿ «·’Ê—…"
emad:

End Sub



Private Sub CommandButton25_Click()
Call automotive
End Sub


Private Sub CommandButton27_Click()
UserForm21.Show
End Sub

Private Sub CommandButton28_Click()
UserForm22.Show
End Sub

Private Sub CommandButton29_Click()
UserForm23.Show
End Sub

Private Sub CommandButton3_Click()
On Error GoTo emad
Unload Me
UserForm65.Show
emad:
End Sub

Private Sub CommandButton30_Click()
UserForm24.Show
End Sub

Private Sub CommandButton31_Click()

' ""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—""
'=================================================== ÕœÌÀ
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text And Sheets(1).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
  '  Sheets(1).Cells(y, 1) = TextBox1.Text
  '  Sheets(1).Cells(y, 2) = TextBox2.Text
 '   Sheets(1).Cells(y, 3) = TextBox3.Text
 '   Sheets(1).Cells(y, 4) = ComboBox4.Value
  '  Sheets(1).Cells(y, 5) = TextBox5.Text
 '   Sheets(1).Cells(y, 6) = ComboBox5.Value
 '   Sheets(1).Cells(y, 7) = TextBox7.Text
 '   Sheets(1).Cells(y, 8) = TextBox8.Text
 '   Sheets(1).Cells(y, 9) = TextBox9.Text
 '   Sheets(1).Cells(y, 10) = TextBox10.Text
 '   Sheets(1).Cells(y, 11) = TextBox11.Text
  '  Sheets(1).Cells(y, 12) = TextBox12.Text
 '    Sheets(1).Cells(y, 13) = ComboBox6.Value
 '   Sheets(1).Cells(y, 14) = TextBox14.Text
'    Sheets(1).Cells(y, 15) = ComboBox7.Value
'    Sheets(1).Cells(y, 16) = TextBox16.Text
'    Sheets(1).Cells(y, 17) = ComboBox8.Value
'    Sheets(1).Cells(y, 18) = TextBox18.Text
'    Sheets(1).Cells(y, 19) = TextBox19.Text
'    Sheets(1).Cells(y, 20) = TextBox20.Text
    Sheets(1).Cells(Y, 21) = TextBox21.Text
    Sheets(1).Cells(Y, 22) = TextBox22.Text
   
'    Sheets(1).Cells(y, 24) = TextBox24.Text
 '   Sheets(1).Cells(y, 25) = TextBox25.Text
'    Sheets(1).Cells(y, 26) = ComboBox10.Value
 '   Sheets(1).Cells(y, 27) = TextBox27.Text
 '   Sheets(1).Cells(y, 28) = ComboBox11.Value
'  Sheets(1).Cells(y, 29) = TextBox29.Text
'    Sheets(1).Cells(y, 30) = TextBox30.Text
 '   Sheets(1).Cells(y, 31) = TextBox31.Text
'  Sheets(1).Cells(y, 32) = TextBox32.Text
'   Sheets(1).Cells(y, 33) = TextBox33.Text
'    Sheets(1).Cells(y, 34) = TextBox34.Text
 '   Sheets(1).Cells(y, 35) = TextBox35.Text
 '   Sheets(1).Cells(y, 36) = TextBox36.Text
 '   Sheets(1).Cells(y, 37) = TextBox37.Text
'    Sheets(1).Cells(y, 38) = TextBox38.Text
'    Sheets(1).Cells(y, 39) = TextBox39.Text
'    Sheets(1).Cells(y, 40) = TextBox40.Text
'    Sheets(1).Cells(y, 41) = TextBox41.Text
'    Sheets(1).Cells(y, 42) = TextBox42.Text
'    Sheets(1).Cells(y, 43) = TextBox43.Text
'    Sheets(1).Cells(y, 44) = TextBox44.Text
'    Sheets(1).Cells(y, 45) = TextBox45.Text
'    Sheets(1).Cells(y, 46) = TextBox46.Text
'    Sheets(1).Cells(y, 47) = TextBox47.Text
'  Sheets(1).Cells(y, 48) = TextBox48.Text
 '  Sheets(1).Cells(y, 49) = TextBox49.Text
 '  Sheets(1).Cells(y, 50) = TextBox50.Text
 '   Sheets(1).Cells(y, 51) = ComboBox9.Value
 '   Sheets(1).Cells(y, 52) = TextBox52.Text
'  Sheets(1).Cells(y, 53) = TextBox53.Text
'   Sheets(1).Cells(y, 54) = TextBox54.Text
'    Sheets(1).Cells(y, 55) = TextBox55.Text
'    Sheets(1).Cells(y, 56) = TextBox56.Text
'    Sheets(1).Cells(y, 57) = TextBox57.Text
'  Sheets(1).Cells(y, 58) = TextBox58.Text
'   Sheets(1).Cells(y, 59) = TextBox59.Text
'    Sheets(1).Cells(y, 60) = TextBox60.Text
'    Sheets(1).Cells(y, 61) = TextBox61.Text
'  Sheets(1).Cells(y, 62) = TextBox62.Text
'   Sheets(1).Cells(y, 63) = TextBox63.Text
'  Sheets(1).Cells(y, 64) = TextBox64.Text
 '   Sheets(1).Cells(y, 66) = Date
    
'    Sheets(1).Cells(y, 66) = TextBox82.Text
'  Sheets(1).Cells(y, 67) = TextBox83.Text
 '  Sheets(1).Cells(y, 68) = TextBox84.Text
'  Sheets(1).Cells(y, 69) = TextBox85.Text
'    Sheets(1).Cells(y, 70) = TextBox86.Text
'  Sheets(1).Cells(y, 71) = TextBox87.Text





'===========================================
Ê—ﬁ…1.Range("x9").Formula2R1C1 = "=iferror(IF(OR(RC[-1]= thariaba1!R1C20:R2C20),VLOOKUP(RC[-14],astktat1,RC[-2]+3,1),IF(OR(RC[-1]=thariaba2!R1C19),VLOOKUP(RC[-14],astktaat2,RC[-2]+3,1),IF(OR(RC[-1]=thariaba3!R1C19:R3C19),VLOOKUP(RC[-14],astktaat3,RC[-2]+3,1),IF(AND(VLOOKUP(RC[-1],sheet1!R1C77:R5C78,2,0)<4,RC[-2]>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(RC[-14],astktaat,VLOOKUP(RC[-1],sheet1!R1C77:R5C78,2,0)+RC[-2],1))))),0)"

Ê—ﬁ…1.Range("x9:x8000").FillDown
 Sheets(1).Cells(Y, 23) = TextBox23.Text
'=========================================
UserForm53.Show
'Dim lastr1, y1 As Integer
'lastr1 = Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
  '  For y1 = 9 To lastr1
     '   If Sheets(1).Cells(y1, "x") = "Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—" Then
 '  k = 0
 ' UserForm53.ListBox1.AddItem
  '     UserForm53.ListBox1.List(k, 0) = Sheets(1).Cells(y1, "e")
       
    '    k = k + 1
   
   
   
  ' MsgBox "Â‰«ﬂ Œÿ√ ›Ì «Õ ”«» „‰ŸÊ„… «·«” ﬁÿ«⁄ «·÷—Ì»Ì Õ«Ê· «’·«Õ… ﬁœ ÌﬂÊ‰ ›Ì «·ﬁÌœ " & Sheets(1).Cells(y, "e"), vbCritical, "—”«·…  ‰»ÌÂ"

'End If
'Next


'MsgBox " „  ⁄„Ì„ «·Ÿ—Ì»Â »‰Ã«Õ", vbDefaultButton1
End Sub

Private Sub CommandButton32_Click()
'============================== ‰ﬁ· „‰ ‘Ì  Œ«—ÃÌ «·Ï ﬁ«⁄œ… «·»Ì«‰«  «·Õ«·ÌÂ »ﬂ· «· ›«’Ì·

On Error Resume Next
Sheets(1).Activate
Call DeleteDataToColumn84


Application.Visible = True
Call open_file
Application.Visible = True
Windows("rwtab_backup.xlsm").Activate
Sheets(1).Select
ActiveSheet.Range("b9:br8000").Select
Selection.Copy
ThisWorkbook.Activate
Sheets(1).Select
Range("b9:br8000").Select
Selection.PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
Windows("rwtab_backup.xlsm").Close
Application.Visible = False

Sheets(1).Range("a9").Formula = "=IF(B9="""","""", SUBTOTAL(3,$B$9:B9))"
Sheets(1).Range("a9:a8000").FillDown


Sheets(13).Activate

Range("b9:br8000").ClearContents

Application.ScreenUpdating = False
Call copy_from_to_no_userform1
Application.ScreenUpdating = True

MsgBox "·ﬁœ  „ ‰ﬁ· Ê«” Ì—«œ «·»Ì«‰«  «·Œ«—ÃÌÂ »‰Ã«Õ", vbCritical, "—”«·…  ‰»ÌÂ"
End Sub

Private Sub CommandButton33_Click()
UserForm26.Show
End Sub

Private Sub CommandButton34_Click()
UserForm67.Show


End Sub



Private Sub CommandButton36_Click()
On Error Resume Next
'     ...... «·Õ«·… «·—«»⁄… ··÷—Ì»Â
'  «·«⁄“»
'  «·„ÿ·ﬁ Ê·Ì” ·œÌÂ Õ÷«‰… ··√Ê·«œ
'  «·„ÊŸ›Â «·„›’Ê· œŒ·Â« ⁄‰ œŒ· “ÊÃÂ«
'   «·„ÊŸ› «·„ “ÊÃ Ê “ÊÃ Â —»… »Ì 
'  «·„ÊŸ›Â «·„ “ÊÃÂ Ê“ÊÃÂ« ⁄«ÿ· ⁄„‰ «·⁄„·


If TextBox23.Text = "«·«⁄“»" Or TextBox23.Text = "«·„ÿ·ﬁ Ê·Ì” ·œÌÂ Õ÷«‰… ··√Ê·«œ" Or TextBox23.Text = "«·„ÊŸ›Â «·„›’Ê· œŒ·Â« ⁄‰ œŒ· “ÊÃÂ«" Or TextBox23.Text = "«·„ÊŸ› «·„ “ÊÃ Ê “ÊÃ Â —»… »Ì " Or TextBox23.Text = "«·„ÊŸ›Â «·„ “ÊÃÂ Ê“ÊÃÂ« ⁄«ÿ· ⁄„‰ «·⁄„·" Then
Dim textvlue1 As Long
Dim textvlue2 As Long
Dim rang1 As Range

Dim rang2 As Range
Set rang1 = Sheets(1).Range("zwjea4")
Set rang2 = Sheets(2).Range("astktaat")

textvlue1 = TextBox22.Text
textvlue2 = TextBox10.Text


vlookup1 = Application.WorksheetFunction.VLookup(TextBox23.Text, rang1, 2, 0) < 4 And TextBox22.value > 0
'TextBox27.Value = vlookup1


If vlookup1 = -1 Then
TextBox22.value = 0
MsgBox "Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—", vbOKCancel, "—”«·…  ÊÃÌÂ"
Exit Sub
Else
TextBox24.value = Application.WorksheetFunction.VLookup(textvlue2, rang2, Application.WorksheetFunction.VLookup(TextBox23.Text, rang1, 2, 0) + textvlue1, 1)
End If
End If






'======================================================
'  «·„ÊŸ›… „ÿ·ﬁ…  / «·„ÊŸ›… «—„·… ...... «·Õ«·… «·«Ê·Ï ··÷—Ì»Â
On Error Resume Next
If TextBox23.Text = "«·„ÊŸ›… „ÿ·ﬁ…" Or TextBox23.Text = "«·„ÊŸ›… «—„·…" Then
Dim textvlue3 As Long
Dim textvlue4 As Long
textvlue3 = TextBox22.Text
textvlue4 = TextBox10.Text
Dim rang3 As Range
Dim rang4 As Range
Set rang3 = Sheets("thariaba1").Range("zwjea1")
Set rang4 = Sheets("thariaba1").Range("astktat1")

If TextBox23.Text = rang3 Then
TextBox24.value = Application.WorksheetFunction.VLookup(textvlue4, rang4, textvlue3 + 3, 1)
End If
End If
'========================================================
'  «·„ÊŸ›… „ “ÊÃ… Ê“ÊÃÂ« ⁄«Ã“ ⁄‰ «·⁄„·/ ......«·Õ«·… «·À«‰Ì… ··÷—Ì»Â
On Error Resume Next
If TextBox23.Text = "«·„ÊŸ›… „ “ÊÃ… Ê“ÊÃÂ« ⁄«Ã“ ⁄‰ «·⁄„·" Then
Dim textvlue5 As Long
Dim textvlue6 As Long
textvlue5 = TextBox22.Text
textvlue6 = TextBox10.Text
Dim rang5 As Range
Dim rang6 As Range
Set rang5 = Sheets("thariaba2").Range("zwjea2")
Set rang6 = Sheets("thariaba2").Range("astktaat2")

If TextBox23.Text = rang5 Then
TextBox24.value = Application.WorksheetFunction.VLookup(textvlue6, rang6, textvlue5 + 3, 1)
End If
End If

'=========================================================
' «·„ÊŸ› «·„ÿ·ﬁ «–« ﬂ«‰ Õ«÷‰« ··√Ê·«œ/ „ Ê›«… “ÊÃ Â / «·„ÊŸ› „ “ÊÃ ÊœŒ·Â „›’Ê· ⁄‰ œŒ· “ÊÃ Â ......«·Õ«·… «·À«·À… ··÷—Ì»Â
On Error Resume Next
If TextBox23.Text = "«·„ÊŸ› «·„ÿ·ﬁ «–« ﬂ«‰ Õ«÷‰« ··√Ê·«œ" Or TextBox23.Text = "„ Ê›«… “ÊÃ Â" Or TextBox23.Text = "«·„ÊŸ› „ “ÊÃ ÊœŒ·Â „›’Ê· ⁄‰ œŒ· “ÊÃ Â" Then
Dim textvlue7 As Long
Dim textvlue8 As Long
textvlue7 = TextBox22.Text
textvlue8 = TextBox10.Text
Dim rang7 As Range
Dim rang8 As Range
Set rang7 = Sheets("thariaba3").Range("zwjea3")
Set rang8 = Sheets("thariaba3").Range("astktaat3")

If TextBox23.Text = rang7 Then
TextBox24.value = Application.WorksheetFunction.VLookup(textvlue8, rang8, textvlue7 + 3, 1)
End If

End If

End Sub

Private Sub CommandButton35_Click()
Sheets(38).Activate
If UserForm1.TextBox58.Text = "—« » Ã“∆Ì" And UserForm1.ComboBox3.value = "—« » Ã“∆Ì" Then
   
 '   iRow = Range("A" & Rows.Count).End(xlUp).Row
 '   Range("A" & iRow + 1).Offset(0, 0) = TextBox2.Text
  '  Range("A" & iRow + 1).Offset(0, 1) = TextBox5.Text
 '   Range("A" & iRow + 1).Offset(0, 2) = TextBox10.Text
 '   Range("A" & iRow + 1).Offset(0, 3) = TextBox49.Text
  '  Range("A" & iRow + 1).Offset(0, 4) = TextBox58.Text
 '   Range("A" & iRow + 1).Offset(0, 5) = Format(Sheets(38).Cells(1, "l"), "yyyy/mm/dd")
        
   
   
   
  lastR = Sheets(38).Cells(Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(38).Cells(Y, 1) = TextBox2.Text And Sheets(38).Cells(Y, 2) = TextBox5.Text Then
   Cells(Y, 1).Select
ActiveCell.EntireRow.Delete
    Exit For
    End If
    Next Y
    
    
    
    
  
  '  Sheets(38).Cells(y, 1) = ""
  '  Sheets(38).Cells(y, 2) = ""
  '  Sheets(38).Cells(y, 3) = ""
   ' Sheets(38).Cells(y, 4) = ""
  '  Sheets(38).Cells(y, 5) = ""
  '  Sheets(38).Cells(y, 6) = ""
   
   
   
   
   
    lastr1 = Sheets(38).Cells(Rows.count, 1).End(xlUp).row
    For y1 = 2 To lastr1
      If Sheets(38).Cells(y1, 1) <> TextBox2.Text And Sheets(38).Cells(y1, 2) <> TextBox5.Text Then
    
    Exit For
    End If
     Next
    iRow = Range("A" & Rows.count).End(xlUp).row
    Range("A" & iRow + 1).Offset(0, 0) = TextBox2.Text
    Range("A" & iRow + 1).Offset(0, 1) = TextBox5.Text
    Range("A" & iRow + 1).Offset(0, 2) = TextBox10.Text
    Range("A" & iRow + 1).Offset(0, 3) = TextBox49.Text
    Range("A" & iRow + 1).Offset(0, 4) = TextBox58.Text
    Range("A" & iRow + 1).Offset(0, 5) = Format(Sheets(38).Cells(1, "l"), "yyyy/mm/dd")
        
   
   
 MsgBox "ﬁœ  „  ”ÃÌ· »Ì«‰« Â ··„œÂ «·«Ê·Ï »«„ﬂ«‰ﬂ «·–Â«» · ”ÃÌ· »Ì«‰« Â ··„œÂ «·À«‰ÌÂ", vbCritical, "—”«·…  ‰»ÌÂ"


   
 'UserForm29.Show

Else
MsgBox "«‰ »Â ⁄·Ìﬂ «Ê·« »«Œ Ì«— «·„Êﬁ› «·ÊŸÌ›Ì ·Â–« «·„‰ ”» »’›… —« » Ã“∆Ì ·Ì ”‰Ï ··‰Ÿ«„ ﬁ—«∆… «·› —Â «·«Ê·Ï Ê„‰ »⁄œÂ« ⁄«Êœ «⁄ÿ«¡ ’›Â Ê„Êﬁ› ÊŸÌ›Ì «·„·«∆„ ·Â", vbCritical, "—”«·…  ‰»ÌÂ"

End If

   

End Sub

Private Sub CommandButton37_Click()
On Error Resume Next
'=IF(AND(BF9="Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« "),J9-(AF9+AG9)&"  ""·œÌÂ «Ã«“Â Œ„” ”‰Ê«  Ê«·„ »ﬁÌ ’«›Ì «·—« » ",IF(AND(BF9=" «Ã«“… »œÊ‰ —« »"),0  &""&"  ·œÌÂ «Ã«“Â »œÊ‰ —« »  ",AE9-AV9))
' «·ŒÊ«—“„ÌÂ  »œ∆
'1 «Õ ”«»  œ«·… «·÷—Ì»Â
'2 «Õ ”«» «·«” Õﬁ«ﬁ« 
'3 «Õ ”«» «·«” ﬁÿ«⁄« 
'4 «Õ ”«» «·’«›Ì
'5 «Õ ”«» ‰”»… «·€Ì«»
'6 «Õ ”«» œ«·… «·„Êﬁ› «·Ê÷Ì›Ì
'7
'8


'======================================== «·«” Õﬁ«ﬁ« 




Select Case TextBox100.value
'==================================================„” „— »«·Œœ„…
Case 1
Call mastmer_5dma_22

'======================================== «Ã«“… Œ„” ”‰Ê« 
Case 2
Call five_snwat_22


'==================================================«Ã«“… »œÊ‰ —« »
Case 3
Call bdoon_ratb1_22
'===========================================«Ã«“… „⁄Ì·
Case 4
Call mo3el_1_22
'=================================—›⁄ Ìœ «Ê «Ìﬁ«› ’—›
Case 5
Call raf3_yead_1_22
'=================================«„Ê„… «Ê· ”  «‘Â—
Case 6
Call mother_hood_1_6_1_22
'=================================«„Ê„… À«‰Ì ”  «‘Â—
Case 7
Call mother_hood_2_6_1_22
'================================= ‰”Ì»
Case 8
Call tanseeb_1_22
'=================================⁄ﬁœ
Case 9
Call contract_1_22
'=================================«Ã«“… ·Ã«‰ ÿ»Ì…

Case 10
Call lejan_tabea_1_22

'=================================„” „— »«·Œœ„Â „⁄  ÷„Ì‰…

Case 12
Call mastmer_5dma_22




'=================================„” „— »«·Œœ„Â „⁄ ”·›…

Case 13
Call mastmer_5dma_22


End Select

'==================================================«·’«›Ì






'=======================================================
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(1).Cells(Y, 1) = TextBox1.Text
    Sheets(1).Cells(Y, 2) = TextBox2.Text
    Sheets(1).Cells(Y, 3) = TextBox3.Text
    Sheets(1).Cells(Y, 4) = ComboBox4.value
    Sheets(1).Cells(Y, 5) = TextBox5.Text
    Sheets(1).Cells(Y, 6) = ComboBox5.value
    Sheets(1).Cells(Y, 7) = TextBox7.Text
    Sheets(1).Cells(Y, 8) = TextBox8.Text
    Sheets(1).Cells(Y, 9) = TextBox9.Text
    Sheets(1).Cells(Y, 10) = TextBox10.Text
    Sheets(1).Cells(Y, 11) = TextBox11.Text
    Sheets(1).Cells(Y, 12) = TextBox12.Text
     Sheets(1).Cells(Y, 13) = ComboBox6.value
    Sheets(1).Cells(Y, 14) = TextBox14.Text
    Sheets(1).Cells(Y, 15) = ComboBox7.value
    Sheets(1).Cells(Y, 16) = TextBox16.Text
    Sheets(1).Cells(Y, 17) = ComboBox8.value
    Sheets(1).Cells(Y, 18) = TextBox18.Text
    Sheets(1).Cells(Y, 19) = TextBox19.Text
    Sheets(1).Cells(Y, 20) = TextBox20.Text
    Sheets(1).Cells(Y, 21) = TextBox21.Text
    Sheets(1).Cells(Y, 22) = TextBox22.Text
    Sheets(1).Cells(Y, 23) = TextBox23.Text
    Sheets(1).Cells(Y, 24) = TextBox24.Text
    Sheets(1).Cells(Y, 25) = TextBox25.Text
    Sheets(1).Cells(Y, 26) = ComboBox10.value
    Sheets(1).Cells(Y, 27) = TextBox27.Text
    Sheets(1).Cells(Y, 28) = ComboBox11.value
  Sheets(1).Cells(Y, 29) = TextBox29.Text
    Sheets(1).Cells(Y, 30) = TextBox30.Text
    Sheets(1).Cells(Y, 31) = TextBox31.Text
  Sheets(1).Cells(Y, 32) = TextBox32.Text
   Sheets(1).Cells(Y, 33) = TextBox33.Text
    Sheets(1).Cells(Y, 34) = TextBox34.Text
    Sheets(1).Cells(Y, 35) = TextBox35.Text
    Sheets(1).Cells(Y, 36) = TextBox36.Text
    Sheets(1).Cells(Y, 37) = TextBox37.Text
    Sheets(1).Cells(Y, 38) = TextBox38.Text
    Sheets(1).Cells(Y, 39) = TextBox39.Text
    Sheets(1).Cells(Y, 40) = TextBox40.Text
    Sheets(1).Cells(Y, 41) = TextBox41.Text
    Sheets(1).Cells(Y, 42) = TextBox42.Text
    Sheets(1).Cells(Y, 43) = TextBox43.Text
    Sheets(1).Cells(Y, 44) = TextBox44.Text
    Sheets(1).Cells(Y, 45) = TextBox45.Text
    Sheets(1).Cells(Y, 46) = TextBox46.Text
    Sheets(1).Cells(Y, 47) = TextBox47.Text
  Sheets(1).Cells(Y, 48) = TextBox48.Text
   Sheets(1).Cells(Y, 49) = TextBox49.Text
   Sheets(1).Cells(Y, 50) = TextBox50.Text
    Sheets(1).Cells(Y, 51) = ComboBox9.value
    Sheets(1).Cells(Y, 52) = TextBox52.Text
  Sheets(1).Cells(Y, 53) = TextBox53.Text
   Sheets(1).Cells(Y, 54) = TextBox54.Text
    Sheets(1).Cells(Y, 55) = TextBox55.Text
    Sheets(1).Cells(Y, 56) = TextBox56.Text
    Sheets(1).Cells(Y, 57) = TextBox57.Text
  Sheets(1).Cells(Y, 58) = TextBox58.Text
   Sheets(1).Cells(Y, 59) = TextBox59.Text
    Sheets(1).Cells(Y, 60) = TextBox60.Text
    Sheets(1).Cells(Y, 61) = TextBox61.Text
  Sheets(1).Cells(Y, 62) = TextBox62.Text
   Sheets(1).Cells(Y, 63) = TextBox63.Text
   Sheets(1).Cells(Y, 64) = TextBox64.Text
      TextBox82.Text = Format(TextBox92.Text, "yyyy/mm/dd")
    Sheets(1).Cells(Y, 66) = Format(TextBox82.Text, "yyyy/mm/dd")
  '  Sheets(1).Cells(y, 68) = TextBox84.Text
 ' Sheets(1).Cells(y, 69) = TextBox85.Text
  ' Sheets(1).Cells(y, 70) = TextBox86.Text
   Sheets(1).Cells(Y, 71) = TextBox87.Text
   Sheets(1).Cells(Y, 72) = TextBox88.Text
   'Sheets(1).Cells(y, 73) = TextBox64.Text
   Sheets(1).Cells(Y, 74) = TextBox91.Text
   
   
   Call emad_pro
         ' For i = 1 To 49
       ' Me.Controls("TextBox" & i).Text = ""
   ' Next i
   '=============================================
 


   
UserForm29.Show
   
End Sub

Private Sub CommandButton38_Click()
UserForm27.Show
End Sub





Private Sub CommandButton39_Click()
Dim rng As Range
Dim cel As Range
Dim lastr1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, "b").End(xlUp).row
Set rng = Sheets(1).Range("au9:au" & lastr1)
For Each cel In rng
If IsEmpty(cel) Or Not IsNumeric(cel) Then
cel = 0
End If
Next cel
End Sub

Private Sub CommandButton4_Click()
'========================================================
'Data migration
 TextBox14.Enabled = False
  TextBox16.Enabled = False
  TextBox18.Enabled = False
  TextBox20.Enabled = False
  TextBox24.Enabled = False
  TextBox27.Enabled = False
  TextBox29.Enabled = False
  TextBox30.Enabled = False
  TextBox31.Enabled = False
  TextBox32.Enabled = False
  TextBox33.Enabled = False
  TextBox47.Enabled = False
  TextBox48.Enabled = False
  TextBox49.Enabled = True
        
   
    Dim lastr1 As Integer
'lastr1 = Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
   ' For y = 9 To lastr1
     If TextBox56.Text = "" Or TextBox57.Text = "" Or TextBox58.Text = "" Then
    MsgBox "„‰ ›÷·ﬂ «œŒ· «·ÃÂ… «·„‰”Ê» «·ÌÂ« Ê «—ÌŒ  —ÕÌ·Â Ê„Êﬁ›Â «·ÊŸÌ›Ì", vbCritical, "—”«·…  ‰»ÌÂ"
     
        Exit Sub
       Else
 
  
        
        
        
        
    Dim iRow As Long, i As Long

    Sheets(3).Activate

    iRow = Range("A" & Rows.count).End(xlUp).row
'=======================================================
  ' TextBox14.Text = 0                                              '
  ' TextBox16.Text = 0
  ' TextBox18.Text = 0                                                                       '
  '  TextBox20.Text = 0
  '  TextBox24.Text = 0
  '  TextBox27.Text = 0
  '  TextBox29.Text = 0
  '  TextBox30.Text = 0
  '  TextBox31.Text = 0
   ' TextBox32.Text = 0                  '
  '  TextBox33.Text = 0
   '  TextBox47.Text = 0
   '  TextBox48.Text = 0
  ' TextBox49.Text = 0



'=========================================================
    Range("A" & iRow + 1).Offset(0, 0).value = iRow
    Range("A" & iRow + 1).Offset(0, 1).value = TextBox2.value
    Range("A" & iRow + 1).Offset(0, 2).value = TextBox3.value
    Range("A" & iRow + 1).Offset(0, 3).value = ComboBox4.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox5.value
    Range("A" & iRow + 1).Offset(0, 5).value = ComboBox5.value
    Range("A" & iRow + 1).Offset(0, 6).value = TextBox7.value
    Range("A" & iRow + 1).Offset(0, 7).value = TextBox8.value
    Range("A" & iRow + 1).Offset(0, 8).value = TextBox9.value
    Range("A" & iRow + 1).Offset(0, 9).value = TextBox10.value
    Range("A" & iRow + 1).Offset(0, 10).value = TextBox11.value
    Range("A" & iRow + 1).Offset(0, 11).value = TextBox12.value
    Range("A" & iRow + 1).Offset(0, 12).value = ComboBox6.value
    Range("A" & iRow + 1).Offset(0, 13).value = TextBox14.value
    Range("A" & iRow + 1).Offset(0, 14).value = ComboBox7.value
    Range("A" & iRow + 1).Offset(0, 15).value = TextBox16.value
    Range("A" & iRow + 1).Offset(0, 16).value = ComboBox8.value
    Range("A" & iRow + 1).Offset(0, 17).value = TextBox18.value
    Range("A" & iRow + 1).Offset(0, 18).value = TextBox19.value
    Range("A" & iRow + 1).Offset(0, 19).value = TextBox20.value
    Range("A" & iRow + 1).Offset(0, 20).value = TextBox21.value
    Range("A" & iRow + 1).Offset(0, 21).value = TextBox22.value
    Range("A" & iRow + 1).Offset(0, 22).value = TextBox23.value
    Range("A" & iRow + 1).Offset(0, 23).value = TextBox24.value
    Range("A" & iRow + 1).Offset(0, 24).value = TextBox25.value
    Range("A" & iRow + 1).Offset(0, 25).value = ComboBox10.value
    Range("A" & iRow + 1).Offset(0, 26).value = TextBox27.value
    Range("A" & iRow + 1).Offset(0, 27).value = ComboBox11.value
    Range("A" & iRow + 1).Offset(0, 28).value = TextBox29.value
    Range("A" & iRow + 1).Offset(0, 29).value = TextBox30.value
    Range("A" & iRow + 1).Offset(0, 30).value = TextBox31.value
    Range("A" & iRow + 1).Offset(0, 31).value = TextBox32.value
    Range("A" & iRow + 1).Offset(0, 32).value = TextBox33.value
    Range("A" & iRow + 1).Offset(0, 33).value = TextBox34.value
    Range("A" & iRow + 1).Offset(0, 34).value = TextBox35.value
    Range("A" & iRow + 1).Offset(0, 35).value = TextBox36.value
    Range("A" & iRow + 1).Offset(0, 36).value = TextBox37.value
    Range("A" & iRow + 1).Offset(0, 37).value = TextBox38.value
    Range("A" & iRow + 1).Offset(0, 38).value = TextBox39.value
    Range("A" & iRow + 1).Offset(0, 39).value = TextBox40.value
    Range("A" & iRow + 1).Offset(0, 40).value = TextBox41.value
    Range("A" & iRow + 1).Offset(0, 41).value = TextBox42.value
    Range("A" & iRow + 1).Offset(0, 42).value = TextBox43.value
    Range("A" & iRow + 1).Offset(0, 43).value = TextBox44.value
    Range("A" & iRow + 1).Offset(0, 44).value = TextBox45.value
    Range("A" & iRow + 1).Offset(0, 45).value = TextBox46.value
    Range("A" & iRow + 1).Offset(0, 46).value = TextBox47.value
    Range("A" & iRow + 1).Offset(0, 47).value = TextBox48.value
    Range("A" & iRow + 1).Offset(0, 48).value = TextBox49.value
    Range("A" & iRow + 1).Offset(0, 49).value = TextBox50.value
    Range("A" & iRow + 1).Offset(0, 50).value = ComboBox9.value
     
   Range("A" & iRow + 1).Offset(0, 51).value = TextBox52.value
   Range("A" & iRow + 1).Offset(0, 52).value = TextBox53.value
   Range("A" & iRow + 1).Offset(0, 53).value = TextBox54.value
   Range("A" & iRow + 1).Offset(0, 54).value = TextBox55.value
   Range("A" & iRow + 1).Offset(0, 55).value = Format(TextBox56.Text, "yyyy/mm/dd")
   Range("A" & iRow + 1).Offset(0, 56).value = TextBox57.value
   Range("A" & iRow + 1).Offset(0, 57).value = TextBox58.value
   Range("A" & iRow + 1).Offset(0, 58).value = TextBox59.value
   Range("A" & iRow + 1).Offset(0, 59).value = TextBox60.value
   Range("A" & iRow + 1).Offset(0, 60).value = TextBox61.value
   Range("A" & iRow + 1).Offset(0, 61).value = TextBox62.value
   Range("A" & iRow + 1).Offset(0, 62).value = TextBox63.value
   Range("A" & iRow + 1).Offset(0, 74).value = TextBox99.value

      
        Sheets(3).Activate
        
'============================================================================

'On Error Resume Next
Sheets(1).Activate
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
For X = 9 To lastR
If Sheets(1).Cells(X, 2) = TextBox2.Text And Sheets(1).Cells(X, 5) = TextBox5.Text Then
 Sheets(1).Cells(X, 1).Select


End If
Next

ActiveCell.EntireRow.Delete
End If
End Sub





Private Sub CommandButton41_Click()
Call SORT_A_TO_Z1
End Sub

Private Sub CommandButton42_Click()
Sheets(1).Activate
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
For X = 9 To lastR
If Sheets(1).Cells(X, 2) = TextBox2.Text And Sheets(1).Cells(X, 5) = TextBox5.Text Then
 Sheets(1).Cells(X, 1).Select


End If
Next

ActiveCell.EntireRow.Delete
 MsgBox " „ Õ–› «·ﬁÌœ »‰Ã«Õ", vbOKOnly, "—”«·…  ‰»ÌÂ"

End Sub



Private Sub CommandButton40_Click()
Unload Me
UserForm34.Show
End Sub



Private Sub CommandButton44_Click()
Unload Me
UserForm51.Show
End Sub

Private Sub CommandButton45_Click()
Dim wsSrc As Worksheet, wsTgt As Worksheet
    Dim lastRowSrc As Long, LastCol As Long
    Dim lastRowTgt As Long, insertRow As Long
    Dim dataArr As Variant, blockArr As Variant
    Dim monthName As String, targetSheet As String
    Dim i As Long

    Set wsSrc = Sheets("Sheet1")

    ' ¬Œ— ⁄„Êœ Õ”» —ƒÊ” «·’› 8
    LastCol = wsSrc.Cells(8, wsSrc.Columns.count).End(xlToLeft).Column

    ' ‰”Œ —ƒÊ” «·√⁄„œ… „‰ «·’› 8 ≈·Ï ‘Ì  «·Âœ›
    monthName = Trim(wsSrc.Range("CB9").value)
    If monthName = "" Then
        MsgBox "⁄„Êœ «·‘Â— CB ›«—€", vbExclamation
        Exit Sub
    End If

    ' —»ÿ «”„ «·‘Â— „⁄ «”„ «·‘Ì 
    Select Case LCase(monthName)
        Case "Ì‰«Ì—", "ﬂ«‰Ê‰ «·À«‰Ì": targetSheet = "jan"
        Case "›»—«Ì—", "‘»«ÿ": targetSheet = "feb"
        Case "„«—”", "¬–«—": targetSheet = "mar"
        Case "«»—Ì·", "‰Ì”«‰": targetSheet = "apr"
        Case "„«ÌÊ", "«Ì«—": targetSheet = "may"
        Case "ÌÊ‰ÌÊ", "Õ“Ì—«‰": targetSheet = "jun"
        Case "ÌÊ·ÌÊ", " „Ê“": targetSheet = "jul"
        Case "«€”ÿ”", "¬»": targetSheet = "aug"
        Case "”» „»—", "«Ì·Ê·": targetSheet = "sep"
        Case "«ﬂ Ê»—", " ‘—Ì‰ «·«Ê·": targetSheet = "oct"
        Case "‰Ê›„»—", " ‘—Ì‰ «·À«‰Ì": targetSheet = "nov"
        Case "œÌ”„»—", "ﬂ«‰Ê‰ «·«Ê·": targetSheet = "dec"
        Case Else
            MsgBox "«”„ «·‘Â— €Ì— „⁄—Ê›: " & monthName, vbCritical
            Exit Sub
    End Select

    Set wsTgt = Sheets(targetSheet)

    ' ‰”Œ —ƒÊ” «·√⁄„œ… „‰ «·’› 8
    wsTgt.Rows(8).value = wsSrc.Rows(8).value

    '  √ﬂÌœ ﬁ»· «· —ÕÌ·
    If MsgBox("Â· √‰  „ √ﬂœ „‰  —ÕÌ· »Ì«‰«  ‘Â— " & monthName & " ø" & vbCrLf & _
              " √ﬂœ √‰ Ã„Ì⁄ «·„ €Ì—«  „ﬂ „·….", vbOKCancel + vbQuestion) = vbCancel Then
        MsgBox "Õ”‰«° ⁄‰œ ÿ·» «· —ÕÌ· ﬂ‰ „ √ﬂœ«.", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ¬Œ— ’› »Ì«‰«  ›Ì Sheet1
    lastRowSrc = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).row
    If lastRowSrc < 9 Then
        MsgBox "·«  ÊÃœ »Ì«‰«  ·· —ÕÌ·", vbExclamation
        Exit Sub
    End If

    '  Õ„Ì· »Ì«‰«  Sheet1 ≈·Ï „’›Ê›…
    dataArr = wsSrc.Range(wsSrc.Cells(9, 1), wsSrc.Cells(lastRowSrc, LastCol)).value

    ' Õ–› »Ì«‰«  Â–« «·‘Â— ›ﬁÿ „‰ ‘Ì  «·Âœ›
    lastRowTgt = wsTgt.Cells(wsTgt.Rows.count, 1).End(xlUp).row
    If lastRowTgt < 9 Then lastRowTgt = 8

    For i = lastRowTgt To 9 Step -1
        If wsTgt.Cells(i, 80).value = monthName Then  ' 80 = CB
            wsTgt.Rows(i).Delete
        End If
    Next i

    ' ≈⁄«œ… Õ”«» ¬Œ— ’› »⁄œ «·Õ–›
    lastRowTgt = wsTgt.Cells(wsTgt.Rows.count, 1).End(xlUp).row
    If lastRowTgt < 9 Then lastRowTgt = 8
    insertRow = lastRowTgt + 1

    ' ≈œ—«Ã »Ì«‰«  «·‘Â— ﬂ„’›Ê›…
    wsTgt.Cells(insertRow, 1).Resize(UBound(dataArr, 1), UBound(dataArr, 2)).value = dataArr

    '  —ÕÌ· «·ﬂ ·… CC1:CQ5 ﬂ„’›Ê›…
    blockArr = wsSrc.Range("CC1:CQ5").value
    wsTgt.Range("CC1:CQ5").value = blockArr

    '  √ﬂœ „‰ „”Õ √Ì Ê÷⁄ ‰”Œ ”«»ﬁ
    Application.CutCopyMode = False

    Application.ScreenUpdating = True

    MsgBox " „  —ÕÌ· »Ì«‰«  ‘Â— " & monthName & " ≈·Ï «·‘Ì  " & targetSheet & " »‰Ã«Õ.", vbInformation

End Sub

Private Sub CommandButton46_Click()
 Dim selectedSheet As String
    selectedSheet = Me.ComboBox16.value
    
    If selectedSheet = "" Then
        MsgBox "—Ã«¡ «Œ «— «·‘Â— „‰ «·ﬁ«∆„….", vbExclamation
        Exit Sub
    End If
    
    ' «” œ⁄«¡ «·≈Ã—«¡ «·„ÊÃÊœ œ«Œ· ‰›” «·‹ UserForm
    Call GetMonthDataToSheet1(selectedSheet)
    
    MsgBox " „ Ã·» »Ì«‰«  «·‘Â— " & selectedSheet & " ≈·Ï Sheet1 »‰Ã«Õ.", vbInformation
End Sub

Private Sub CommandButton47_Click()
Unload Me
UserForm59.Show
End Sub

Private Sub CommandButton48_Click()
Unload Me
UserForm61.Show
End Sub

Private Sub CommandButton49_Click()
  Dim ws As Worksheet, wsMap As Worksheet
    Dim lastRow As Long, i As Long, r As Long, c As Long
    Dim jobTitle As String, jobDesc As String
    Dim cellValue As String
    Dim found As Boolean

    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set wsMap = ThisWorkbook.Sheets("Sheet10")

    lastRow = ws.Cells(ws.Rows.count, "F").End(xlUp).row

    For i = 9 To lastRow
        jobTitle = Trim(ws.Cells(i, "F").value)
        jobDesc = ""
        found = False

        If jobTitle <> "" Then
            ' Õ«·… Œ«’…: "„œÌ—" ·ÊÕœÂ
            If LCase(jobTitle) = "„œÌ—" Then
                jobDesc = "«œ«—Ì"
                found = True
            Else
                jobTitle = NormalizeText(jobTitle)
                ' «·»ÕÀ «·œﬁÌﬁ √Ê·«
                For r = 2 To 61
                    For c = 2 To 12
                        cellValue = NormalizeText(wsMap.Cells(r, c).value)
                        If cellValue <> "" Then
                            If jobTitle = cellValue Then
                                jobDesc = wsMap.Cells(r, "A").value
                                found = True
                                Exit For
                            End If
                        End If
                    Next c
                    If found Then Exit For
                Next r

                ' «·»ÕÀ «· ﬁ—Ì»Ì »‰«¡ ⁄·Ï «·ﬁÊ«⁄œ «·À«» …
                If Not found Then
                    jobDesc = GetFixedRule(jobTitle)
                    If jobDesc <> "" Then found = True
                End If

                ' ≈–« ·„ ‰Ãœ ‘Ì¡ ? „Õ«Ê·… „ÿ«»ﬁ…  ﬁ—Ì»Ì… ⁄«„…
                If Not found Then
                    For r = 2 To 61
                        For c = 2 To 12
                            cellValue = NormalizeText(wsMap.Cells(r, c).value)
                            If cellValue <> "" Then
                                If IsApproxMatch(jobTitle, cellValue) Then
                                    jobDesc = wsMap.Cells(r, "A").value
                                    found = True
                                    Exit For
                                End If
                            End If
                        Next c
                        If found Then Exit For
                    Next r
                End If
            End If
        End If

        ws.Cells(i, "CF").value = jobDesc
    Next i
    MsgBox "·ﬁœ  „ «⁄«œ… »‰«¡ «·Ê’› «·ÊŸÌ›Ì »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
End Sub
Private Function NormalizeText(txt As String) As String
    txt = Trim(txt)
    txt = Replace(txt, "√", "«")
    txt = Replace(txt, "≈", "«")
    txt = Replace(txt, "¬", "«")
    txt = Replace(txt, "Ï", "Ì")
    txt = Replace(txt, "∆", "Ì")
    txt = Replace(txt, "¡", "")
    txt = Replace(txt, "„‰»Ì", "„‰»∆")
    NormalizeText = txt
End Function
Private Function IsApproxMatch(text1 As String, text2 As String) As Boolean
    Dim t1 As String, t2 As String
    t1 = Trim(text1)
    t2 = Trim(text2)

    ' ≈“«·… “Ì«œ«  Õ—› «·Ì«¡ √Ê "Ì‰"
    Do While Right(t1, 1) = "Ì" Or Right(t1, 2) = "Ì‰"
        If Right(t1, 2) = "Ì‰" Then t1 = Left(t1, Len(t1) - 2) Else t1 = Left(t1, Len(t1) - 1)
    Loop
    Do While Right(t2, 1) = "Ì" Or Right(t2, 2) = "Ì‰"
        If Right(t2, 2) = "Ì‰" Then t2 = Left(t2, Len(t2) - 2) Else t2 = Left(t2, Len(t2) - 1)
    Loop

    If InStr(t1, t2) > 0 Or InStr(t2, t1) > 0 Then
        IsApproxMatch = True
    Else
        IsApproxMatch = False
    End If
End Function
Private Function GetFixedRule(jobTitle As String) As String
    Dim title As String
    title = LCase(jobTitle)

    ' ﬁ«⁄œ… ›‰Ì: √Ì ⁄‰Ê«‰ ÌÕ ÊÌ "›‰Ì" √Ê "›‰ÌÌ‰" √Ê "›‰Ì‰"
    If InStr(title, "›‰Ì") > 0 Or InStr(title, "›‰ÌÌ‰") > 0 Or InStr(title, "›‰Ì‰") > 0 Then
        GetFixedRule = "›‰Ì"
        Exit Function
    End If

    ' «·ﬁÊ«⁄œ «·À«» … «·√Œ—Ï
    Select Case title
        Case "—∆Ì” „·«ÕŸÌ‰", "—∆Ì” „·«ÕŸÌ‰ ›‰ÌÌ‰", "„·«ÕŸ"
            GetFixedRule = "›‰Ì"
        Case "„œÌ—  ‰›Ì–Ì"
            GetFixedRule = "„»—„Ã"
        Case "„⁄«Ê‰ „œÌ— ›‰Ì"
            GetFixedRule = "›‰Ì"
        Case Else
            GetFixedRule = "" ' ·« ÌÊÃœ ﬁ«⁄œ… À«» …
    End Select
End Function


Private Sub CommandButton50_Click()
UserForm63.Show
End Sub

Private Sub CommandButton51_Click()
UserForm70.Show
End Sub

Private Sub CommandButton6_Click()

'On Error GoTo emad
 
'========================================================

If MsgBox("   ***Â·  —€» »«œŒ«· «·»Ì«‰«  ⁄»— „‰’… Ê«·Ê«ÃÂ… «·Õ«·ÌÂ  øø    **** " & vbCrLf & " Ê«–« ﬂ«‰ ·œÌﬂ ”Ã·«  „ Ê›—Â ›Ì „’‰› Œ«—ÃÌ „‰ «·„„ﬂ‰ «” Ì—«œÂ« ⁄»— »Ê«»… «·« ’«· «·ﬂ —Ê‰Ì «÷€ÿ ⁄·Ï “— NO    ", vbYesNo, "—”«·…  ÊÃÌÂ") = vbYes Then
     
     Sheets(1).Activate
     Dim iRow As Long, i As Long

    iRow = Cells(Rows.count, "b").End(xlUp).row
    TextBox1.value = iRow - 7
     TextBox2.value = ""
     TextBox3.value = ""
     
     TextBox5.value = ""
     TextBox7.value = ""
  
  TextBox7.value = ""
   TextBox8.value = 0
  TextBox9.value = 0
  TextBox10.value = 0
   TextBox11.value = 0
    TextBox12.value = 0
      
     TextBox14.value = ""
    
   TextBox16.value = ""
   
   TextBox18.value = ""
    TextBox19.value = 0
   TextBox20.value = ""
    TextBox21.value = 0
  TextBox22.value = 0
   TextBox23.value = ""
TextBox24.value = ""
   TextBox25.value = 0
  
   TextBox27.value = ""
   
  TextBox29.value = ""
   TextBox30.value = ""
 TextBox31.value = ""
  TextBox32.value = ""
   'TextBox33.Text = ""
   TextBox34.value = 0
   TextBox35.value = 0
   TextBox36.value = 0
  TextBox37.value = ""
 TextBox38.value = 0
     TextBox39.value = 0
    TextBox40.value = 0
     TextBox41.value = ""
   TextBox42.value = 0
    TextBox43.value = ""
 TextBox44.value = 0
   TextBox45.value = ""
    TextBox46.value = 0
    TextBox47.value = ""
   TextBox48.value = ""
   TextBox49.value = ""
      TextBox50.value = ""
   TextBox52.value = ""
   TextBox56.value = ""
   
   TextBox54.value = ""
    TextBox55.value = ""
   TextBox57.value = ""
   
'============================================================
   
 TextBox14.Enabled = False
    
  TextBox16.Enabled = False
  TextBox18.Enabled = False
  TextBox20.Enabled = False
  TextBox24.Enabled = False
  TextBox27.Enabled = False
  TextBox29.Enabled = False
  TextBox30.Enabled = False
  TextBox31.Enabled = False
  TextBox32.Enabled = False
  TextBox33.Enabled = False
  TextBox48.Enabled = False
  TextBox49.Enabled = False
   
    
  
     
   
    
   
   Else
   UserForm19.Show
   End If
'emad:
End Sub



Private Sub CommandButton8_Click()
Sheets(1).Activate
For X = 9 To 1500
If Cells(X, 5) = TextBox5.Text Then
Cells(X, 5).Select
Exit For
End If
Next X
'==================================


ActiveCell.EntireRow.Delete
End Sub

Private Sub CommandButton7_Click()

End Sub

Private Sub CommandButton9_Click()
On Error GoTo emad
Dim simadd1 As Double
Dim simsubtract1 As Double
Dim item1 As ListItem
Dim last1, frw1 As Integer
Dim rngData As Range
Dim wksSource As Worksheet
 Dim fString As Variant
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row



'=============================================

   Sheets(1).Activate
    
    
    For frw1 = 9 To last1
                
 If Sheets(1).Cells(frw1, "h") = TextBox67.Text And Sheets(1).Cells(frw1, "f") = ComboBox2.Text And ComboBox3.value = Sheets(1).Cells(frw1, "bf").value Then
     
'==================================================

Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ad9:ad8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ae9:ae8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("av9:av8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aw9:aw8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("x9:x8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text, Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*")






UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")




End If
Next frw1
End With

'============================================= «·Õ”«»«  «·ﬂ·ÌÂ

Dim simadd As Double
Dim simsubtract As Double
Dim item2 As ListItem
Dim last2, frw2 As Integer
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2



If ComboBox2.value = "«·ﬂ·" And Sheets(1).Cells(frw2, "bf") = ComboBox3.value Then

Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))




item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(2) = Sheets(1).Cells(frw2, "c")
item2.SubItems(3) = Sheets(1).Cells(frw2, "d")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(7) = Sheets(1).Cells(frw2, "h")
item2.SubItems(8) = Sheets(1).Cells(frw2, "i")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(10) = Sheets(1).Cells(frw2, "k")
item2.SubItems(11) = Sheets(1).Cells(frw2, "l")
item2.SubItems(12) = Sheets(1).Cells(frw2, "m")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(14) = Sheets(1).Cells(frw2, "o")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(16) = Sheets(1).Cells(frw2, "q")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(21) = Sheets(1).Cells(frw2, "v")
item2.SubItems(22) = Sheets(1).Cells(frw2, "w")
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(25) = Sheets(1).Cells(frw2, "z")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(1).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(1).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(38) = Sheets(1).Cells(frw2, "am")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(40) = Sheets(1).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(1).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(1).Cells(frw2, "as")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(46) = Sheets(1).Cells(frw2, "au")
item2.SubItems(47) = Sheets(1).Cells(frw2, "av")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(1).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(1).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(1).Cells(frw2, "az")
item2.SubItems(52) = Sheets(1).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(1).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(1).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(1).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(1).Cells(frw2, "be")
item2.SubItems(57) = Sheets(1).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(1).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(1).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(1).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(1).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(1).Cells(frw2, "bk")
'item2.SubItems(63) = Sheets(1).Cells(frw2, "bl")
simadd = Application.WorksheetFunction.Sum(Sheets(1).Range("bh9:bh" & frw2))
simsubtract = Application.WorksheetFunction.Sum(Sheets(1).Range("bj9:bj" & last2))
TextBox73.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("j9:j" & last2))
TextBox74.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("ad9:ad" & last2))
TextBox79.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("ae9:ae" & last2))
TextBox80.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("av9:av" & last2))
TextBox66.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("aw9:aw" & last2))
TextBox76.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("X9:X" & last2))
TextBox67.Text = ""

TextBox66.Text = (TextBox66.Text + simadd) - simsubtract
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")




ElseIf Sheets(1).Cells(frw1, "h") Like "*" & TextBox67.Text & "*" And Sheets(1).Cells(frw1, "f") Like "*" & ComboBox2.Text & "*" And ComboBox3.value = Sheets(1).Cells(frw1, "bf").value Then
  item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(2) = Sheets(1).Cells(frw2, "c")
item2.SubItems(3) = Sheets(1).Cells(frw2, "d")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(7) = Sheets(1).Cells(frw2, "h")
item2.SubItems(8) = Sheets(1).Cells(frw2, "i")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(10) = Sheets(1).Cells(frw2, "k")
item2.SubItems(11) = Sheets(1).Cells(frw2, "l")
item2.SubItems(12) = Sheets(1).Cells(frw2, "m")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(14) = Sheets(1).Cells(frw2, "o")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(16) = Sheets(1).Cells(frw2, "q")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(21) = Sheets(1).Cells(frw2, "v")
item2.SubItems(22) = Sheets(1).Cells(frw2, "w")
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(25) = Sheets(1).Cells(frw2, "z")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(1).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(1).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(38) = Sheets(1).Cells(frw2, "am")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(40) = Sheets(1).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(1).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(1).Cells(frw2, "as")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(46) = Sheets(1).Cells(frw2, "au")
item2.SubItems(47) = Sheets(1).Cells(frw2, "av")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(1).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(1).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(1).Cells(frw2, "az")
item2.SubItems(52) = Sheets(1).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(1).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(1).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(1).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(1).Cells(frw2, "be")
item2.SubItems(57) = Sheets(1).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(1).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(1).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(1).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(1).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(1).Cells(frw2, "bk")
item2.SubItems(63) = Sheets(1).Cells(frw2, "bl")





'==================================================
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ad9:ad800"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ae9:ae8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("av9:av8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aw9:aw8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("x9:x8000"), Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("h9:h8000"), "=" & Me.TextBox67.Text, Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*")






UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")
End If
Next






'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(1).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("BF9:BF571"), "=" & Me.ComboBox3.Value)


'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("BF9:BF" & frw2), "=" & Me.ComboBox3.Text)
TextBox75.Text = ListView1.ListItems.count




emad:



End Sub



Private Sub Frame_Click()

End Sub

Private Sub Frame6_Click()

End Sub

Private Sub Label108_Click()


End Sub



Private Sub Label125_Click()

End Sub

Private Sub Label124_Click()
UserForm41.Show
End Sub

Private Sub Label127_Click()
 Label127.BackColor = RGB(0, 255, 255)
    
    '  ‘€Ì· Õ—ﬂ… «· ﬁ·’ («·ÿÊ· 0 ·ÌŒ ›Ì ÊÌŸÂ— „«  Õ Â)
    AnimateLabel 78, 0
    Label127.Visible = False
End Sub

Private Sub ListBox1_AfterUpdate()
On Error Resume Next
Call arb

End Sub

Private Sub ListBox1_Click()
On Error Resume Next
TextBox92.Text = Sheets(1).Cells(1, "cb")
Call arb
'=================================

Call UpdateSummaryFormula









'=================================


Sheets(1).Activate
If TextBox2.Text = "" And TextBox5.Text = "" Then
Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
End If
For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) = True Then
        For j = 1 To 63
        Controls("TextBox" & j).Text = Cells(ListBox1.List(i, 1), j)
       TextBox10.Visible = False
        Next j
    End If
Next i
ComboBox1.Visible = False
'===================================
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(1).Cells(frw1, "b") = TextBox2.Text And Sheets(1).Cells(frw1, "e") = TextBox5.Text Then
Image1.Picture = LoadPicture(Sheets(1).Cells(frw1, 64))
Image2.Picture = LoadPicture(Sheets(1).Cells(frw1, 65))
ComboBox4.Text = Sheets(1).Cells(frw1, 4)
ComboBox5.Text = Sheets(1).Cells(frw1, 6)
ComboBox6.Text = Sheets(1).Cells(frw1, "m")
ComboBox7.Text = Sheets(1).Cells(frw1, "o")
ComboBox8.Text = Sheets(1).Cells(frw1, "q")
ComboBox9.Text = Sheets(1).Cells(frw1, "ay")
ComboBox10.Text = Sheets(1).Cells(frw1, "z")
ComboBox11.Text = Sheets(1).Cells(frw1, "ab")
ComboBox3.Text = Sheets(1).Cells(frw1, "bf")
IsLoading = True
CheckBox5.value = Sheets(1).Cells(frw1, "cg")
IsLoading = False
TextBox82.Text = Format(Sheets(1).Cells(frw1, "bn"), "yyyy/mm/dd")

TextBox83.Text = Sheets(1).Cells(frw1, "bo")
TextBox84.Text = Sheets(1).Cells(frw1, "bp")
TextBox85.Text = Sheets(1).Cells(frw1, "bq")
TextBox86.Text = Sheets(1).Cells(frw1, "br")
TextBox87.Text = Sheets(1).Cells(frw1, "bs")
TextBox88.Text = Sheets(1).Cells(frw1, "bt")
TextBox91.Text = Sheets(1).Cells(frw1, "bv")
TextBox99.Text = Sheets(1).Cells(frw1, "bu")
TextBox100.Text = Sheets(1).Cells(frw1, "ca")
TextBox105.Text = Sheets(1).Cells(frw1, "ce")
TextBox106.Text = Sheets(1).Cells(frw1, "cf")
TextBox107.Text = Sheets(1).Cells(4, "ci")

CheckBox4.value = Sheets(1).Cells(frw1, "cd")


'=====================

'=========================


If TextBox58.Text = "⁄ﬁœ" Then
TextBox10.Enabled = False
Label123.Visible = True
TextBox10.Text = Sheets(1).Cells(frw1, "j")
TextBox10.Visible = False
TextBox101.Visible = True

TextBox101.Text = Sheets(1).Cells(frw1, "j")

Else
TextBox101.Visible = False
TextBox10.Visible = True
Label123.Visible = False
End If
End If
Next frw1
If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If


If TextBox58.Text = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Then

TextBox20.Text = 0

End If


TextBox89.value = Sheets(1).Cells(2, "bw").value
TextBox90.value = Sheets(1).Cells(2, "bx").value

If TextBox58.Text = "—« » Ã“∆Ì" Or TextBox58.Text = "«„Ê„Â À«‰Ì 6 «‘Â—" Then
CommandButton16.Enabled = False
Else
CommandButton16.Enabled = True
End If

'===========================================
Sheets(1).Cells(4, "cb").value = TextBox58.Text

TextBox93.value = Sheets(1).Cells(4, "cd")
TextBox94.value = Sheets(1).Cells(4, "ce")
TextBox97.value = Sheets(1).Cells(4, "cf")
TextBox98.value = Sheets(1).Cells(4, "cg")
TextBox95.value = Sheets(1).Cells(4, "ch")
TextBox96.value = Sheets(1).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

TextBox75.Text = ListView1.ListItems.count
UserForm1.TextBox93.value = Format(UserForm1.TextBox93.value, "#,## IQD")
UserForm1.TextBox94.value = Format(UserForm1.TextBox94.value, "#,## IQD")
UserForm1.TextBox97.value = Format(UserForm1.TextBox97.value, "#,## IQD")
UserForm1.TextBox98.value = Format(UserForm1.TextBox98.value, "#,## IQD")
UserForm1.TextBox95.value = Format(UserForm1.TextBox95.value, "#,## IQD")
UserForm1.TextBox96.value = Format(UserForm1.TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & TextBox58.Text

TextBox89.value = Format(val(TextBox85) / val(TextBox10) * 100, "00.0") & "%"
TextBox90.value = Format(val(TextBox86) / val(TextBox10) * 100, "00.0") & "%"

ComboBox12.value = ""
ComboBox13.value = ""
TextBox103.value = "„ﬁœ«— ‰”»… «·„‹‹‹Œ’’«  «·„‹‹∆ÊÌ…" & " = " & (val(TextBox14) + val(TextBox16) + val(TextBox18) + val(TextBox27) + val(TextBox29)) / val(TextBox10) * 100 & "%"
'==========================

If ComboBox4.value = "«‰ÀÏ" And TextBox23.Text = "«·„ÊŸ› «·„ “ÊÃ Ê “ÊÃ Â —»… »Ì " Then
MsgBox "«·‰Ÿ«„ «·Ÿ—Ì»Ì Ì—›÷ Â–Â «·„⁄«ÌÌ— ·ﬂÊ‰Â« «‰ÀÏ Ê·« Ì„ﬂ‰ «‰  ‰ÿ»ﬁ ⁄·Ï “ÊÃÂ« —»  »Ì øøø", vbCritical, "—”«·…  ‰»ÌÂ"



ElseIf TextBox23.Text = "«·«⁄“»" And TextBox22.Text > 0 Then
MsgBox "Â‰«ﬂ ⁄œ„  Ê«›ﬁ „‰ÿﬁÌ ··‰Ÿ«„ «·÷—Ì»Ì ﬂ√‰ ÌﬂÊ‰ «⁄“» Ê·œÌÂ «ÿ›«· øøø ", vbCritical, "—”«·…  ‰»ÌÂ"


End If
Call UpdateSummaryFormula
Call CaptureSnapshot(Me, DictOld)
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ComboBox1.Visible = False

If ComboBox6.value = "" Or ComboBox7.value = "" Or ComboBox8.value = "" Or ComboBox10.value = "" Or ComboBox11.value = "" Then
ComboBox6.value = 0
ComboBox7.value = 0
ComboBox8.value = 0
ComboBox10.value = 0
ComboBox11.value = 0
Else
Exit Sub
End If











End Sub



Private Sub ListBox1_Enter()

End Sub

Private Sub ListBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
Call arb

End Sub

Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
Call arb




End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub ListView1_Click()
On Error Resume Next
TextBox1.Text = ListView1.SelectedItem.ListSubItems(0)
TextBox2.Text = ListView1.SelectedItem.ListSubItems(1)
TextBox3.Text = ListView1.SelectedItem.ListSubItems(2)
ComboBox4.value = ListView1.SelectedItem.ListSubItems(3)
TextBox5.Text = ListView1.SelectedItem.ListSubItems(4)
ComboBox5.value = ListView1.SelectedItem.ListSubItems(5)
TextBox7.Text = ListView1.SelectedItem.ListSubItems(6)
TextBox8.Text = ListView1.SelectedItem.ListSubItems(7)
TextBox9.Text = ListView1.SelectedItem.ListSubItems(8)
TextBox10.Text = ListView1.SelectedItem.ListSubItems(9)
TextBox101.Text = ListView1.SelectedItem.ListSubItems(9)
TextBox11.Text = ListView1.SelectedItem.ListSubItems(10)
TextBox12.Text = ListView1.SelectedItem.ListSubItems(11)
ComboBox6.Text = ListView1.SelectedItem.ListSubItems(12)
TextBox14.Text = ListView1.SelectedItem.ListSubItems(13)
ComboBox7.Text = ListView1.SelectedItem.ListSubItems(14)
TextBox16.Text = ListView1.SelectedItem.ListSubItems(15)
ComboBox8.Text = ListView1.SelectedItem.ListSubItems(16)
TextBox18.Text = ListView1.SelectedItem.ListSubItems(17)
TextBox19.Text = ListView1.SelectedItem.ListSubItems(18)
TextBox20.Text = ListView1.SelectedItem.ListSubItems(19)
TextBox21.Text = ListView1.SelectedItem.ListSubItems(20)
TextBox22.Text = ListView1.SelectedItem.ListSubItems(21)

ComboBox1.Text = ListView1.SelectedItem.ListSubItems(22)
TextBox24.Text = ListView1.SelectedItem.ListSubItems(23)
TextBox25.Text = ListView1.SelectedItem.ListSubItems(24)
TextBox27.Text = ListView1.SelectedItem.ListSubItems(26)
ComboBox10.value = ListView1.SelectedItem.ListSubItems(25)
ComboBox11.value = ListView1.SelectedItem.ListSubItems(27)
TextBox29.value = ListView1.SelectedItem.ListSubItems(28)
TextBox30.Text = ListView1.SelectedItem.ListSubItems(29)
TextBox31.Text = ListView1.SelectedItem.ListSubItems(30)
TextBox32.Text = ListView1.SelectedItem.ListSubItems(31)
TextBox33.Text = ListView1.SelectedItem.ListSubItems(32)
TextBox34.Text = ListView1.SelectedItem.ListSubItems(33)
TextBox35.Text = ListView1.SelectedItem.ListSubItems(34)
TextBox36.Text = ListView1.SelectedItem.ListSubItems(35)
TextBox37.Text = ListView1.SelectedItem.ListSubItems(36)
TextBox38.Text = ListView1.SelectedItem.ListSubItems(37)
TextBox39.Text = ListView1.SelectedItem.ListSubItems(38)
TextBox40.Text = ListView1.SelectedItem.ListSubItems(39)
TextBox41.Text = ListView1.SelectedItem.ListSubItems(40)
TextBox42.Text = ListView1.SelectedItem.ListSubItems(41)
TextBox43.Text = ListView1.SelectedItem.ListSubItems(42)
TextBox46.Text = ListView1.SelectedItem.ListSubItems(45)
TextBox48.Text = ListView1.SelectedItem.ListSubItems(47)
TextBox49.Text = ListView1.SelectedItem.ListSubItems(48)
ComboBox9.value = ListView1.SelectedItem.ListSubItems(50)
TextBox52.Text = ListView1.SelectedItem.ListSubItems(51)
TextBox56.Text = ListView1.SelectedItem.ListSubItems(55)
TextBox57.Text = ListView1.SelectedItem.ListSubItems(56)
TextBox58.Text = ListView1.SelectedItem.ListSubItems(57)
ComboBox3.Text = ListView1.SelectedItem.ListSubItems(57)
TextBox54.Text = ListView1.SelectedItem.ListSubItems(53)
TextBox59.Text = ListView1.SelectedItem.ListSubItems(58)
TextBox55.Text = ListView1.SelectedItem.ListSubItems(54)
TextBox60.Text = ListView1.SelectedItem.ListSubItems(59)
TextBox62.Text = ListView1.SelectedItem.ListSubItems(61)
TextBox53.Text = ListView1.SelectedItem.ListSubItems(52)
TextBox61.Text = ListView1.SelectedItem.ListSubItems(60)
TextBox82.Text = ListView1.SelectedItem.ListSubItems(65)
TextBox82.Text = ListView1.SelectedItem.ListSubItems(65)
TextBox83.Text = ListView1.SelectedItem.ListSubItems(66)
TextBox84.Text = ListView1.SelectedItem.ListSubItems(67)
TextBox85.Text = ListView1.SelectedItem.ListSubItems(68)
TextBox86.Text = ListView1.SelectedItem.ListSubItems(69)
TextBox50.Text = ListView1.SelectedItem.ListSubItems(49)
TextBox63.Text = ListView1.SelectedItem.ListSubItems(62)











'========================================




'TextBox58.Text = ComboBox3.Value
'TextBox100.Text = ComboBox3.Column(1)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
TextBox1.Text = ListView1.SelectedItem + 1 - 1
TextBox2.Text = ListView1.SelectedItem.ListSubItems(1)
TextBox3.Text = ListView1.SelectedItem.ListSubItems(2)
ComboBox4.value = ListView1.SelectedItem.ListSubItems(3)
TextBox5.Text = ListView1.SelectedItem.ListSubItems(4)
ComboBox5.value = ListView1.SelectedItem.ListSubItems(5)
TextBox7.Text = ListView1.SelectedItem.ListSubItems(6)
TextBox8.Text = ListView1.SelectedItem.ListSubItems(7)
TextBox9.Text = ListView1.SelectedItem.ListSubItems(8)
TextBox10.Text = ListView1.SelectedItem.ListSubItems(9)
TextBox101.Text = ListView1.SelectedItem.ListSubItems(9)
TextBox11.Text = ListView1.SelectedItem.ListSubItems(10)
TextBox12.Text = ListView1.SelectedItem.ListSubItems(11)
ComboBox6.Text = ListView1.SelectedItem.ListSubItems(12)
TextBox14.Text = ListView1.SelectedItem.ListSubItems(13)
ComboBox7.Text = ListView1.SelectedItem.ListSubItems(14)
TextBox16.Text = ListView1.SelectedItem.ListSubItems(15)
ComboBox8.Text = ListView1.SelectedItem.ListSubItems(16)
TextBox18.Text = ListView1.SelectedItem.ListSubItems(17)
TextBox19.Text = ListView1.SelectedItem.ListSubItems(18)
TextBox20.Text = ListView1.SelectedItem.ListSubItems(19)
TextBox21.Text = ListView1.SelectedItem.ListSubItems(20)
TextBox22.Text = ListView1.SelectedItem.ListSubItems(21)

ComboBox1.Text = ListView1.SelectedItem.ListSubItems(22)
TextBox24.Text = ListView1.SelectedItem.ListSubItems(23)
TextBox25.Text = ListView1.SelectedItem.ListSubItems(24)
TextBox27.Text = ListView1.SelectedItem.ListSubItems(26)
ComboBox10.value = ListView1.SelectedItem.ListSubItems(25)
ComboBox11.value = ListView1.SelectedItem.ListSubItems(27)
TextBox29.value = ListView1.SelectedItem.ListSubItems(28)
TextBox30.Text = ListView1.SelectedItem.ListSubItems(29)
TextBox31.Text = ListView1.SelectedItem.ListSubItems(30)
TextBox32.Text = ListView1.SelectedItem.ListSubItems(31)
TextBox33.Text = ListView1.SelectedItem.ListSubItems(32)
TextBox34.Text = ListView1.SelectedItem.ListSubItems(33)
TextBox35.Text = ListView1.SelectedItem.ListSubItems(34)
TextBox36.Text = ListView1.SelectedItem.ListSubItems(35)
TextBox37.Text = ListView1.SelectedItem.ListSubItems(36)
TextBox38.Text = ListView1.SelectedItem.ListSubItems(37)
TextBox39.Text = ListView1.SelectedItem.ListSubItems(38)
TextBox40.Text = ListView1.SelectedItem.ListSubItems(39)
TextBox41.Text = ListView1.SelectedItem.ListSubItems(40)
TextBox42.Text = ListView1.SelectedItem.ListSubItems(41)
TextBox43.Text = ListView1.SelectedItem.ListSubItems(42)
TextBox46.Text = ListView1.SelectedItem.ListSubItems(45)
TextBox48.Text = ListView1.SelectedItem.ListSubItems(47)
TextBox49.Text = ListView1.SelectedItem.ListSubItems(48)
ComboBox9.value = ListView1.SelectedItem.ListSubItems(50)
TextBox52.Text = ListView1.SelectedItem.ListSubItems(51)
TextBox56.Text = ListView1.SelectedItem.ListSubItems(55)
TextBox57.Text = ListView1.SelectedItem.ListSubItems(56)
TextBox54.Text = ListView1.SelectedItem.ListSubItems(53)
TextBox59.Text = ListView1.SelectedItem.ListSubItems(58)
TextBox55.Text = ListView1.SelectedItem.ListSubItems(54)
TextBox60.Text = ListView1.SelectedItem.ListSubItems(59)
TextBox62.Text = ListView1.SelectedItem.ListSubItems(61)
TextBox53.Text = ListView1.SelectedItem.ListSubItems(52)
TextBox61.Text = ListView1.SelectedItem.ListSubItems(60)
TextBox82.Text = ListView1.SelectedItem.ListSubItems(65)
TextBox82.Text = ListView1.SelectedItem.ListSubItems(65)
TextBox83.Text = ListView1.SelectedItem.ListSubItems(66)
TextBox84.Text = ListView1.SelectedItem.ListSubItems(67)
TextBox85.Text = ListView1.SelectedItem.ListSubItems(68)
TextBox86.Text = ListView1.SelectedItem.ListSubItems(69)
TextBox50.Text = ListView1.SelectedItem.ListSubItems(49)
TextBox63.Text = ListView1.SelectedItem.ListSubItems(62)




End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)

On Error Resume Next
TextBox1.Text = ListView1.SelectedItem.ListSubItems(0)
TextBox2.Text = ListView1.SelectedItem.ListSubItems(1)
TextBox3.Text = ListView1.SelectedItem.ListSubItems(2)
ComboBox4.value = ListView1.SelectedItem.ListSubItems(3)
TextBox5.Text = ListView1.SelectedItem.ListSubItems(4)
ComboBox5.value = ListView1.SelectedItem.ListSubItems(5)
TextBox7.Text = ListView1.SelectedItem.ListSubItems(6)
TextBox8.Text = ListView1.SelectedItem.ListSubItems(7)
TextBox9.Text = ListView1.SelectedItem.ListSubItems(8)
TextBox10.Text = ListView1.SelectedItem.ListSubItems(9)
TextBox101.Text = ListView1.SelectedItem.ListSubItems(9)
TextBox11.Text = ListView1.SelectedItem.ListSubItems(10)
TextBox12.Text = ListView1.SelectedItem.ListSubItems(11)
ComboBox6.Text = ListView1.SelectedItem.ListSubItems(12)
TextBox14.Text = ListView1.SelectedItem.ListSubItems(13)
ComboBox7.Text = ListView1.SelectedItem.ListSubItems(14)
TextBox16.Text = ListView1.SelectedItem.ListSubItems(15)
ComboBox8.Text = ListView1.SelectedItem.ListSubItems(16)
TextBox18.Text = ListView1.SelectedItem.ListSubItems(17)
TextBox19.Text = ListView1.SelectedItem.ListSubItems(18)
TextBox20.Text = ListView1.SelectedItem.ListSubItems(19)
TextBox21.Text = ListView1.SelectedItem.ListSubItems(20)
TextBox22.Text = ListView1.SelectedItem.ListSubItems(21)

ComboBox1.Text = ListView1.SelectedItem.ListSubItems(22)
TextBox24.Text = ListView1.SelectedItem.ListSubItems(23)
TextBox25.Text = ListView1.SelectedItem.ListSubItems(24)
TextBox27.Text = ListView1.SelectedItem.ListSubItems(26)
ComboBox10.value = ListView1.SelectedItem.ListSubItems(25)
ComboBox11.value = ListView1.SelectedItem.ListSubItems(27)
TextBox29.value = ListView1.SelectedItem.ListSubItems(28)
TextBox30.Text = ListView1.SelectedItem.ListSubItems(29)
TextBox31.Text = ListView1.SelectedItem.ListSubItems(30)
TextBox32.Text = ListView1.SelectedItem.ListSubItems(31)
TextBox33.Text = ListView1.SelectedItem.ListSubItems(32)
TextBox34.Text = ListView1.SelectedItem.ListSubItems(33)
TextBox35.Text = ListView1.SelectedItem.ListSubItems(34)
TextBox36.Text = ListView1.SelectedItem.ListSubItems(35)
TextBox37.Text = ListView1.SelectedItem.ListSubItems(36)
TextBox38.Text = ListView1.SelectedItem.ListSubItems(37)
TextBox39.Text = ListView1.SelectedItem.ListSubItems(38)
TextBox40.Text = ListView1.SelectedItem.ListSubItems(39)
TextBox41.Text = ListView1.SelectedItem.ListSubItems(40)
TextBox42.Text = ListView1.SelectedItem.ListSubItems(41)
TextBox43.Text = ListView1.SelectedItem.ListSubItems(42)
TextBox46.Text = ListView1.SelectedItem.ListSubItems(45)
TextBox48.Text = ListView1.SelectedItem.ListSubItems(47)
TextBox49.Text = ListView1.SelectedItem.ListSubItems(48)
ComboBox9.value = ListView1.SelectedItem.ListSubItems(50)
TextBox52.Text = ListView1.SelectedItem.ListSubItems(51)
TextBox56.Text = ListView1.SelectedItem.ListSubItems(55)
TextBox57.Text = ListView1.SelectedItem.ListSubItems(56)
TextBox54.Text = ListView1.SelectedItem.ListSubItems(53)
TextBox59.Text = ListView1.SelectedItem.ListSubItems(58)
TextBox55.Text = ListView1.SelectedItem.ListSubItems(54)
TextBox60.Text = ListView1.SelectedItem.ListSubItems(59)
TextBox62.Text = ListView1.SelectedItem.ListSubItems(61)
TextBox53.Text = ListView1.SelectedItem.ListSubItems(52)
TextBox61.Text = ListView1.SelectedItem.ListSubItems(60)
TextBox82.Text = ListView1.SelectedItem.ListSubItems(65)
TextBox82.Text = ListView1.SelectedItem.ListSubItems(65)
TextBox83.Text = ListView1.SelectedItem.ListSubItems(66)
TextBox84.Text = ListView1.SelectedItem.ListSubItems(67)
TextBox85.Text = ListView1.SelectedItem.ListSubItems(68)
TextBox86.Text = ListView1.SelectedItem.ListSubItems(69)
TextBox50.Text = ListView1.SelectedItem.ListSubItems(49)
TextBox63.Text = ListView1.SelectedItem.ListSubItems(62)


























'TextBox72.Text = ListView1.SelectedItem.ListSubItems(4)
'TextBox58.Text = ComboBox3.Value
'TextBox100.Text = ComboBox3.Column(1)
End Sub

Private Sub ListView1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As stdole.OLE_XPOS_PIXELS, ByVal Y As stdole.OLE_YPOS_PIXELS)

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

Private Sub TextBox1_Change()
Call arb
If TextBox1.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox1.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox1.Text = ""
TextBox1.SetFocus
End If
End If
 
End Sub



Private Sub TextBox10_Change()

End Sub

Private Sub TextBox101_Change()

End Sub

Private Sub TextBox102_Change()

End Sub

Private Sub TextBox104_Change()

End Sub

Private Sub TextBox11_AfterUpdate()

If TextBox49.value = 0 Or TextBox58.Text = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Or TextBox58.Text = "≈Ã«“… „⁄Ì· «·„ ›—€" _
Or TextBox58.Text = "—›⁄ Ìœ «Ê ≈Ìﬁ«› ’—›" Or TextBox58.Text = "«„Ê„Â «Ê· 6 «‘Â—" Or TextBox58.Text = "«„Ê„Â À«‰Ì 6 «‘Â—" Or TextBox58.Text = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" Then
TextBox11.value = 0
TextBox12.value = 30
TextBox47.value = 0
MsgBox "·« Ì„ﬂ‰ «œŒ«·  ﬁÌ„… ··€Ì«» ·«‰Â «·„‰ ”» Œ«—Ã „Êﬁ› «·Œœ„… Ê·Â–« ”Ì„‰⁄ «·‰Ÿ«„ „‰ «Õ ”«» ﬁÌ„… «·€Ì«»", vbCritical, "—”«·…  ‰»ÌÂ"
Else

Call true_salary_1

End If
End Sub

Private Sub TextBox11_Change()

If TextBox49.value = 0 Or TextBox58.Text = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Or TextBox58.Text = "≈Ã«“… „⁄Ì· «·„ ›—€" Or TextBox58.Text = "" _
Or TextBox58.Text = "—›⁄ Ìœ" Or TextBox58.Text = "«„Ê„Â «Ê· 6 «‘Â—" Or TextBox58.Text = "«„Ê„Â À«‰Ì 6 «‘Â—" Or TextBox58.Text = "≈Ã«“… ·Ã«‰ ÿ»ÌÂ" Then
TextBox11.value = 0
TextBox12.value = 30
TextBox47.value = 0
'MsgBox "·« Ì„ﬂ‰ «œŒ«·  ﬁÌ„… ··€Ì«» ·«‰ «·„‰ ”» Œ«—Ã „Êﬁ› «·Œœ„… Ê·Â–« ”Ì„‰⁄ «·‰Ÿ«„ „‰ «Õ ”«» ﬁÌ„… «·€Ì«»", vbCritical, "—”«·…  ‰»ÌÂ"

Else

Call true_salary_1

End If
End Sub

Private Sub TextBox12_Change()


End Sub

Private Sub TextBox13_Change()


End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox15_Change()

End Sub

Private Sub TextBox16_Change()

End Sub

Private Sub TextBox17_Change()


End Sub

Private Sub TextBox18_Change()

End Sub

Private Sub TextBox19_Change()


End Sub

Private Sub TextBox2_Change()


End Sub

Private Sub TextBox20_Change()

End Sub

Private Sub TextBox21_Change()

End Sub

Private Sub TextBox22_Change()

Dim numlong As Long
Dim numlong1 As Long
numlong = 10000
TextBox20.value = numlong1
If TextBox22.value >= 5 Then
numlong1 = 40000
TextBox20.value = numlong1
ElseIf TextBox22.value <= 4 Then
TextBox20.value = val(TextBox22.value) * numlong
End If
End Sub

Private Sub TextBox23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next

ComboBox1.Visible = True
ComboBox1.List = Sheets(4).Range("f1:f11").value
If IsNumeric(TextBox23.Text) Then
MsgBox "Â–« «·Õﬁ· ·« Ìﬁ»· «·« «·Õ—Ê› «·‰’ÌÂ"
TextBox23.Text = ""
TextBox23.SetFocus
Else
Exit Sub
End If
End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub TextBox25_Change()

If TextBox25.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox25.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox25.Text = ""
TextBox25.SetFocus
End If
End If
End Sub

Private Sub TextBox26_Change()
If TextBox26.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox26.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox26.Text = ""
TextBox26.SetFocus
End If
End If
End Sub

Private Sub TextBox28_Change()
If TextBox28.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox28.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox28.Text = ""
TextBox28.SetFocus
End If
End If
End Sub

Private Sub TextBox27_Change()

End Sub

Private Sub TextBox29_Change()

End Sub

Private Sub TextBox3_Change()

If TextBox3.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox3.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox3.Text = ""
TextBox3.SetFocus
End If
End If
End Sub

Private Sub TextBox30_Change()

End Sub

Private Sub TextBox31_Change()

End Sub

Private Sub TextBox32_Change()

End Sub

Private Sub TextBox33_Change()

End Sub

Private Sub TextBox34_Change()

End Sub

Private Sub TextBox35_Change()

If TextBox35.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox35.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox35.Text = ""
TextBox35.SetFocus
End If
End If
End Sub

Private Sub TextBox36_Change()

If TextBox36.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox36.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox36.Text = ""
TextBox36.SetFocus
End If
End If
End Sub

Private Sub TextBox37_Change()

End Sub

Private Sub TextBox38_Change()

If TextBox38.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox38.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox38.Text = ""
TextBox38.SetFocus
End If
End If
End Sub

Private Sub TextBox4_Change()
arb
If IsNumeric(TextBox4.Text) Then
MsgBox "Â–« «·Õﬁ· ·« Ìﬁ»· «·« «·Õ—Ê› «·‰’ÌÂ"
TextBox4.Text = ""

Else
Exit Sub
End If

TextBox4.SetFocus

End Sub

Private Sub TextBox39_Change()

End Sub

Private Sub TextBox40_Change()

End Sub

Private Sub TextBox41_Change()

End Sub

Private Sub TextBox42_Change()

If TextBox42.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox42.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox42.Text = ""
TextBox42.SetFocus
End If
End If
End Sub

Private Sub TextBox43_Change()

End Sub

Private Sub TextBox44_Change()

If TextBox44.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox44.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox44.Text = ""
TextBox44.SetFocus
End If
End If
End Sub

Private Sub TextBox45_Change()

End Sub

Private Sub TextBox46_Change()

End Sub

Private Sub TextBox47_Change()

End Sub

Private Sub TextBox48_Change()

End Sub

Private Sub TextBox49_Change()

End Sub

Private Sub TextBox5_Change()
arb
End Sub

Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If IsNumeric(TextBox5.Text) Then
MsgBox "Â–« «·Õﬁ· ·« Ìﬁ»· «·« «·Õ—Ê› «·‰’ÌÂ"
TextBox5.Text = ""
TextBox5.SetFocus
Else
Exit Sub
End If
TextBox27.value = ""
ListBox1.Clear
End Sub



Private Sub TextBox6_Change()
arb
If IsNumeric(TextBox6.Text) Then
MsgBox "Â–« «·Õﬁ· ·« Ìﬁ»· «·« «·Õ—Ê› «·‰’ÌÂ"
TextBox6.Text = ""
TextBox6.SetFocus
Else
Exit Sub
End If
End Sub



Private Sub TextBox52_Change()

End Sub

Private Sub TextBox53_Change()

End Sub

Private Sub TextBox54_Change()

End Sub

Private Sub TextBox55_Change()

End Sub

Private Sub TextBox56_Change()

End Sub

Private Sub TextBox57_Change()

End Sub

Private Sub TextBox58_AfterUpdate()







On Error Resume Next





With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If TextBox58.Text = Sheets(1).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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

UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")

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
   End With

End Sub

Private Sub TextBox58_Change()





On Error Resume Next




With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If TextBox58.Text = Sheets(1).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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

UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")

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
   End With
   TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("bf9:bf" & last1), Me.TextBox58.Text)

End Sub

Private Sub TextBox59_Change()

End Sub

Private Sub TextBox60_Change()

End Sub



Private Sub TextBox62_Change()

End Sub

Private Sub TextBox65_AfterUpdate()

If TextBox72.Text = "#N/A" Then

 Beep
 Beep
Application.Speech.Speak " al    rakum   gair     mau juod"
End If
End Sub

Private Sub TextBox65_Change()
On Error GoTo emad
Sheets(1).Cells(3, 64) = TextBox65.Text
TextBox72.Text = Sheets(1).Cells(3, 65).Text
ListBox1.Text = TextBox72.Text
If ListBox1.Text <> "" Then

2 Beep
Application.Speech.Speak " al    rakum       mau juod"

End If
emad:
End Sub

Private Sub TextBox65_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox65.Text = ""
TextBox77.Text = ""
TextBox72.Text = ""

End Sub

Private Sub TextBox66_Change()

End Sub

Private Sub TextBox67_Change()
Dim item1 As ListItem
Dim last1, frw1 As Integer
Dim rngData As Range
Dim wksSource As Worksheet
 Dim fString As Variant
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row



'=============================================

   
    
    
    For frw1 = 9 To last1
                
  If Sheets(1).Cells(frw1, "e") Like "*" & TextBox67.Text & "*" Then
            
   Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
   
   
   
'==================================================

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
'item1.SubItems(63) = Sheets(1).Cells(frw1, "bl")
'item1.SubItems(64) = Sheets(1).Cells(frw1, "bm")

End If
Next frw1
End With
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("e9:e8000"), "*" & Me.TextBox67.Text & "*", Sheets(1).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*")

End Sub

Private Sub TextBox67_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo emad
Call rightToLeftListView
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
ListView1.ListItems.Clear
With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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
item1.SubItems(55) = Sheets(1).Cells(frw1, "bf")




Next frw1

UserForm1.TextBox62.value = Format(UserForm1.TextBox62.value, "#,## IQD")
UserForm1.TextBox63.value = Format(UserForm1.TextBox63.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
End With


'==============================================
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
    
'=================================================




End With

Me.TextBox61.value = ListView1.ListItems.count
 CheckBox1.Locked = False
CommandButton5.Enabled = False

ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
emad:
End Sub

Private Sub TextBox7_Change()
arb

End Sub



Private Sub TextBox72_Change()
On Error Resume Next
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

 arb



ListBox1.Clear

Sheets(1).Activate
    For i = 1 To 63
            Controls("TextBox" & i).Text = ""
    Next i

    If TextBox72 = "" Then Exit Sub
    Sheets(1).Activate

    ss = Sheets(1).Cells(Rows.count, 5).End(xlUp).row
     k = 0
     
     
For Each c In Range("E8:E" & ss)
    If c Like TextBox72.value & "*" Then
        ListBox1.AddItem
        ListBox1.List(k, 0) = Cells(c.row, 5).value
        ListBox1.List(k, 1) = c.row
        k = k + 1
    End If
    
Next c
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Private Sub TextBox72_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets(1).Activate
On Error Resume Next
ComboBox15.Text = "«·»ÕÀ ›Ì «·„Êﬁ› «·ÊŸÌ›Ì"
TextBox102.value = Format(Sheets(26).Cells(3, "t"), "00.0")
If TextBox102.value = 0 Then
 Label122.Caption = "«·ﬁÌ„ „ ÿ«»ﬁ…"
 Else
 Label122.Caption = "«·ﬁÌ„ €Ì— „ ÿ«»ﬁ…"

End If

TextBox72.Text = ""


For i = 9 To Sheets(1).Cells(Rows.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 9, 0) = Cells(i, 5).value
  ListBox1.List(i - 9, 1) = i
      Next i
      If TextBox22.value = "" Then
TextBox20.value = ""
End If

Dim last1 As Integer
last1 = Sheets(1).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last1 - 8

End Sub



Private Sub TextBox72_Enter()
arb

End Sub

Private Sub TextBox72_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

End Sub

Private Sub TextBox75_Change()

End Sub

Private Sub TextBox8_Change()
On Error Resume Next
Sheets(9).Cells(18, "g") = TextBox8.Text
Sheets(9).Cells(18, "f") = TextBox9.Text
TextBox10.Text = Sheets(9).Cells(18, "h")

If TextBox8.Text = "" Then
Exit Sub
Else
If IsNumeric(TextBox8.Text) Then
Else
MsgBox "Â–« «·Õﬁ· ·« ÌÌﬁ»· «·« «·«—ﬁ«„"
TextBox8.Text = ""
TextBox8.SetFocus
End If
End If

Sheets(9).Cells(17, "q") = TextBox8.Text
Sheets(9).Cells(17, "r") = TextBox9.Text
TextBox91.Text = Sheets(9).Cells(17, "s")

 Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet10")

    '  ÕœÌÀ «·Ê’› «·ÊŸÌ›Ì
    ws.Range("D69").value = Me.TextBox106.value

    '  ÕœÌÀ «·œ—Ã…
    ws.Range("C69").value = Me.TextBox8.value

    '  ÕœÌÀ ComboBox5 »«·ﬁÌ„… «·„ÊÃÊœ… ›Ì E69
    Me.ComboBox5.value = ws.Range("E69").value






End Sub

Private Sub TextBox85_AfterUpdate()
TextBox89.value = Format(val(TextBox85) / val(TextBox10) * 100, "00.0") & "%"
End Sub

Private Sub TextBox85_Change()
On Error Resume Next
TextBox89.value = Format(val(TextBox85) / val(TextBox10) * 100, "00.0") & "%"

End Sub

Private Sub TextBox86_AfterUpdate()
On Error Resume Next
TextBox90.value = Format(val(TextBox86) / val(TextBox10) * 100, "00.0") & "%"
End Sub

Private Sub TextBox86_Change()
On Error Resume Next
TextBox90.value = Format(val(TextBox86) / val(TextBox10) * 100, "00.0") & "%"
End Sub

Private Sub TextBox89_Change()

End Sub

Private Sub TextBox9_Change()
On Error Resume Next

Sheets(9).Cells(18, "g") = TextBox8.Text
Sheets(9).Cells(18, "f") = TextBox9.Text
TextBox10.Text = Sheets(9).Cells(18, "h")

Sheets(9).Cells(17, "q") = TextBox8.Text
Sheets(9).Cells(17, "r") = TextBox9.Text
TextBox91.Text = Sheets(9).Cells(17, "s")
TextBox10.Text = Sheets(9).Cells(18, "h")








If TextBox9.Text = "" Then
Exit Sub
End If
End Sub
Private Sub UserForm_Activate()

On Error GoTo emad
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Sheets(1).Activate

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)
ComboBox1.List = Sheets(4).Range("f1:f11").value
ComboBox2.List = Sheets(5).Range("a1:a373").value
ComboBox3.List = Sheets(4).Range("o1:p13").value
ComboBox15.List = Sheets(4).Range("o1:p13").value
ComboBox4.List = Array("–ﬂ—", "«‰ÀÏ")
ComboBox14.List = Array("ŒÿÊ—… ”«∆ﬁ", "·« Ì” Õﬁ")

ComboBox5.List = Sheets(5).Range("a1:a373").value

ComboBox6.List = Array("0", "0.15", "0.25", "0.35", "0.45", "0.55", "0.65", "0.75", " 0.85", "0.95", "1", "1.05", "1.10", "1.15", "1.20", "1.25", "1.30", "1.35", "1.40", "1.45", "1.50")

ComboBox7.List = Array("0.35", "0.5", "0")
ComboBox8.List = Array("0.5", "0.3", "0.25", "0.20", "0.15", "0")
ComboBox9.List = Array("„«” —", "ﬂ«‘")
ComboBox10.List = Array("0.3", "0.25", "0.20", "0.15", "0")
ComboBox11.List = Array("1", "0.80", "0.30", "0.25", "0.20", "0.15", "0")
ComboBox12.List = Array("ﬂ«‰Ê‰ «·À«‰Ì", "‘»«ÿ", "¬–«—", "‰Ì”«‰", "√Ì«—", "Õ“Ì—«‰", " „Ê“", "¬»", "√Ì·Ê·", " ‘—Ì‰†«·√Ê·", " ‘—Ì‰†«·À«‰Ì", "ﬂ«‰Ê‰†«·√Ê·")
ComboBox13.List = Array("ﬂ«‰Ê‰ «·À«‰Ì", "‘»«ÿ", "¬–«—", "‰Ì”«‰", "√Ì«—", "Õ“Ì—«‰", " „Ê“", "¬»", "√Ì·Ê·", " ‘—Ì‰†«·√Ê·", " ‘—Ì‰†«·À«‰Ì", "ﬂ«‰Ê‰†«·√Ê·")







Me.TextBox67.Text = ""
Me.TextBox92.Text = Sheets(1).Cells(1, "cb")
'========================================================= «·Õ”«»«  «·ﬂ·ÌÂ Õ”»  ›«’Ì·Â«
 TextBox76.Text = Sheets(1).Cells(2, "cc")
TextBox73.Text = Sheets(1).Cells(2, "cd")
TextBox74.Text = Sheets(1).Cells(2, "ce")
TextBox79.Text = Sheets(1).Cells(2, "cf")
TextBox80.Text = Sheets(1).Cells(2, "cg")
TextBox82.Text = Sheets(1).Cells(1, "bn")
TextBox66.Text = Sheets(1).Cells(2, "ch")
TextBox107.Text = Sheets(1).Cells(3, "cl")


'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")









TextBox102.value = Format(Sheets(26).Cells(3, "t"), "00.0")
If TextBox102.value = 0 Then
 Label122.Caption = "«·ﬁÌ„ „ ÿ«»ﬁ…"
 Else
 Label122.Caption = "«·ﬁÌ„ €Ì— „ ÿ«»ﬁ…"

End If




Dim lastr1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, "b").End(xlUp).row

Ê—ﬁ…1.Range("cb9").Formula = "=TEXT(TODAY(),""mmm"")"
Ê—ﬁ…1.Range("cb9:cb" & lastr1).FillDown
Dim last1 As Integer
last1 = Sheets(1).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last1 - 8

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

Call CalculateSalaryPositions
Call UpdateLiveFormulas_Sheet1
Call FillAT
emad:
End Sub





Private Sub UserForm_Click()
Dim last1 As Integer
last1 = Sheets(1).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last1 - 8
Call CaptureSnapshot(Me, DictOld)
 Label127.Visible = True
    ' ≈⁄«œ… «··Ê‰ «·√’·Ì ﬁ»· «· „œœ
    Label127.BackColor = OrigColor
    
    AnimateLabel OrigW, OrigH
End Sub

Private Sub UserForm_Initialize()
'===============================
On Error Resume Next

 OrigW = Label127.Width
    OrigH = Label127.Height
    OrigColor = Label127.BackColor


 Dim ws As Worksheet
    Me.ComboBox16.Clear
    
    ' ≈÷«›… √”„«¡ ‘Ì «  «·√‘Â—  ·ﬁ«∆Ì«
    For Each ws In ThisWorkbook.Worksheets
        Select Case LCase(ws.Name)
            Case "jan", "feb", "mar", "apr", "may", "jun", _
                 "jul", "aug", "sep", "oct", "nov", "dec"
                Me.ComboBox16.AddItem ws.Name
        End Select
    Next ws










Label104.Caption = Sheets("title_factory").Cells(2, 1)

Label105.Caption = Sheets("title_factory").Cells(2, 3)
On Error GoTo ResetCursor ' ????? ????? ?????? ??? ??? ?????? ??????? ?? ??? ???? ???
       Application.Cursor = xlWait ' ????? ?????? ???????
    Application.ScreenUpdating = False ' ????? ???????
    
    ' ??? ???????? ?????? ?? ??? (????? ????? ??????? ???? 2025)
   

Label124.Enabled = False
TextBox101.Visible = False
If Me.Width = Application.Width Then

Me.Width = 1010.5
Me.Height = 720
Me.Zoom = 49



Else
Me.Width = Application.Width
Me.Height = Application.Height
X = Me.Width / 1010.5
Me.Zoom = Me.Zoom * X

End If
Me.Left = (Application.Width - Me.Width) / 2
Me.Top = (Application.Height - Me.Height) / 2



'=========================
Call rightToLeftListView
With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True

.ColumnHeaders.Add , , "«·—ﬁ„", 40
.ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
.ColumnHeaders.Add , , "—ﬁ„ «·„«” —", 120
.ColumnHeaders.Add , , "«·Ã‰”", 40
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 210
.ColumnHeaders.Add , , "«·⁄‰Ê«‰ «·ÊŸÌ›Ì", 100
.ColumnHeaders.Add , , " «·ﬁ”„", 100
.ColumnHeaders.Add , , " «·œ—Ã…", 50
.ColumnHeaders.Add , , "«·„—Õ·…", 50
.ColumnHeaders.Add , , "«·—« » «·«”„Ì ", 140
.ColumnHeaders.Add , , "«Ì«„ «·€Ì«» ", 140
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
.ColumnHeaders.Add , , "‰”»… „Â‰ÌÂ ", 100
.ColumnHeaders.Add , , "  „Œ’’«  „Â‰ÌÂ ", 100
.ColumnHeaders.Add , , "‰”»… «·ŒÿÊ—Â", 100
.ColumnHeaders.Add , , " „Œ’’«  «·ŒÿÊ—… ", 140
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·„Œ’’« ", 120
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·«” Õﬁ«ﬁ« ", 120
.ColumnHeaders.Add , , "  Êﬁ›«   ﬁ«⁄œÌÂ 10", 100
.ColumnHeaders.Add , , "« Êﬁ›«   ﬁ«⁄œÌÂ 15%", 40
.ColumnHeaders.Add , , "«· «„Ì‰ «·’ÕÌ ", 50
.ColumnHeaders.Add , , " «ÃÊ— «·‰ﬁ·", 40
.ColumnHeaders.Add , , " ‰›ﬁ…", 40
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄", 40

.ColumnHeaders.Add , , " —”„ «·ÿ«»⁄", 100
.ColumnHeaders.Add , , " Ê“«—… «·„«·ÌÂ", 50
.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄1", 50
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄1  ", 140
.ColumnHeaders.Add , , "‰«œÌ «·ÂÌ√… «·’‰«⁄«  «·Õ—»Ì", 40
.ColumnHeaders.Add , , " ÃÂ… «·«” ﬁÿ«⁄ ", 120
.ColumnHeaders.Add , , " ‰«œÌ «·ÂÌ√… «·’‰«⁄«  «·Õ—»Ì", 120
.ColumnHeaders.Add , , "ÃÂ… «·«” ﬁÿ«⁄ ", 40
.ColumnHeaders.Add , , "«” ﬁÿ«⁄«  —⁄«Ì… ", 140
.ColumnHeaders.Add , , "„»·€ «·€Ì«» ", 40
.ColumnHeaders.Add , , "„Ã„Ê⁄ «·«” ﬁÿ«⁄« ", 100
.ColumnHeaders.Add , , " «·—« » «·’«›Ì", 50
.ColumnHeaders.Add , , " «·”·›…", 50
.ColumnHeaders.Add , , "ÿ—Ìﬁ… «·œ›⁄", 150
.ColumnHeaders.Add , , " —ﬁ„ «·«Ì»«‰ ", 140
.ColumnHeaders.Add , , "  «—ÌŒ «·«÷«›… ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ’Ê·Â ⁄·Ï ⁄·«Ê… ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ’Ê·Â ⁄·Ï  —›Ì⁄  ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ–›Â Ê —ÕÌ·Â  ", 140

.ColumnHeaders.Add , , "«·ÃÂ… «·„«‰Õ… ··—« » ", 140
.ColumnHeaders.Add , , "«·„Êﬁ› «·ÊŸÌ›Ì ", 140
.ColumnHeaders.Add , , " «—ÌŒ Õ’Ê·Â ⁄·Ï ﬂ «» ‘ﬂ— Ê ﬁœÌ—", 140


.ColumnHeaders.Add , , " «·«÷«›… «·„«·Ì… «·„ƒﬁ Â ", 140
.ColumnHeaders.Add , , "  «—ÌŒ «· ⁄œÌ· ", 140

.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ  ", 140
.ColumnHeaders.Add , , " —„“ «·«” Õﬁ«ﬁ", 140
.ColumnHeaders.Add , , " „”«— «·’Ê—…", 140
.ColumnHeaders.Add , , " barcode ", 140
.ColumnHeaders.Add , , "  «—ÌŒ «Œ—  ÕœÌÀ ", 140
.ColumnHeaders.Add , , " «” ﬁÿ«⁄ ⁄«„ ", 140
.ColumnHeaders.Add , , " «” —Ã«⁄ ⁄«„ ", 140
.ColumnHeaders.Add , , " «” ﬁÿ«⁄ „»·€ ‰”»… ", 140
.ColumnHeaders.Add , , " «÷«›… „»·€ ‰”»… ", 140
.ColumnHeaders.Add , , " ”»» «·«” ﬁÿ«⁄", 140
.ColumnHeaders.Add , , " ”»» «·«” —Ã«⁄ ", 140




ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With

'==========================================================================
Call rightToLeftListView








'=================================================
TextBox61.Text = ""
TextBox62.Text = ""
TextBox63.Text = ""
TextBox64.Text = ""
TextBox80.Text = ""
TextBox66.Text = ""


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 14 'ÕÃ„ «·Œÿ
.Font.Bold = True  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
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

item1.SubItems(63) = Sheets(1).Cells(frw1, "bl")
item1.SubItems(64) = Sheets(1).Cells(frw1, "bm")
item1.SubItems(65) = Sheets(1).Cells(frw1, "bn")
item1.SubItems(66) = Sheets(1).Cells(frw1, "bo")
item1.SubItems(67) = Sheets(1).Cells(frw1, "bp")
item1.SubItems(68) = Sheets(1).Cells(frw1, "bq")
item1.SubItems(69) = Sheets(1).Cells(frw1, "br")
item1.SubItems(70) = Sheets(1).Cells(frw1, "bs")
item1.SubItems(71) = Sheets(1).Cells(frw1, "bt")





Next frw1

UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")

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
Me.TextBox75.value = ListView1.ListItems.count
 CheckBox1.Locked = False
CommandButton5.Enabled = False

ListView1.Font = Bold
ListView1.Font.Size = 14

'ListView1.ForeColor = "blue"
 If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If

Label123.Visible = False


Dim last As Integer
last = Sheets(1).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last - 8


    
 
    
ResetCursor:
    Application.Cursor = xlDefault ' ????? ?????? ????? ???????
    Application.ScreenUpdating = True


End Sub


Private Sub GetMonthDataToSheet1(monthSheetName As String)

    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, LastCol As Long
    Dim dataRange As Range, dataArr As Variant, headersArr As Variant
    Dim nRows As Long, nCols As Long
    Dim actualCells As Long

    Set wsSrc = Sheets(monthSheetName)
    Set wsDest = Sheets("Sheet1")

    '  ÕœÌœ ¬Œ— ’› Ê¬Œ— ⁄„Êœ
    lastRow = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).row
    LastCol = wsSrc.Cells(8, wsSrc.Columns.count).End(xlToLeft).Column

    ' ===== «· Õﬁﬁ „‰ ÊÃÊœ √Ì »Ì«‰«  =====
    If lastRow < 9 Then
        MsgBox "«·‘Â— [" & monthSheetName & "] ›«—€ Ê·« ÌÕ ÊÌ ⁄·Ï √Ì »Ì«‰« ." & vbCrLf & _
               " √ﬂœ „‰  —ÕÌ· «·»Ì«‰«  ≈·ÌÂ √Ê·«. ‘ﬂ—« · ⁄«Ê‰ﬂ„.", vbExclamation, " ‰»ÌÂ"
        Exit Sub
    End If

    Set dataRange = wsSrc.Range(wsSrc.Cells(9, 1), wsSrc.Cells(lastRow, LastCol))
    
    ' ⁄œ √Ì Œ·Ì…  Õ ÊÌ ⁄·Ï »Ì«‰« 
    actualCells = Application.WorksheetFunction.CountA(dataRange)
    If actualCells = 0 Then
        MsgBox "«·‘Â— [" & monthSheetName & "] ›«—€  „«„« Ê·« ÌÕ ÊÌ ⁄·Ï √Ì »Ì«‰« ." & vbCrLf & _
               " √ﬂœ „‰  —ÕÌ· «·»Ì«‰«  ≈·ÌÂ √Ê·«. ‘ﬂ—« · ⁄«Ê‰ﬂ„.", vbExclamation, " ‰»ÌÂ"
        Exit Sub
    End If

    ' ﬁ—«¡… «·ÂÌœ— Ê«·»Ì«‰« 
    headersArr = wsSrc.Rows(8).Resize(1, LastCol).value
    dataArr = dataRange.value
    nRows = UBound(dataArr, 1)
    nCols = UBound(dataArr, 2)

    ' ‰”Œ «·ÂÌœ— ≈·Ï Sheet1 ›Ì ’› 8
    wsDest.Cells(8, 1).Resize(1, nCols).value = headersArr

    ' ‰”Œ «·»Ì«‰«  ≈·Ï Sheet1 »œ¡« „‰ ’› 9
    wsDest.Cells(9, 1).Resize(nRows, nCols).value = dataArr

End Sub
Sub UpdateCGInSheet(EmpID As String, Status As Boolean)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    '  ÕœÌœ «·‘Ì  («· ”·”· 1)
    Set ws = ThisWorkbook.Sheets(1)
    
    ' ≈ÌÃ«œ ¬Œ— ’› ›Ì «·⁄„Êœ B («·—ﬁ„ «·ÊŸÌ›Ì)
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    ' Õ·ﬁ…  ﬂ—«—Ì…  »œ√ „‰ «·’› 9
    For i = 9 To lastRow
        ' „ﬁ«—‰… «·—ﬁ„ «·ÊŸÌ›Ì „⁄ «·ﬁÌ„… «·„ÊÃÊœ… ›Ì «·‘Ì 
        If Trim(ws.Cells(i, "B").value) = Trim(EmpID) Then
            '  ÕœÌÀ «·Œ·Ì… ›Ì «·⁄„Êœ CG »«·ﬁÌ„… «·„ÿ·Ê»… (True √Ê False)
            ws.Cells(i, "CG").value = Status
            Exit For ' «·Œ—ÊÃ „‰ «·Õ·ﬁ… »⁄œ ≈ „«„ «· ÕœÌÀ
        End If
    Next i
End Sub

Sub AnimateLabel(TargetW As Single, TargetH As Single)
    Dim i As Integer
    ' “Ì«œ… ⁄œœ «·ŒÿÊ«   Ã⁄· «·Õ—ﬂ… √‰⁄„ Ê√»ÿ√ (ﬂ«‰  15 √’»Õ  40)
    Dim Steps As Integer: Steps = 40
    Dim DiffW As Single, DiffH As Single
    
    DiffW = (TargetW - Label127.Width) / Steps
    DiffH = (TargetH - Label127.Height) / Steps

    For i = 1 To Steps
        Label127.Width = Label127.Width + DiffW
        Label127.Height = Label127.Height + DiffH
        
        DoEvents
        ' “Ì«œ… «· √ŒÌ— «·“„‰Ì (0.015 À«‰Ì…) · ﬁ·Ì· «·”—⁄…
        Dim t As Single: t = Timer: Do Until Timer > t + 0.015: Loop
    Next i
    
    Label127.Width = TargetW
    Label127.Height = TargetH
End Sub
