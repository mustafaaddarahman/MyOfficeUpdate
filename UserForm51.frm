VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm51 
   Caption         =   "UserForm51"
   ClientHeight    =   13815
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   20160
   OleObjectBlob   =   "UserForm51.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm51"
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

If val(TextBox27) > 0 And val(TextBox29) > 0 Then
TextBox29 = 0
ComboBox11.value = 0
'MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
TextBox29 = TextBox29
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

If val(TextBox27) > 0 And val(TextBox29) > 0 Then
TextBox29 = 0
ComboBox11.value = 0
'MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
TextBox29 = TextBox29
End If
'===============================================

End Sub

Private Sub ComboBox11_AfterUpdate()

If val(TextBox27) > 0 And val(TextBox29) > 0 Then
TextBox27 = 0
ComboBox10.value = 0
MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
TextBox27 = TextBox27
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

If val(TextBox27) > 0 And val(TextBox29) > 0 Then
TextBox27 = 0
ComboBox10.value = 0
MsgBox "·« Ì„ﬂ‰ Ã„⁄ ŒÿÊ—… „Â‰ÌÂ „⁄ ŒÿÊ—… ‰”»Â «·„Œ’’Â ·«Œ ’«’«  «·„⁄‰Ê‰Â „«·Ì«", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
Else
TextBox27 = TextBox27
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



ComboBox3.value = ""

Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox15.value

If Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox15.value Then

TextBox93.value = Sheets(ComboBox16.value).Cells(4, "cd")
TextBox94.value = Sheets(ComboBox16.value).Cells(4, "ce")
TextBox97.value = Sheets(ComboBox16.value).Cells(4, "cf")
TextBox98.value = Sheets(ComboBox16.value).Cells(4, "cg")
TextBox95.value = Sheets(ComboBox16.value).Cells(4, "ch")
TextBox96.value = Sheets(ComboBox16.value).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
TextBox93.value = Format(TextBox93.value, "#,## IQD")
TextBox94.value = Format(TextBox94.value, "#,## IQD")
TextBox97.value = Format(TextBox97.value, "#,## IQD")
TextBox98.value = Format(TextBox98.value, "#,## IQD")
TextBox95.value = Format(TextBox95.value, "#,## IQD")
TextBox96.value = Format(TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & ComboBox15.value
End If
'===========================================
ListBox1.Clear

TextBox72.Text = ""
Sheets(ComboBox16.value).Activate
Dim frow1 As Integer
     ss = Sheets(ComboBox16.value).Cells(Rows.count, 5).End(xlUp).row
     
     For frow1 = 9 To ss
    If ComboBox15.value = Sheets(ComboBox16.value).Cells(frow1, 58) Then
        ListBox1.AddItem
        ListBox1.List(ListBox1.ListCount - 1, 0) = Cells(frow1, 5).value
       
    End If
    
Next
Dim last2 As Integer
     last2 = Sheets(ComboBox16.value).Cells(Rows.count, 5).End(xlUp).row
     
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(ComboBox16.value).Range("bf9:bf" & last2), Me.ComboBox15.Text)


'================================ Õ„Ì· «·»Ì«‰«  ›Ì «· listview

With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If ComboBox15.value = Sheets(ComboBox16.value).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")

End If



Next frw1

TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox64.value = Format(TextBox64.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")

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
On Error Resume Next
ComboBox3.value = ""

Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox15.value

If Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox15.value Then

TextBox93.value = Sheets(ComboBox16.value).Cells(4, "cd")
TextBox94.value = Sheets(ComboBox16.value).Cells(4, "ce")
TextBox97.value = Sheets(ComboBox16.value).Cells(4, "cf")
TextBox98.value = Sheets(ComboBox16.value).Cells(4, "cg")
TextBox95.value = Sheets(ComboBox16.value).Cells(4, "ch")
TextBox96.value = Sheets(ComboBox16.value).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
TextBox93.value = Format(TextBox93.value, "#,## IQD")
TextBox94.value = Format(TextBox94.value, "#,## IQD")
TextBox97.value = Format(TextBox97.value, "#,## IQD")
TextBox98.value = Format(TextBox98.value, "#,## IQD")
TextBox95.value = Format(TextBox95.value, "#,## IQD")
TextBox96.value = Format(TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & ComboBox15.value
End If

'===================================
ListBox1.Clear

Sheets(ComboBox16.value).Activate
Dim frow1 As Integer
     ss = Sheets(ComboBox16.value).Cells(Rows.count, 5).End(xlUp).row
     
     For frow1 = 9 To ss
    If ComboBox15.value = Sheets(ComboBox16.value).Cells(frow1, 58) Then
        ListBox1.AddItem
        ListBox1.List(ListBox1.ListCount - 1, 0) = Cells(frow1, 5).value
       
    End If
    
Next
Dim last2 As Integer
     last2 = Sheets(ComboBox16.value).Cells(Rows.count, 5).End(xlUp).row
     
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(ComboBox16.value).Range("bf9:bf" & last2), Me.ComboBox15.Text)


'================================ Õ„Ì· «·»Ì«‰«  ›Ì «· listview

With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If ComboBox15.value = Sheets(ComboBox16.value).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")

End If



Next frw1

TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox64.value = Format(TextBox64.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")

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



Private Sub ComboBox16_Change()
 Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long, LastCol As Long
    Dim hasData As Boolean
    
    sheetName = Me.ComboBox16.value
    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Set ws = Sheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "«·‘Ì  [" & sheetName & "] €Ì— „ÊÃÊœ.", vbExclamation
        Exit Sub
    End If

    ' ¬Œ— ’› Ê¬Œ— ⁄„Êœ
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    '  Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  »⁄œ «·’› 8
    hasData = (lastRow >= 9) And (Application.WorksheetFunction.CountA(ws.Range(ws.Cells(9, 1), ws.Cells(lastRow, LastCol))) > 0)

    '  ›—Ì€ ListBox ›Ê—« ⁄‰œ  €Ì— «·„Õ ÊÏ
    ListBox1.Clear
    ListBox1.ColumnCount = LastCol

    If hasData Then
        MsgBox "«·„Õ ÊÏ ··‘Â— [" & sheetName & "] ÌÕ ÊÌ ⁄·Ï »Ì«‰« .", vbInformation
    Else
        ' «·‘Ì  ›«—€ ? «” œ⁄«¡ «· ÊÃÌÂ ··„’›Ê›…
        Call GetMonthDataToSheet1(sheetName)
    End If
    
    
   TextBox76.value = Sheets(ComboBox16.value).Cells(2, "cc")
TextBox73.value = Sheets(ComboBox16.value).Cells(2, "cd")
TextBox74.value = Sheets(ComboBox16.value).Cells(2, "ce")
TextBox79.Text = Sheets(ComboBox16.value).Cells(2, "cf")
TextBox80.Text = Sheets(ComboBox16.value).Cells(2, "cg")
TextBox82.Text = Sheets(ComboBox16.value).Cells(1, "bn")
TextBox66.Text = Sheets(ComboBox16.value).Cells(2, "ch")

 
    
End Sub

Private Sub ComboBox2_Change()
ListView1.ListItems.Clear
End Sub

Private Sub ComboBox3_AfterUpdate()
On Error Resume Next

Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox3.value
TextBox58.Text = ComboBox3.value
TextBox100.Text = ComboBox3.Column(1)

If Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox3.value Then

TextBox93.value = Sheets(ComboBox16.value).Cells(4, "cd")
TextBox94.value = Sheets(ComboBox16.value).Cells(4, "ce")
TextBox97.value = Sheets(ComboBox16.value).Cells(4, "cf")
TextBox98.value = Sheets(ComboBox16.value).Cells(4, "cg")
TextBox95.value = Sheets(ComboBox16.value).Cells(4, "ch")
TextBox96.value = Sheets(ComboBox16.value).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
TextBox93.value = Format(TextBox93.value, "#,## IQD")
TextBox94.value = Format(TextBox94.value, "#,## IQD")
TextBox97.value = Format(TextBox97.value, "#,## IQD")
TextBox98.value = Format(TextBox98.value, "#,## IQD")
TextBox95.value = Format(TextBox95.value, "#,## IQD")
TextBox96.value = Format(TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & TextBox58.Text
End If


End Sub

Private Sub ComboBox3_Change()
On Error Resume Next

TextBox58.Text = ComboBox3.value
TextBox100.Text = ComboBox3.Column(1)


Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox3.value
If Sheets(ComboBox16.value).Cells(4, "cb").value = ComboBox3.value Then

TextBox93.value = Sheets(ComboBox16.value).Cells(4, "cd")
TextBox94.value = Sheets(ComboBox16.value).Cells(4, "ce")
TextBox97.value = Sheets(ComboBox16.value).Cells(4, "cf")
TextBox98.value = Sheets(ComboBox16.value).Cells(4, "cg")
TextBox95.value = Sheets(ComboBox16.value).Cells(4, "ch")
TextBox96.value = Sheets(ComboBox16.value).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
TextBox93.value = Format(TextBox93.value, "#,## IQD")
TextBox94.value = Format(TextBox94.value, "#,## IQD")
TextBox97.value = Format(TextBox97.value, "#,## IQD")
TextBox98.value = Format(TextBox98.value, "#,## IQD")
TextBox95.value = Format(TextBox95.value, "#,## IQD")
TextBox96.value = Format(TextBox96.value, "#,## IQD")

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







Private Sub CommandButton18_Click()
On Error Resume Next
Call trans_data
Dim filnam, filpath As String
filnam = "reportxl" & "  " & Format(Date, "dd-mm")
filpath = ThisWorkbook.path & "\excel files\" & filnam & ".xlsx"
Dim ws As Worksheet
Set ws = Sheets(11)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
With ws
.Copy
Application.ActiveWorkbook.SaveAs fileName:=filnam
Application.ActiveWorkbook.Close False
End With
Application.ScreenUpdating = True
Application.DisplayAlerts = True
MsgBox " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "—”«·Â  ÊÃÌÂ"

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
Sheets(ComboBox16.value).Activate
Dim lastR As Integer
lastR = Sheets(ComboBox16.value).Cells(Rows.count, 1).End(xlUp).row

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
On Error Resume Next
Unload Me
UserForm9.Show


End Sub

Private Sub CommandButton24_Click()
'On Error GoTo emad
Dim lastR As Integer
Dim file_path As String
lastR = Sheets(ComboBox16.value).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(ComboBox16.value).Cells(Y, 5) = TextBox5.Text And Sheets(ComboBox16.value).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    
  
   
   'Call emad_pro

 '=========================================
 file_path = "D:\employ_pic\" & TextBox5.Text & ".jpg"
 
 TextBox64.Text = file_path
 SavePicture Image1.Picture, TextBox64.Text

Sheets(ComboBox16.value).Cells(Y, 64) = TextBox64.Text
MsgBox " „ Õ›Ÿ «·’Ê—…"
'emad:

End Sub



















Private Sub CommandButton32_Click()

End Sub

Private Sub CommandButton33_Click()
UserForm26.Show
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
Set rang1 = Sheets(ComboBox16.value).Range("zwjea4")
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











Private Sub CommandButton39_Click()
Dim rng As Range
Dim cel As Range
Dim lastr1 As Integer
lastr1 = Sheets(ComboBox16.value).Cells(Rows.count, "b").End(xlUp).row
Set rng = Sheets(ComboBox16.value).Range("au9:au" & lastr1)
For Each cel In rng
If IsEmpty(cel) Or Not IsNumeric(cel) Then
cel = 0
End If
Next cel
End Sub







Private Sub CommandButton41_Click()
Call SORT_A_TO_Z1
End Sub













Private Sub CommandButton8_Click()
Sheets(ComboBox16.value).Activate
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
Sheets(ComboBox16.value).Activate
last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row



'=============================================

   Sheets(ComboBox16.value).Activate
    
    
    For frw1 = 9 To last1
                
 If Sheets(ComboBox16.value).Cells(frw1, "h") = TextBox67.Text And Sheets(ComboBox16.value).Cells(frw1, "f") = ComboBox2.Text And ComboBox3.value = Sheets(ComboBox16.value).Cells(frw1, "bf").value Then
     
'==================================================

Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("j9:j8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("ad9:ad8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("ae9:ae8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("av9:av8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("aw9:aw8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("x9:x8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text, Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*")






TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox79.value = Format(TextBox79.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")
TextBox76.value = Format(TextBox76.value, "#,## IQD")




End If
Next frw1
End With

'============================================= «·Õ”«»«  «·ﬂ·ÌÂ

Dim simadd As Double
Dim simsubtract As Double
Dim item2 As ListItem
Dim last2, frw2 As Integer
last2 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2



If ComboBox2.value = "«·ﬂ·" And Sheets(ComboBox16.value).Cells(frw2, "bf") = ComboBox3.value Then

Set item2 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw2, "A"))




item2.SubItems(1) = Sheets(ComboBox16.value).Cells(frw2, "b")
item2.SubItems(2) = Sheets(ComboBox16.value).Cells(frw2, "c")
item2.SubItems(3) = Sheets(ComboBox16.value).Cells(frw2, "d")
item2.SubItems(4) = Sheets(ComboBox16.value).Cells(frw2, "e")
item2.SubItems(5) = Sheets(ComboBox16.value).Cells(frw2, "f")
item2.SubItems(6) = Sheets(ComboBox16.value).Cells(frw2, "g")
item2.SubItems(7) = Sheets(ComboBox16.value).Cells(frw2, "h")
item2.SubItems(8) = Sheets(ComboBox16.value).Cells(frw2, "i")
item2.SubItems(9) = Sheets(ComboBox16.value).Cells(frw2, "j")
item2.SubItems(10) = Sheets(ComboBox16.value).Cells(frw2, "k")
item2.SubItems(11) = Sheets(ComboBox16.value).Cells(frw2, "l")
item2.SubItems(12) = Sheets(ComboBox16.value).Cells(frw2, "m")
item2.SubItems(13) = Sheets(ComboBox16.value).Cells(frw2, "n")
item2.SubItems(14) = Sheets(ComboBox16.value).Cells(frw2, "o")
item2.SubItems(15) = Sheets(ComboBox16.value).Cells(frw2, "p")
item2.SubItems(16) = Sheets(ComboBox16.value).Cells(frw2, "q")
item2.SubItems(17) = Sheets(ComboBox16.value).Cells(frw2, "r")
item2.SubItems(18) = Sheets(ComboBox16.value).Cells(frw2, "s")
item2.SubItems(19) = Sheets(ComboBox16.value).Cells(frw2, "t")
item2.SubItems(20) = Sheets(ComboBox16.value).Cells(frw2, "u")
item2.SubItems(21) = Sheets(ComboBox16.value).Cells(frw2, "v")
item2.SubItems(22) = Sheets(ComboBox16.value).Cells(frw2, "w")
item2.SubItems(23) = Sheets(ComboBox16.value).Cells(frw2, "x")
item2.SubItems(24) = Sheets(ComboBox16.value).Cells(frw2, "y")
item2.SubItems(25) = Sheets(ComboBox16.value).Cells(frw2, "z")
item2.SubItems(26) = Sheets(ComboBox16.value).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(ComboBox16.value).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(ComboBox16.value).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(ComboBox16.value).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(ComboBox16.value).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(ComboBox16.value).Cells(frw2, "af")
item2.SubItems(32) = Sheets(ComboBox16.value).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(ComboBox16.value).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(ComboBox16.value).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(ComboBox16.value).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(ComboBox16.value).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(ComboBox16.value).Cells(frw2, "al")
item2.SubItems(38) = Sheets(ComboBox16.value).Cells(frw2, "am")
item2.SubItems(39) = Sheets(ComboBox16.value).Cells(frw2, "an")
item2.SubItems(40) = Sheets(ComboBox16.value).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(ComboBox16.value).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(ComboBox16.value).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(ComboBox16.value).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(ComboBox16.value).Cells(frw2, "as")
item2.SubItems(45) = Sheets(ComboBox16.value).Cells(frw2, "at")
item2.SubItems(46) = Sheets(ComboBox16.value).Cells(frw2, "au")
item2.SubItems(47) = Sheets(ComboBox16.value).Cells(frw2, "av")
item2.SubItems(48) = Sheets(ComboBox16.value).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(ComboBox16.value).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(ComboBox16.value).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(ComboBox16.value).Cells(frw2, "az")
item2.SubItems(52) = Sheets(ComboBox16.value).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(ComboBox16.value).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(ComboBox16.value).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(ComboBox16.value).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(ComboBox16.value).Cells(frw2, "be")
item2.SubItems(57) = Sheets(ComboBox16.value).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(ComboBox16.value).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(ComboBox16.value).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(ComboBox16.value).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(ComboBox16.value).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(ComboBox16.value).Cells(frw2, "bk")
'item2.SubItems(63) = Sheets(combobox16.value).Cells(frw2, "bl")
simadd = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("bh9:bh" & frw2))
simsubtract = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("bj9:bj" & last2))
TextBox73.Text = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("j9:j" & last2))
TextBox74.Text = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("ad9:ad" & last2))
TextBox79.Text = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("ae9:ae" & last2))
TextBox80.Text = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("av9:av" & last2))
TextBox66.Text = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("aw9:aw" & last2))
TextBox76.Text = Application.WorksheetFunction.Sum(Sheets(ComboBox16.value).Range("X9:X" & last2))
TextBox67.Text = ""

TextBox66.Text = (TextBox66.Text + simadd) - simsubtract
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox79.value = Format(TextBox79.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")
TextBox76.value = Format(TextBox76.value, "#,## IQD")




ElseIf Sheets(ComboBox16.value).Cells(frw1, "h") Like "*" & TextBox67.Text & "*" And Sheets(ComboBox16.value).Cells(frw1, "f") Like "*" & ComboBox2.Text & "*" And ComboBox3.value = Sheets(ComboBox16.value).Cells(frw1, "bf").value Then
  item2.SubItems(1) = Sheets(ComboBox16.value).Cells(frw2, "b")
item2.SubItems(2) = Sheets(ComboBox16.value).Cells(frw2, "c")
item2.SubItems(3) = Sheets(ComboBox16.value).Cells(frw2, "d")
item2.SubItems(4) = Sheets(ComboBox16.value).Cells(frw2, "e")
item2.SubItems(5) = Sheets(ComboBox16.value).Cells(frw2, "f")
item2.SubItems(6) = Sheets(ComboBox16.value).Cells(frw2, "g")
item2.SubItems(7) = Sheets(ComboBox16.value).Cells(frw2, "h")
item2.SubItems(8) = Sheets(ComboBox16.value).Cells(frw2, "i")
item2.SubItems(9) = Sheets(ComboBox16.value).Cells(frw2, "j")
item2.SubItems(10) = Sheets(ComboBox16.value).Cells(frw2, "k")
item2.SubItems(11) = Sheets(ComboBox16.value).Cells(frw2, "l")
item2.SubItems(12) = Sheets(ComboBox16.value).Cells(frw2, "m")
item2.SubItems(13) = Sheets(ComboBox16.value).Cells(frw2, "n")
item2.SubItems(14) = Sheets(ComboBox16.value).Cells(frw2, "o")
item2.SubItems(15) = Sheets(ComboBox16.value).Cells(frw2, "p")
item2.SubItems(16) = Sheets(ComboBox16.value).Cells(frw2, "q")
item2.SubItems(17) = Sheets(ComboBox16.value).Cells(frw2, "r")
item2.SubItems(18) = Sheets(ComboBox16.value).Cells(frw2, "s")
item2.SubItems(19) = Sheets(ComboBox16.value).Cells(frw2, "t")
item2.SubItems(20) = Sheets(ComboBox16.value).Cells(frw2, "u")
item2.SubItems(21) = Sheets(ComboBox16.value).Cells(frw2, "v")
item2.SubItems(22) = Sheets(ComboBox16.value).Cells(frw2, "w")
item2.SubItems(23) = Sheets(ComboBox16.value).Cells(frw2, "x")
item2.SubItems(24) = Sheets(ComboBox16.value).Cells(frw2, "y")
item2.SubItems(25) = Sheets(ComboBox16.value).Cells(frw2, "z")
item2.SubItems(26) = Sheets(ComboBox16.value).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(ComboBox16.value).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(ComboBox16.value).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(ComboBox16.value).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(ComboBox16.value).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(ComboBox16.value).Cells(frw2, "af")
item2.SubItems(32) = Sheets(ComboBox16.value).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(ComboBox16.value).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(ComboBox16.value).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(ComboBox16.value).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(ComboBox16.value).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(ComboBox16.value).Cells(frw2, "al")
item2.SubItems(38) = Sheets(ComboBox16.value).Cells(frw2, "am")
item2.SubItems(39) = Sheets(ComboBox16.value).Cells(frw2, "an")
item2.SubItems(40) = Sheets(ComboBox16.value).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(ComboBox16.value).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(ComboBox16.value).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(ComboBox16.value).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(ComboBox16.value).Cells(frw2, "as")
item2.SubItems(45) = Sheets(ComboBox16.value).Cells(frw2, "at")
item2.SubItems(46) = Sheets(ComboBox16.value).Cells(frw2, "au")
item2.SubItems(47) = Sheets(ComboBox16.value).Cells(frw2, "av")
item2.SubItems(48) = Sheets(ComboBox16.value).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(ComboBox16.value).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(ComboBox16.value).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(ComboBox16.value).Cells(frw2, "az")
item2.SubItems(52) = Sheets(ComboBox16.value).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(ComboBox16.value).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(ComboBox16.value).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(ComboBox16.value).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(ComboBox16.value).Cells(frw2, "be")
item2.SubItems(57) = Sheets(ComboBox16.value).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(ComboBox16.value).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(ComboBox16.value).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(ComboBox16.value).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(ComboBox16.value).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(ComboBox16.value).Cells(frw2, "bk")
item2.SubItems(63) = Sheets(ComboBox16.value).Cells(frw2, "bl")





'==================================================
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("j9:j8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("ad9:ad800"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("ae9:ae8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("av9:av8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("aw9:aw8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(ComboBox16.value).Range("x9:x8000"), Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*", Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(ComboBox16.value).Range("h9:h8000"), "=" & Me.TextBox67.Text, Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*")






TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox79.value = Format(TextBox79.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")
TextBox76.value = Format(TextBox76.value, "#,## IQD")
End If
Next






'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(combobox16.value).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(combobox16.value).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*", Sheets(combobox16.value).Range("BF9:BF571"), "=" & Me.ComboBox3.Value)


'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(combobox16.value).Range("BF9:BF" & frw2), "=" & Me.ComboBox3.Text)
TextBox75.Text = ListView1.ListItems.count




emad:



End Sub



Private Sub Frame_Click()

End Sub

Private Sub Label108_Click()


End Sub



Private Sub Label125_Click()

End Sub

Private Sub Label124_Click()
UserForm41.Show
End Sub

Private Sub ListBox1_AfterUpdate()
On Error Resume Next
Call arb

End Sub

Private Sub ListBox1_Click()
On Error Resume Next
Call arb

Sheets(ComboBox16.value).Activate
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
last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(ComboBox16.value).Cells(frw1, "b") = TextBox2.Text And Sheets(ComboBox16.value).Cells(frw1, "e") = TextBox5.Text Then
Image1.Picture = LoadPicture(Sheets(ComboBox16.value).Cells(frw1, 64))
Image2.Picture = LoadPicture(Sheets(ComboBox16.value).Cells(frw1, 65))
ComboBox4.Text = Sheets(ComboBox16.value).Cells(frw1, 4)
ComboBox5.Text = Sheets(ComboBox16.value).Cells(frw1, 6)
ComboBox6.Text = Sheets(ComboBox16.value).Cells(frw1, "m")
ComboBox7.Text = Sheets(ComboBox16.value).Cells(frw1, "o")
ComboBox8.Text = Sheets(ComboBox16.value).Cells(frw1, "q")
ComboBox9.Text = Sheets(ComboBox16.value).Cells(frw1, "ay")
ComboBox10.Text = Sheets(ComboBox16.value).Cells(frw1, "z")
ComboBox11.Text = Sheets(ComboBox16.value).Cells(frw1, "ab")
ComboBox3.Text = Sheets(ComboBox16.value).Cells(frw1, "bf")

TextBox82.Text = Format(Sheets(ComboBox16.value).Cells(frw1, "bn"), "yyyy/mm/dd")

TextBox83.Text = Sheets(ComboBox16.value).Cells(frw1, "bo")
TextBox84.Text = Sheets(ComboBox16.value).Cells(frw1, "bp")
TextBox85.Text = Sheets(ComboBox16.value).Cells(frw1, "bq")
TextBox86.Text = Sheets(ComboBox16.value).Cells(frw1, "br")
TextBox87.Text = Sheets(ComboBox16.value).Cells(frw1, "bs")
TextBox88.Text = Sheets(ComboBox16.value).Cells(frw1, "bt")
TextBox91.Text = Sheets(ComboBox16.value).Cells(frw1, "bv")
TextBox99.Text = Sheets(ComboBox16.value).Cells(frw1, "bu")
TextBox100.Text = Sheets(ComboBox16.value).Cells(frw1, "ca")
TextBox105.Text = Sheets(ComboBox16.value).Cells(frw1, "ce")

CheckBox4.value = Sheets(ComboBox16.value).Cells(frw1, "cd")

'=====================

'=========================


If TextBox58.Text = "⁄ﬁœ" Then
TextBox10.Enabled = False
Label123.Visible = True
TextBox10.Text = Sheets(ComboBox16.value).Cells(frw1, "j")
TextBox10.Visible = False
TextBox101.Visible = True

TextBox101.Text = Sheets(ComboBox16.value).Cells(frw1, "j")

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


TextBox89.value = Sheets(ComboBox16.value).Cells(2, "bw").value
TextBox90.value = Sheets(ComboBox16.value).Cells(2, "bx").value

If TextBox58.Text = "—« » Ã“∆Ì" Or TextBox58.Text = "«„Ê„Â À«‰Ì 6 «‘Â—" Then
CommandButton16.Enabled = False
Else
CommandButton16.Enabled = True
End If

'===========================================
Sheets(ComboBox16.value).Cells(4, "cb").value = TextBox58.Text

TextBox93.value = Sheets(ComboBox16.value).Cells(4, "cd")
TextBox94.value = Sheets(ComboBox16.value).Cells(4, "ce")
TextBox97.value = Sheets(ComboBox16.value).Cells(4, "cf")
TextBox98.value = Sheets(ComboBox16.value).Cells(4, "cg")
TextBox95.value = Sheets(ComboBox16.value).Cells(4, "ch")
TextBox96.value = Sheets(ComboBox16.value).Cells(4, "cc")
  
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

TextBox75.Text = ListView1.ListItems.count
TextBox93.value = Format(TextBox93.value, "#,## IQD")
TextBox94.value = Format(TextBox94.value, "#,## IQD")
TextBox97.value = Format(TextBox97.value, "#,## IQD")
TextBox98.value = Format(TextBox98.value, "#,## IQD")
TextBox95.value = Format(TextBox95.value, "#,## IQD")
TextBox96.value = Format(TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & TextBox58.Text

TextBox89.value = Format(val(TextBox85) / val(TextBox10) * 100, "00.0") & "%"
TextBox90.value = Format(val(TextBox86) / val(TextBox10) * 100, "00.0") & "%"

ComboBox12.value = ""
ComboBox13.value = ""
TextBox103.value = "„ﬁœ«— ‰”»… «·„‹‹‹Œ’’«  «·„‹‹∆ÊÌ…" & " = " & (val(TextBox14) + val(TextBox16) + val(TextBox18) + val(TextBox27) + val(TextBox29)) / val(TextBox10) * 100 & "%"
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
If Not Intersect(Target, ActiveControl) Is Nothing Then
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
With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If TextBox58.Text = Sheets(ComboBox16.value).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")

End If



Next frw1

TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox64.value = Format(TextBox64.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")

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
With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If TextBox58.Text = Sheets(ComboBox16.value).Cells(frw1, "bf") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")

End If



Next frw1

TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox64.value = Format(TextBox64.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")

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
   TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(ComboBox16.value).Range("bf9:bf" & last1), Me.TextBox58.Text)

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
Sheets(ComboBox16.value).Cells(3, 64) = TextBox65.Text
TextBox72.Text = Sheets(ComboBox16.value).Cells(3, 65).Text
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
Sheets(ComboBox16.value).Activate
last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row



'=============================================

   
    
    
    For frw1 = 9 To last1
                
  If Sheets(ComboBox16.value).Cells(frw1, "e") Like "*" & TextBox67.Text & "*" Then
            
   Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
   
   
   
'==================================================

item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")
'item1.SubItems(63) = Sheets(combobox16.value).Cells(frw1, "bl")
'item1.SubItems(64) = Sheets(combobox16.value).Cells(frw1, "bm")

End If
Next frw1
End With
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(ComboBox16.value).Range("e9:e8000"), "*" & Me.TextBox67.Text & "*", Sheets(ComboBox16.value).Range("f9:f8000"), "*" & Me.ComboBox2.Text & "*")

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
last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bf")




Next frw1

TextBox62.value = Format(TextBox62.value, "#,## IQD")
TextBox63.value = Format(TextBox63.value, "#,## IQD")
TextBox64.value = Format(TextBox64.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")
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


 arb




ListBox1.Clear

Sheets(ComboBox16.value).Activate
    For i = 1 To 63
            Controls("TextBox" & i).Text = ""
    Next i

    If TextBox72 = "" Then Exit Sub
    Sheets(ComboBox16.value).Activate

    ss = Sheets(ComboBox16.value).Cells(Rows.count, 5).End(xlUp).row
     k = 0
     
     
For Each c In Range("E8:E" & ss)
    If c Like TextBox72.value & "*" Then
        ListBox1.AddItem
        ListBox1.List(k, 0) = Cells(c.row, 5).value
        ListBox1.List(k, 1) = c.row
        k = k + 1
    End If
    
Next c

End Sub

Private Sub TextBox72_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Sheets(ComboBox16.value).Activate

ComboBox15.Text = "«·»ÕÀ ›Ì «·„Êﬁ› «·ÊŸÌ›Ì"
TextBox102.value = Format(Sheets(26).Cells(3, "t"), "00.0")
If TextBox102.value = 0 Then
 Label122.Caption = "«·ﬁÌ„ „ ÿ«»ﬁ…"
 Else
 Label122.Caption = "«·ﬁÌ„ €Ì— „ ÿ«»ﬁ…"

End If

TextBox72.Text = ""


For i = 9 To Sheets(ComboBox16.value).Cells(Rows.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 9, 0) = Cells(i, 5).value
  ListBox1.List(i - 9, 1) = i
      Next i
      If TextBox22.value = "" Then
TextBox20.value = ""
End If

Dim last1 As Integer
last1 = Sheets(ComboBox16.value).Cells(Rows.count, "e").End(xlUp).row
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


Sheets(ComboBox16.value).Activate

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)
ComboBox1.List = Sheets(4).Range("f1:f11").value
ComboBox2.List = Sheets(5).Range("a1:a373").value
ComboBox3.List = Sheets(4).Range("o1:p13").value
ComboBox15.List = Sheets(4).Range("o1:p13").value
ComboBox4.List = Array("–ﬂ—", "«‰ÀÏ")
ComboBox14.List = Array("ŒÿÊ—… ”«∆ﬁ", "·« Ì” Õﬁ")

ComboBox5.List = Sheets(5).Range("a1:a373").value
ComboBox6.List = Array("0.15", "0.25", "0.35", "0.45", "0.55", "0.75", "1.25", "1.50", "0")
ComboBox7.List = Array("0.35", "0.5", "0")
ComboBox8.List = Array("0.5", "0.3", "0.25", "0.20", "0.15", "0")
ComboBox9.List = Array("„«” —", "ﬂ«‘")
ComboBox10.List = Array("0.3", "0.25", "0.20", "0.15", "0")
ComboBox11.List = Array("1", "0.80", "0.30", "0.25", "0.20", "0.15", "0")
ComboBox12.List = Array("ﬂ«‰Ê‰ «·À«‰Ì", "‘»«ÿ", "¬–«—", "‰Ì”«‰", "√Ì«—", "Õ“Ì—«‰", " „Ê“", "¬»", "√Ì·Ê·", " ‘—Ì‰†«·√Ê·", " ‘—Ì‰†«·À«‰Ì", "ﬂ«‰Ê‰†«·√Ê·")
ComboBox13.List = Array("ﬂ«‰Ê‰ «·À«‰Ì", "‘»«ÿ", "¬–«—", "‰Ì”«‰", "√Ì«—", "Õ“Ì—«‰", " „Ê“", "¬»", "√Ì·Ê·", " ‘—Ì‰†«·√Ê·", " ‘—Ì‰†«·À«‰Ì", "ﬂ«‰Ê‰†«·√Ê·")








Me.TextBox92.Text = Sheets(ComboBox16.value).Cells(1, "cb")
'========================================================= «·Õ”«»«  «·ﬂ·ÌÂ Õ”»  ›«’Ì·Â«
TextBox76.value = Sheets(ComboBox16.value).Cells(2, "cc")
TextBox73.value = Sheets(ComboBox16.value).Cells(2, "cd")
TextBox74.value = Sheets(ComboBox16.value).Cells(2, "ce")
TextBox79.Text = Sheets(ComboBox16.value).Cells(2, "cf")
TextBox80.Text = Sheets(ComboBox16.value).Cells(2, "cg")
TextBox82.Text = Sheets(ComboBox16.value).Cells(1, "bn")
TextBox66.Text = Sheets(ComboBox16.value).Cells(2, "ch")


'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(combobox16.value).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox79.value = Format(TextBox79.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")
TextBox76.value = Format(TextBox76.value, "#,## IQD")









TextBox102.value = Format(Sheets(26).Cells(3, "t"), "00.0")
If TextBox102.value = 0 Then
 Label122.Caption = "«·ﬁÌ„ „ ÿ«»ﬁ…"
 Else
 Label122.Caption = "«·ﬁÌ„ €Ì— „ ÿ«»ﬁ…"

End If




Dim lastr1 As Integer
lastr1 = Sheets(ComboBox16.value).Cells(Rows.count, "b").End(xlUp).row

Ê—ﬁ…1.Range("cb9").Formula = "=TEXT(TODAY(),""mmm"")"
Ê—ﬁ…1.Range("cb9:cb" & lastr1).FillDown
Dim last1 As Integer
last1 = Sheets(ComboBox16.value).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last1 - 8




emad:
End Sub





Private Sub UserForm_Click()
Dim last1 As Integer
last1 = Sheets(ComboBox16.value).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last1 - 8

End Sub

Private Sub UserForm_Initialize()
'===============================
On Error Resume Next

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


ComboBox15.List = Sheets(4).Range("o1:p13").value













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
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 140
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
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(ComboBox16.value).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(ComboBox16.value).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(ComboBox16.value).Cells(frw1, "b")
item1.SubItems(2) = Sheets(ComboBox16.value).Cells(frw1, "c")
item1.SubItems(3) = Sheets(ComboBox16.value).Cells(frw1, "d")
item1.SubItems(4) = Sheets(ComboBox16.value).Cells(frw1, "e")
item1.SubItems(5) = Sheets(ComboBox16.value).Cells(frw1, "f")
item1.SubItems(6) = Sheets(ComboBox16.value).Cells(frw1, "g")
item1.SubItems(7) = Sheets(ComboBox16.value).Cells(frw1, "h")
item1.SubItems(8) = Sheets(ComboBox16.value).Cells(frw1, "i")
item1.SubItems(9) = Sheets(ComboBox16.value).Cells(frw1, "j")
item1.SubItems(10) = Sheets(ComboBox16.value).Cells(frw1, "k")
item1.SubItems(11) = Sheets(ComboBox16.value).Cells(frw1, "l")
item1.SubItems(12) = Sheets(ComboBox16.value).Cells(frw1, "m")
item1.SubItems(13) = Sheets(ComboBox16.value).Cells(frw1, "n")
item1.SubItems(14) = Sheets(ComboBox16.value).Cells(frw1, "o")
item1.SubItems(15) = Sheets(ComboBox16.value).Cells(frw1, "p")
item1.SubItems(16) = Sheets(ComboBox16.value).Cells(frw1, "q")
item1.SubItems(17) = Sheets(ComboBox16.value).Cells(frw1, "r")
item1.SubItems(18) = Sheets(ComboBox16.value).Cells(frw1, "s")
item1.SubItems(19) = Sheets(ComboBox16.value).Cells(frw1, "t")
item1.SubItems(20) = Sheets(ComboBox16.value).Cells(frw1, "u")
item1.SubItems(21) = Sheets(ComboBox16.value).Cells(frw1, "v")
item1.SubItems(22) = Sheets(ComboBox16.value).Cells(frw1, "w")
item1.SubItems(23) = Sheets(ComboBox16.value).Cells(frw1, "x")
item1.SubItems(24) = Sheets(ComboBox16.value).Cells(frw1, "y")
item1.SubItems(25) = Sheets(ComboBox16.value).Cells(frw1, "z")
item1.SubItems(26) = Sheets(ComboBox16.value).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(ComboBox16.value).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(ComboBox16.value).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(ComboBox16.value).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(ComboBox16.value).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(ComboBox16.value).Cells(frw1, "af")
item1.SubItems(32) = Sheets(ComboBox16.value).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(ComboBox16.value).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(ComboBox16.value).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(ComboBox16.value).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(ComboBox16.value).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(ComboBox16.value).Cells(frw1, "al")
item1.SubItems(38) = Sheets(ComboBox16.value).Cells(frw1, "am")
item1.SubItems(39) = Sheets(ComboBox16.value).Cells(frw1, "an")
item1.SubItems(40) = Sheets(ComboBox16.value).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(ComboBox16.value).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(ComboBox16.value).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(ComboBox16.value).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(ComboBox16.value).Cells(frw1, "as")
item1.SubItems(45) = Sheets(ComboBox16.value).Cells(frw1, "at")
item1.SubItems(46) = Sheets(ComboBox16.value).Cells(frw1, "au")
item1.SubItems(47) = Sheets(ComboBox16.value).Cells(frw1, "av")
item1.SubItems(48) = Sheets(ComboBox16.value).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(ComboBox16.value).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(ComboBox16.value).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(ComboBox16.value).Cells(frw1, "az")
item1.SubItems(52) = Sheets(ComboBox16.value).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(ComboBox16.value).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(ComboBox16.value).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(ComboBox16.value).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(ComboBox16.value).Cells(frw1, "be")
item1.SubItems(57) = Sheets(ComboBox16.value).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(ComboBox16.value).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(ComboBox16.value).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(ComboBox16.value).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(ComboBox16.value).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(ComboBox16.value).Cells(frw1, "bk")

item1.SubItems(63) = Sheets(ComboBox16.value).Cells(frw1, "bl")
item1.SubItems(64) = Sheets(ComboBox16.value).Cells(frw1, "bm")
item1.SubItems(65) = Sheets(ComboBox16.value).Cells(frw1, "bn")
item1.SubItems(66) = Sheets(ComboBox16.value).Cells(frw1, "bo")
item1.SubItems(67) = Sheets(ComboBox16.value).Cells(frw1, "bp")
item1.SubItems(68) = Sheets(ComboBox16.value).Cells(frw1, "bq")
item1.SubItems(69) = Sheets(ComboBox16.value).Cells(frw1, "br")
item1.SubItems(70) = Sheets(ComboBox16.value).Cells(frw1, "bs")
item1.SubItems(71) = Sheets(ComboBox16.value).Cells(frw1, "bt")





Next frw1

TextBox73.value = Format(TextBox73.value, "#,## IQD")
TextBox74.value = Format(TextBox74.value, "#,## IQD")
TextBox64.value = Format(TextBox64.value, "#,## IQD")
TextBox80.value = Format(TextBox80.value, "#,## IQD")
TextBox66.value = Format(TextBox66.value, "#,## IQD")

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
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
 If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If
Label104.Caption = Sheets("title_factory").Cells(2, 1)

Label105.Caption = Sheets("title_factory").Cells(2, 3)
Label123.Visible = False


Dim last As Integer
last = Sheets(ComboBox16.value).Cells(Rows.count, "e").End(xlUp).row
TextBox75.Text = last - 8











End Sub

Private Sub GetMonthDataToSheet1(monthSheetName As String)

    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, LastCol As Long
    Dim dataRange As Range, dataArr As Variant, headersArr As Variant
    Dim nRows As Long, nCols As Long
    Dim actualCells As Long

    Set wsSrc = Sheets(monthSheetName)
    Set wsDest = Sheets(ComboBox16.value)

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
