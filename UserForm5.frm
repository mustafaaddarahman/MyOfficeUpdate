VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   " «—ÌŒ «Œ—  ”ﬂÌ‰ Ê⁄·«Ê…"
   ClientHeight    =   11760
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   18684
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
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

Private Sub CheckBox1_Click()
If CheckBox1.value = True Or CheckBox2.value = True Then
CommandButton6.Enabled = True

Else
CommandButton6.Enabled = False
End If
End Sub

Private Sub CheckBox2_Change()
If CheckBox1.value = True Or CheckBox2.value = True Then
CommandButton6.Enabled = True
Else
CommandButton6.Enabled = False
End If
End Sub

Private Sub ComboBox1_AfterUpdate()
TextBox21.Text = ComboBox1.value
CheckBox1.value = False
CheckBox2.value = False
For i = 1 To 20
        Me.Controls("TextBox" & i).Text = ""
    Next i
   
End Sub

Private Sub ComboBox1_Change()
TextBox21.Text = ComboBox1.value
For i = 1 To 20
        Me.Controls("TextBox" & i).Text = ""
    Next i
    
          TextBox23.Text = ""
          Me.TextBox13.Text = "00/00/0000"
       Me.TextBox14.Text = "00/00/0000"
End Sub

Private Sub CommandButton1_Click()
'=COUNTIFS($L$2:$L$30;">=" &$S4;$B$2:$B$30;R4;E$2:E$30;$Q4)
Me.ListView1.ListItems.Clear

Sheets(6).Activate

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True
.ColumnHeaders.Add , , "«·—ﬁ„", 40
.ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
.ColumnHeaders.Add , , "—ﬁ„ «·„«” —", 140
.ColumnHeaders.Add , , "«·Ã‰”", 60
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 140
.ColumnHeaders.Add , , "«·⁄‰Ê«‰ «·ÊŸÌ›Ì", 100
.ColumnHeaders.Add , , " «·ﬁ”„", 100
.ColumnHeaders.Add , , " «·œ—Ã…", 60
.ColumnHeaders.Add , , "«·„—Õ·…", 50
.ColumnHeaders.Add , , " «·—« » «·«”„Ì", 100
.ColumnHeaders.Add , , " «—ÌŒ «Œ— ⁄·«Ê…", 140
.ColumnHeaders.Add , , "  «—ÌŒ ﬂ «» «·‘ﬂ— Ê«· ﬁœÌ—", 140
.ColumnHeaders.Add , , "  «—ÌŒ «” Õﬁ«ﬁ «·⁄·«Ê…", 140
.ColumnHeaders.Add , , " «—ÌŒ «Œ—  —›Ì⁄", 140
.ColumnHeaders.Add , , "  «—ÌŒ «” Õﬁ«ﬁ «· —›Ì⁄", 140
.ColumnHeaders.Add , , "⁄œœ «·”‰Ê«  «· Ì ﬁ÷«Â« „‰ «Œ—  —›Ì⁄ ", 190


ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"
End With
'======================================== „—Õ·… «„·«¡ «·»Ì«‰« 
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
last1 = Sheets(6).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If TextBox17.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox11, Sheets(6).Range("$b$2:$b$10000"), _
  Me.TextBox2, Sheets(6).Range("$e$2:$e$10000"), Me.TextBox5.Text) And Sheets(6).Cells(frw1, "e") Like "*" & TextBox5.Text & "*" _
 And TextBox11.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("k2:k10000"), Sheets(6).Range("e2:e10000"), Me.TextBox5.Text), "yyyy/mm/dd") _
 And TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("e2:e10000"), Me.TextBox5.Text), "yyyy/mm/dd") Then
Set item1 = ListView1.ListItems.Add(, , Sheets(6).Cells(frw1, "A"))
 item1.SubItems(1) = Sheets(6).Cells(frw1, "b")
item1.SubItems(2) = Sheets(6).Cells(frw1, "c")
item1.SubItems(3) = Sheets(6).Cells(frw1, "d")
item1.SubItems(4) = Sheets(6).Cells(frw1, "e")
item1.SubItems(5) = Sheets(6).Cells(frw1, "f")
item1.SubItems(6) = Sheets(6).Cells(frw1, "g")
item1.SubItems(7) = Sheets(6).Cells(frw1, "h")
item1.SubItems(8) = Sheets(6).Cells(frw1, "i")
item1.SubItems(9) = Sheets(6).Cells(frw1, "j")
item1.SubItems(10) = Sheets(6).Cells(frw1, "k")
item1.SubItems(11) = Sheets(6).Cells(frw1, "l")
item1.SubItems(12) = Sheets(6).Cells(frw1, "m")
item1.SubItems(13) = Sheets(6).Cells(frw1, "n")
item1.SubItems(14) = Sheets(6).Cells(frw1, "o")
item1.SubItems(15) = Sheets(6).Cells(frw1, "p")

End If
Next frw1
'=============================================
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
    

End With
End With
End With
ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"

End Sub

Private Sub CommandButton2_Click()


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
TextBox11.Text = ""
TextBox23.Text = ""
TextBox12.Text = ""
TextBox13.Text = ""
TextBox14.Text = ""




On Error Resume Next
'Dim row1, last1 As Integer
'last1 = Sheets(1).Range("l" & Rows.Count).End(xlUp).row
'For row1 = 2 To last1

For Each c1 In Worksheets(6).Range("b:b")
If c1 = TextBox21.Text Then
TextBox2.value = c1

TextBox1.value = Application.WorksheetFunction.MaxIfs(Sheets(6).Range("a2:a10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text)
'=MAXIFS(A2:A8,B2:B8,Q4)
TextBox3.value = c1.Offset(0, 1)

TextBox4.value = c1.Offset(0, 2)

TextBox5.value = c1.Offset(0, 3)
TextBox6.value = c1.Offset(0, 4)
TextBox7.value = c1.Offset(0, 5)
TextBox8.value = c1.Offset(0, 6)
TextBox9.value = c1.Offset(0, 7)
TextBox10.value = c1.Offset(0, 8)
TextBox15.value = c1.Offset(0, 15)

 
 
 
 
TextBox11.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("k2:k10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
'=MAXIFS(L2:L9,B2:B9,Q4)

TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")


TextBox12.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("m2:m10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox13.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("n2:n10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox14.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("o2:o10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox15.value = Application.WorksheetFunction.MaxIfs(Sheets(6).Range("p2:p10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text)
TextBox17.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox11.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)


 If TextBox17.value > 3 Then
Me.TextBox12 = Format(Application.WorksheetFunction.EDate(Me.TextBox11.Text, 12 - 3), "yyyy/mm/dd")
Else
 Me.TextBox12 = Format(Application.WorksheetFunction.EDate(Me.TextBox11.Text, 12 - Me.TextBox17.value), "yyyy/mm/dd")
 End If
 If Year(TextBox13.Text) < 2000 Then
 TextBox14.Text = ""
 TextBox13.Text = ""
 Else
 TextBox14.Text = Format(Application.WorksheetFunction.EDate(Me.TextBox13.Text, (12 * 5) - Me.TextBox17.value), "yyyy/mm/dd")
 '«ÌÃ«œ «·› —… «· Ì ﬁ÷«Â« ›Ì «Œ—  —›Ì⁄
'===============================================
Dim xday, xmonth, xyear As Integer
Dim X, Y As Date
Y = TextBox13.Text
X = Date
xyear = DateDiff("yyyy", Y, X)
xmonth = DateDiff("m", Y, X) Mod 12
Dim a, b As Integer
b = Day(X) - Day(Y)
xday = b
If b < 0 Then
xmonth = xmonth - 1
xday = xday + 30
End If
a = Month(X) - Month(Y)
If a < 0 Then
xyear = xyear - 1
End If
Me.TextBox15.value = xyear & "”‰…/" & xmonth & "‘Â—/" & xday & "ÌÊ„"
 '========================================
  
 

 
 End If
 Me.ListView1.ListItems.Clear
 If Format(TextBox23.Text, "yyyy/mm/dd") = "30/12/1899" Then
TextBox23.Text = ""
Else
TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
End If
Exit Sub
End If
Next

'============================
'=COUNTIFS(L2:L15;">=" &S4;B2:B15;R4;E2:E15;Q4)

Dim fondnam As Long
If fondnam = Application.WorksheetFunction.CountIf(Sheets(6).Range("b2:b100000"), Me.TextBox21.value) Then
fondnam = 0
MsgBox "Â–« «·ﬁÌœ €Ì— „ÊÃÊœ Õ«Ê· «· «ﬂœ „‰Â «Ê ﬁœ ÌﬂÊ‰ Œÿ√ ›Ì ﬂ «» Â", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "—”«·Â  Õ–Ì—ÌÂ ⁄„«œ «·‰⁄Ì„Ì"
Exit Sub
End If

 '========================================
 '=======================================
 
 '=VLOOKUP(S27,$B$1:$J$39,4,0)
 '=VLOOKUP(S27,data,4,0)

TextBox13.value = ""
TextBox14.value = ""
End Sub

Private Sub CommandButton3_Click()
On Error Resume Next
'Dim row1, last1 As Integer
'last1 = Sheets(1).Range("l" & Rows.Count).End(xlUp).row
'For row1 = 2 To last1

For Each c1 In Worksheets(6).Range("b:b")
If c1 = TextBox21.Text Then
TextBox2.value = c1

TextBox1.value = Application.WorksheetFunction.MaxIfs(Sheets(6).Range("a2:a10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text)
'=MAXIFS(A2:A8,B2:B8,Q4)
TextBox3.value = c1.Offset(0, 1)

TextBox4.value = c1.Offset(0, 2)

TextBox5.value = c1.Offset(0, 3)
TextBox6.value = c1.Offset(0, 4)
TextBox7.value = c1.Offset(0, 5)
TextBox8.value = c1.Offset(0, 6)
TextBox9.value = c1.Offset(0, 7)
TextBox10.value = c1.Offset(0, 8)
TextBox15.value = c1.Offset(0, 15)

 
 
 
 
TextBox11.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("k2:k10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
'=MAXIFS(L2:L9,B2:B9,Q4)

TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")


TextBox12.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("m2:m10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox13.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("n2:n10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox14.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("o2:o10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox15.value = Application.WorksheetFunction.MaxIfs(Sheets(6).Range("p2:p10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text)
TextBox17.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox11.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)


 If TextBox17.value > 3 Then
Me.TextBox12 = Format(Application.WorksheetFunction.EDate(Me.TextBox11.Text, 12 - 3), "yyyy/mm/dd")
Else
 Me.TextBox12 = Format(Application.WorksheetFunction.EDate(Me.TextBox11.Text, 12 - Me.TextBox17.value), "yyyy/mm/dd")
 End If
 If Year(TextBox13.Text) < 2000 Then
 TextBox14.Text = ""
 TextBox13.Text = ""
 Else
 TextBox14.Text = Format(Application.WorksheetFunction.EDate(Me.TextBox13.Text, (12 * 5) - Me.TextBox17.value), "yyyy/mm/dd")
 '«ÌÃ«œ «·› —… «· Ì ﬁ÷«Â« ›Ì «Œ—  —›Ì⁄
'===============================================
Dim xday, xmonth, xyear As Integer
Dim X, Y As Date
Y = TextBox13.Text
X = Date
xyear = DateDiff("yyyy", Y, X)
xmonth = DateDiff("m", Y, X) Mod 12
Dim a, b As Integer
b = Day(X) - Day(Y)
xday = b
If b < 0 Then
xmonth = xmonth - 1
xday = xday + 30
End If
a = Month(X) - Month(Y)
If a < 0 Then
xyear = xyear - 1
End If
Me.TextBox15.value = xyear & "”‰…/" & xmonth & "‘Â—/" & xday & "ÌÊ„"
 '========================================
  
 

 
 End If
 Me.ListView1.ListItems.Clear
 If Format(TextBox23.Text, "yyyy/mm/dd") = "" Then
TextBox23.Text = ""
Else
TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
End If
Exit Sub
End If
Next

'============================
'=COUNTIFS(L2:L15;">=" &S4;B2:B15;R4;E2:E15;Q4)

Dim fondnam As Long
If fondnam = Application.WorksheetFunction.CountIf(Sheets(6).Range("b2:b100000"), Me.TextBox21.value) Then
fondnam = 0
MsgBox "Â–« «·ﬁÌœ €Ì— „ÊÃÊœ Õ«Ê· «· «ﬂœ „‰Â «Ê ﬁœ ÌﬂÊ‰ Œÿ√ ›Ì ﬂ «» Â", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "—”«·Â  Õ–Ì—ÌÂ ⁄„«œ «·‰⁄Ì„Ì"
Exit Sub
End If

 '========================================
 '=======================================
 
 '=VLOOKUP(S27,$B$1:$J$39,4,0)
 '=VLOOKUP(S27,data,4,0)

TextBox13.value = ""
TextBox14.value = ""
End Sub

Private Sub CommandButton4_Click()
Me.ListView1.ListItems.Clear

Sheets(6).Activate

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True
.ColumnHeaders.Add , , "«·—ﬁ„", 40
.ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
.ColumnHeaders.Add , , "—ﬁ„ «·„«” —", 140
.ColumnHeaders.Add , , "«·Ã‰”", 60
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 140
.ColumnHeaders.Add , , "«·⁄‰Ê«‰ «·ÊŸÌ›Ì", 100
.ColumnHeaders.Add , , " «·ﬁ”„", 100
.ColumnHeaders.Add , , " «·œ—Ã…", 60
.ColumnHeaders.Add , , "«·„—Õ·…", 50
.ColumnHeaders.Add , , " «·—« » «·«”„Ì", 100
.ColumnHeaders.Add , , " «—ÌŒ «Œ— ⁄·«Ê…", 140
.ColumnHeaders.Add , , "  «—ÌŒ ﬂ «» «·‘ﬂ— Ê«· ﬁœÌ—", 140
.ColumnHeaders.Add , , "  «—ÌŒ «” Õﬁ«ﬁ «·⁄·«Ê…", 140
.ColumnHeaders.Add , , " «—ÌŒ «Œ—  —›Ì⁄", 140
.ColumnHeaders.Add , , "  «—ÌŒ «” Õﬁ«ﬁ «· —›Ì⁄", 140
.ColumnHeaders.Add , , "⁄œœ «·”‰Ê«  «· Ì ﬁ÷«Â« „‰ «Œ—  —›Ì⁄ ", 190


ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"
End With
'======================================== „—Õ·… «„·«¡ «·»Ì«‰« 
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
last1 = Sheets(6).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(6).Cells(frw1, "b").value = Me.TextBox2.Text Or Sheets(6).Cells(frw1, "e") Like "*" & TextBox5.Text & "*" Then
Set item1 = ListView1.ListItems.Add(, , Sheets(6).Cells(frw1, "A"))
 item1.SubItems(1) = Sheets(6).Cells(frw1, "b")
item1.SubItems(2) = Sheets(6).Cells(frw1, "c")
item1.SubItems(3) = Sheets(6).Cells(frw1, "d")
item1.SubItems(4) = Sheets(6).Cells(frw1, "e")
item1.SubItems(5) = Sheets(6).Cells(frw1, "f")
item1.SubItems(6) = Sheets(6).Cells(frw1, "g")
item1.SubItems(7) = Sheets(6).Cells(frw1, "h")
item1.SubItems(8) = Sheets(6).Cells(frw1, "i")
item1.SubItems(9) = Sheets(6).Cells(frw1, "j")
item1.SubItems(10) = Sheets(6).Cells(frw1, "k")
item1.SubItems(11) = Sheets(6).Cells(frw1, "l")
item1.SubItems(12) = Sheets(6).Cells(frw1, "m")
item1.SubItems(13) = Sheets(6).Cells(frw1, "n")
item1.SubItems(14) = Sheets(6).Cells(frw1, "o")
item1.SubItems(15) = Sheets(6).Cells(frw1, "p")

End If
Next frw1
'=============================================
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
    

End With
End With
End With
ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"

End Sub

Private Sub CommandButton6_Click()
'=============================
If CheckBox1.value = True Then
TextBox9.value = TextBox9.value
End If

If CheckBox2.value = True Then
TextBox8.value = TextBox8.value
End If



'====================================
Sheets(6).Activate
'=========================================================
 For Y = 2 To 1500
    
    If Sheets(6).Cells(Y, 1) = TextBox1.Text Or Sheets(6).Cells(Y, 2) = TextBox2.Text _
    Or TextBox13.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("k2:k10000"), Sheets(6).Range("b2:b10000"), Me.TextBox2.Text), "yyyy/mm/dd") Then
    Exit For
    End If
    Next Y
   ' Sheets(6).Cells(Y, 1) = TextBox1.Text
   ' Sheets(6).Cells(Y, 2) = TextBox2.Text
   ' Sheets(6).Cells(Y, 3) = TextBox3.Text
   ' Sheets(6).Cells(Y, 4) = TextBox4.Text
   ' Sheets(6).Cells(Y, 5) = TextBox5.Text
    Sheets(6).Cells(Y, 6) = TextBox6.Text
    Sheets(6).Cells(Y, 7) = TextBox7.Text
    Sheets(6).Cells(Y, 8) = TextBox8.Text
    Sheets(6).Cells(Y, 9) = TextBox9.Text
    Sheets(6).Cells(Y, 10) = TextBox10.Text
  '  Sheets(6).Cells(Y, 11) = TextBox11.Text
   ' Sheets(6).Cells(Y, 12) = TextBox23.Text
     Sheets(6).Cells(Y, 13) = TextBox12.Text
    Sheets(6).Cells(Y, 14) = TextBox13.Text
    Sheets(6).Cells(Y, 15) = TextBox14.Text
    Sheets(6).Cells(Y, 16) = TextBox15.Text
    Sheets(6).Cells(Y, 17) = TextBox16.Text
    Sheets(6).Cells(Y, 18) = TextBox18.Text
   Sheets(6).Cells(Y, 19) = ComboBox2.value
  
          
       '   For i = 1 To 23
      '  Me.Controls("TextBox" & i).Text = ""
  '  Next i
'????? ?????? ???? ????
   
'????? ??????
   
'????? ???? ???????

 Sheets(8).Activate
'=========================================================










'==================================================== ‘€«·
  For y1 = 2 To 1500
 If Sheets(8).Cells(y1, 1) = "" Or Sheets(8).Cells(y1, 2) = "" Then
 Sheets(8).Cells(y1, 1) = TextBox1.Text
    Sheets(8).Cells(y1, 2) = TextBox2.Text
    Sheets(8).Cells(y1, 3) = TextBox3.Text
    Sheets(8).Cells(y1, 4) = TextBox4.Text
    Sheets(8).Cells(y1, 5) = TextBox5.Text
    Sheets(8).Cells(y1, 6) = TextBox6.Text
    Sheets(8).Cells(y1, 7) = TextBox7.Text
    Sheets(8).Cells(y1, 8) = TextBox8.Text
    Sheets(8).Cells(y1, 9) = TextBox9.Text
    Sheets(8).Cells(y1, 10) = TextBox10.Text
   Sheets(8).Cells(y1, 11) = TextBox11.Text
 Sheets(8).Cells(y1, 12) = TextBox12.Text
     Sheets(8).Cells(y1, 13) = TextBox13.Text
    Sheets(8).Cells(y1, 14) = TextBox14.Text
    Sheets(8).Cells(y1, 15) = TextBox15.Text
    Sheets(8).Cells(y1, 16) = TextBox16.Text
    Sheets(8).Cells(y1, 17) = TextBox18.Text
   Sheets(8).Cells(y1, 18) = ComboBox2.value
 End If
 
 
 
 '=====================================
 

    
       If Sheets(8).Cells(y1, 1) = TextBox1.Text Or Sheets(8).Cells(y1, 2) = TextBox2.Text Then
    Exit For
    End If
    Next y1
   
    Sheets(8).Cells(y1, 1) = TextBox1.Text
    Sheets(8).Cells(y1, 2) = TextBox2.Text
    Sheets(8).Cells(y1, 3) = TextBox3.Text
    Sheets(8).Cells(y1, 4) = TextBox4.Text
    Sheets(8).Cells(y1, 5) = TextBox5.Text
    Sheets(8).Cells(y1, 6) = TextBox6.Text
    Sheets(8).Cells(y1, 7) = TextBox7.Text
    Sheets(8).Cells(y1, 8) = TextBox8.Text
    Sheets(8).Cells(y1, 9) = TextBox9.Text
    Sheets(8).Cells(y1, 10) = TextBox10.Text
   Sheets(8).Cells(y1, 11) = TextBox11.Text
 Sheets(8).Cells(y1, 12) = TextBox12.Text
     Sheets(8).Cells(y1, 13) = TextBox13.Text
    Sheets(8).Cells(y1, 14) = TextBox14.Text
    Sheets(8).Cells(y1, 15) = TextBox15.Text
    Sheets(8).Cells(y1, 16) = TextBox16.Text
    Sheets(8).Cells(y1, 17) = TextBox18.Text
   Sheets(8).Cells(y1, 18) = ComboBox2.value
     
  
  '===============================================
    '  For y2 = 2 To 1500
    
    ' If Sheets(8).Cells(y2, 1) = TextBox1.Text Or Sheets(8).Cells(y2, 2) = TextBox2.Text Then
   
   ' Exit For
  '  End If
  '  Next y2
   
  '  Sheets(8).Cells(y2, 1) = TextBox1.Text
  '  Sheets(8).Cells(y2, 2) = TextBox2.Text
  '  Sheets(8).Cells(y2, 3) = TextBox3.Text
  '  Sheets(8).Cells(y2, 4) = TextBox4.Text
 '   Sheets(8).Cells(y2, 5) = TextBox5.Text
  '  Sheets(8).Cells(y2, 6) = TextBox6.Text
  '  Sheets(8).Cells(y2, 7) = TextBox7.Text
  '  Sheets(8).Cells(y2, 8) = TextBox8.Text
  '  Sheets(8).Cells(y2, 9) = TextBox9.Text
  '  Sheets(8).Cells(y2, 10) = TextBox10.Text
 '  Sheets(8).Cells(y2, 11) = TextBox11.Text
' Sheets(8).Cells(y2, 12) = TextBox12.Text
 '    Sheets(8).Cells(y2, 13) = TextBox13.Text
  '  Sheets(8).Cells(y2, 14) = TextBox14.Text
 '   Sheets(8).Cells(y2, 15) = TextBox15.Text
 '   Sheets(8).Cells(y2, 16) = TextBox16.Text
  '  Sheets(8).Cells(y2, 17) = TextBox18.Text
  ' Sheets(8).Cells(y1, 18) = TextBox18.Text
     
 
     
    
    
    
    
   
    
   
End Sub

Private Sub CommandButton7_Click()
On Error Resume Next



For Each c1 In Worksheets(6).Range("b:b")
If c1 = TextBox25.Text Then
TextBox2.value = c1

TextBox1.value = Application.WorksheetFunction.MaxIfs(Sheets(6).Range("a2:a10000"), Sheets(6).Range("b2:b10000"), Me.TextBox25.Text)
'=MAXIFS(A2:A8,B2:B8,Q4)
TextBox3.value = c1.Offset(0, 1)

TextBox4.value = c1.Offset(0, 2)

TextBox5.value = c1.Offset(0, 3)
TextBox6.value = c1.Offset(0, 4)
TextBox7.value = c1.Offset(0, 5)
TextBox8.value = c1.Offset(0, 6)
TextBox9.value = c1.Offset(0, 7)
TextBox10.value = c1.Offset(0, 8)
TextBox15.value = c1.Offset(0, 15)

 
 
 
 
TextBox11.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("k2:k10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
'=MAXIFS(L2:L9,B2:B9,Q4)

TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")


TextBox12.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("m2:m10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox13.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("n2:n10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox14.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("o2:o10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
TextBox15.value = Application.WorksheetFunction.MaxIfs(Sheets(6).Range("p2:p10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text)
TextBox17.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox11.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)


 If TextBox17.value > 3 Then
Me.TextBox12 = Format(Application.WorksheetFunction.EDate(Me.TextBox11.Text, 12 - 3), "yyyy/mm/dd")
Else
 Me.TextBox12 = Format(Application.WorksheetFunction.EDate(Me.TextBox11.Text, 12 - Me.TextBox17.value), "yyyy/mm/dd")
 End If
 If Year(TextBox13.Text) < 2000 Then
 TextBox14.Text = ""
 TextBox13.Text = ""
 Else
 TextBox14.Text = Format(Application.WorksheetFunction.EDate(Me.TextBox13.Text, (12 * 5) - Me.TextBox17.value), "yyyy/mm/dd")
 '«ÌÃ«œ «·› —… «· Ì ﬁ÷«Â« ›Ì «Œ—  —›Ì⁄
'===============================================
Dim xday, xmonth, xyear As Integer
Dim X, Y As Date
Y = TextBox13.Text
X = Date
xyear = DateDiff("yyyy", Y, X)
xmonth = DateDiff("m", Y, X) Mod 12
Dim a, b As Integer
b = Day(X) - Day(Y)
xday = b
If b < 0 Then
xmonth = xmonth - 1
xday = xday + 30
End If
a = Month(X) - Month(Y)
If a < 0 Then
xyear = xyear - 1
End If
Me.TextBox15.value = xyear & "”‰…/" & xmonth & "‘Â—/" & xday & "ÌÊ„"
 '========================================
  
 

 
 End If
 Me.ListView1.ListItems.Clear
 If Format(TextBox23.Text, "yyyy/mm/dd") = "1899/12/30" Then
TextBox23.Text = ""
Else
TextBox23.value = Format(Application.WorksheetFunction.MaxIfs(Sheets(6).Range("l2:l10000"), Sheets(6).Range("b2:b10000"), Me.TextBox21.Text), "yyyy/mm/dd")
End If
Exit Sub
End If
Next

'============================
'=COUNTIFS(L2:L15;">=" &S4;B2:B15;R4;E2:E15;Q4)

Dim fondnam As LongLong
If fondnam = Application.WorksheetFunction.CountIf(Sheets(6).Range("b2:b100000"), Me.TextBox21.value) Then
fondnam = 0
MsgBox "Â–« «·ﬁÌœ €Ì— „ÊÃÊœ Õ«Ê· «· «ﬂœ „‰Â «Ê ﬁœ ÌﬂÊ‰ Œÿ√ ›Ì ﬂ «» Â", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "—”«·Â  Õ–Ì—ÌÂ ⁄„«œ «·‰⁄Ì„Ì"
Exit Sub
End If

 '========================================
 '=======================================
 
 '=VLOOKUP(S27,$B$1:$J$39,4,0)
 '=VLOOKUP(S27,data,4,0)

TextBox13.value = ""
TextBox14.value = ""
End Sub

Private Sub Label77_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub TextBox11_AfterUpdate()
TextBox11.Text = Format(TextBox11.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox12_AfterUpdate()
TextBox12.Text = Format(TextBox12.Text, "yyyy/mm/dd")
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



Private Sub TextBox13_AfterUpdate()
TextBox13.Text = Format(TextBox13.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox14_AfterUpdate()
TextBox14.Text = Format(TextBox14.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox22_Change()
Dim item1 As ListItem
Dim last1, frw1 As Integer

Sheets(6).Activate
last1 = Sheets(6).Range("A" & Rows.count).End(xlUp).row
     For frw1 = 2 To last1
 If Sheets(6).Cells(frw1, "e") Like "*" & TextBox22.Text & "*" Then
 TextBox21.Text = Sheets(6).Cells(frw1, "b")
End If
Next frw1
End Sub

Private Sub TextBox23_AfterUpdate()
TextBox23.Text = Format(TextBox23.Text, "yyyy/mm/dd")
End Sub

Private Sub UserForm_Activate()

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

 
ComboBox1.List = Sheets(7).Range("a1:a571").value

Dim last1, frw1 As Integer
Dim Sh As Worksheet
Set Sh = ThisWorkbook.Sheets(10)
ComboBox2.Clear

last1 = Sheets(10).Range("A" & Rows.count).End(xlUp).row
For i = 2 To last1

ComboBox2.AddItem Sheets(10).Cells(i, 1).value



Next i



End Sub



