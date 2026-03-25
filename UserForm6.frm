VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "»Ì«‰«  «·„—Õ·Ì‰ Ê«·„÷«›Ì‰"
   ClientHeight    =   12120
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   18492
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm6"
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
On Error Resume Next
Dim result As VbMsgBoxResult
If CheckBox1.value = True Then
'CommandButton13.Enabled = True
CommandButton20.Enabled = False
CheckBox2.value = False

TextBox44.Text = TextBox44.Text + 1
Else
CommandButton13.Enabled = False
TextBox44.Text = ""
End If
result = MsgBox(" Â·  —Ìœ  —ÕÌ· «·⁄·«Ê…«·Œ«’… »«·„ÊŸ› " & TextBox5.Text, vbYesNo + vbQuestion, "—”«·…  ‰»ÌÂ")
If result = vbYes Then
MsgBox "”ÌﬁÊ„ «·‰Ÿ«„ » —ÕÌ· «·⁄·«ÊÂ Ê«÷«›… „—Õ·… Ê“Ì«œ… «·—« » Õ”» «·«” Õﬁ«ﬁ", vbOKOnly, "—”«·…  «ﬂÌœ"
CommandButton13.Enabled = True
ElseIf result = vbNo Then
CheckBox1.value = False
CheckBox2.value = False
CommandButton13.Enabled = False
CommandButton20.Enabled = False
Exit Sub
End If

End Sub

Private Sub CheckBox2_Click()
Dim result As VbMsgBoxResult
If TextBox42.Text = "" Then
TextBox42.Text = 0
Else
Exit Sub
End If
If CheckBox2.value = True Then
CommandButton20.Enabled = True
CommandButton13.Enabled = False
 CheckBox1.value = False
 
 TextBox42.Text = TextBox42.Text + 1
ElseIf CheckBox2.value = False Then
CommandButton20.Enabled = False
TextBox42.Text = ""
End If
result = MsgBox(" Â·  —Ìœ  —ÕÌ· «· —›Ì⁄ «·Œ«’… »«·„ÊŸ› " & TextBox5.Text, vbYesNo + vbQuestion, "—”«·…  ‰»ÌÂ")
If result = vbYes Then
MsgBox "”ÌﬁÊ„ «·‰Ÿ«„ » —ÕÌ· «· —›Ì⁄ Ê«÷«›… „—Õ·… Ê“Ì«œ… «·—« » Õ”» «·«” Õﬁ«ﬁ", vbOKOnly, "—”«·…  «ﬂÌœ"
CommandButton20.Enabled = True
ElseIf result = vbNo Then
CheckBox2.value = False
CheckBox1.value = False
CommandButton13.Enabled = False
CommandButton20.Enabled = False
Exit Sub
End If

'=================================

End Sub

Private Sub ComboBox1_AfterUpdate()
ComboBox1.TextAlign = fmTextAlignCenter


End Sub

Private Sub ComboBox1_Change()

On Error GoTo emad
Sheets(1).Activate
Dim namecomb As Long
Dim namecombv As Long
Dim result As VbMsgBoxResult
Dim result2 As VbMsgBoxResult
Dim i As Integer
Dim i1 As Integer
Dim ss As Integer
Dim ss1 As Integer

 ss = Sheets(1).Cells(Rows.count, 2).End(xlUp).row
    For i = 2 To ss
' =VLOOKUP(BG4,$B$2:$W$28,4,0)
ss1 = Sheets(8).Cells(Rows.count, 2).End(xlUp).row
   
If ComboBox1.value = Sheets(1).Cells(i, "e") Then
TextBox32.Text = Sheets(1).Cells(i, "b")
result = MsgBox("«”„ «·„ÊŸ› „ÊÃÊœ ÷„‰ «·ﬁÌÊœ «·⁄«„Â Ê‰Ÿ«„ «·—Ê« » Â·  —Ìœ «· «ﬂœ „‰ ﬁÌœ «·⁄·«Ê« ", vbOKCancel + vbQuestion, "—”«·…  ÊÃÌÂ ⁄„«œ «·‰⁄Ì„Ì")
TextBox21.Text = TextBox32.Text
 If result = vbOK Then
    Sheets(8).Activate
     For i1 = 2 To ss1
If TextBox32.Text = Sheets(8).Cells(i1, "b").value And ComboBox1.value = Sheets(8).Cells(i1, "e") Then

result2 = MsgBox("«”„ «·„ÊŸ› „ÊÃÊœ ÷„‰ ﬁÌÊœ ‰Ÿ«„ «·⁄·«Ê«  Ê«· —›Ì⁄« ", vbYesNo + vbQuestion, "—”«·…  ÊÃÌÂ")

If result2 = vbYes Then
Frame4.Visible = False
Exit Sub
End If
End If

If TextBox32.Text <> Sheets(8).Cells(i1, "b").value And ComboBox1.value = Sheets(8).Cells(i1, "e").value Then
'If TextBox32.Text <> Sheets(8).Cells(i1, "b") And ComboBox1.Value <> Sheets(8).Cells(i1, "e").Value Then
MsgBox "⁄–—« «”„ «·„ÊŸ› €Ì— „ÊÃÊœ ÷„‰ ﬁÌÊœ ‰Ÿ«„ «·⁄·«Ê«  Ê«· —›Ì⁄«  Õ«·Ì« —«Ã⁄ Ê«ÃÂ… «· —ÕÌ· ", vbYesNo + vbQuestion, "—”«·…  ÊÃÌÂ"


End If
Next i1
End If
End If
Next i
MsgBox "⁄–—« «”„ «·„ÊŸ› €Ì— „ÊÃÊœ ÷„‰ ﬁÌÊœ ‰Ÿ«„ «·⁄·«Ê«  Ê«· —›Ì⁄«  Õ«·Ì« —«Ã⁄ Ê«ÃÂ… «· —ÕÌ· ", vbYesNo + vbQuestion, "—”«·…  ÊÃÌÂ"
emad:
End Sub
Private Sub CommandButton1_Click()
On Error Resume Next
'=================================
'====================================
Sheets(8).Activate
For Each c1 In Worksheets(8).Range("b:b")
If c1 = TextBox21.Text Then
TextBox2.value = c1

TextBox1.value = Application.WorksheetFunction.MaxIfs(Sheets(8).Range("a2:a10000"), Sheets(8).Range("b2:b10000"), Me.TextBox21.Text)
'=MAXIFS(A2:A8,B2:B8,Q4)
TextBox3.value = c1.Offset(0, 1)

TextBox4.value = c1.Offset(0, 2)

TextBox5.value = c1.Offset(0, 3)
TextBox6.value = c1.Offset(0, 4)
TextBox7.value = c1.Offset(0, 5)
TextBox8.value = c1.Offset(0, 6)
TextBox9.value = c1.Offset(0, 7)
TextBox10.value = c1.Offset(0, 8)
TextBox11.value = c1.Offset(0, 9)
TextBox12.value = c1.Offset(0, 10)
TextBox13.value = c1.Offset(0, 11)
TextBox14.value = c1.Offset(0, 12)
TextBox15.Text = c1.Offset(0, 13)
TextBox16.value = c1.Offset(0, 14)
TextBox17.Text = c1.Offset(0, 15)
TextBox18.Text = c1.Offset(0, 16)
 TextBox19.Text = c1.Offset(0, 17)
 TextBox44.Text = c1.Offset(0, 18)
 TextBox38.Text = c1.Offset(0, 19)
 TextBox39.Text = c1.Offset(0, 20)
 TextBox41.Text = c1.Offset(0, 21)
'TextBox42.Text = c1.Offset(0, 23)
 TextBox43.Text = c1.Offset(0, 22)
' TextBox44.Text = c1.Offset(0, 18)
'TextBox38.Text = c1.Offset(0, 19)
 'TextBox39.Text = c1.Offset(0, 20)
 
 
 
 
 
 
 
 
 '=============================
 
 
 
 
 '===========================
 Dim sdat As Date

sdate = Application.WorksheetFunction.EDate(Format(TextBox12.Text, "yyyy/mm/dd"), -1)

 TextBox26.Text = Format(TextBox12.Text, "yyyy/mm/dd")


TextBox27.value = Format(sdate, "yyyy/mm/dd")

TextBox30.value = Format(Date, "yyyy/mm/dd")
 
'===============================
Dim result As VbMsgBoxResult
Dim result1 As VbMsgBoxResult
  If CDate(TextBox30.Text) >= CDate(TextBox27.Text) And CDate(TextBox30.Text) <= CDate(TextBox26.Text) Then
' Or TextBox12.Text <= Application.WorksheetFunction.EDate(Date, 1 - 1.3)
result = MsgBox("Ì—ÃÏ  »·Ì€ «·„‰ ”» Ê ÊÕœ… «·⁄·«Ê«  Ê«· —›Ì⁄«  " & TextBox5.Text & Chr(10) & "»«‰  «—ÌŒ «” Õﬁ«ﬁ «·⁄·«Ê… » «—ÌŒ" & TextBox12.Text, vbYesNo + vbQuestion, "—”«·…  ÊÃÌÂ")
If result = vbYes Then
Exit Sub
 Else
 Dim contday As Integer
 
 contday = DateDiff("d", Format(CDate(TextBox30.Text), "yyyy/mm/dd"), Format(CDate(TextBox26.Text), "yyyy/mm/dd"))
 TextBox29.value = contday
result1 = MsgBox(" ‘ﬂ—« ·ﬁœ  „ «·€«¡ «· ÊÃÌÂ" & Chr(10) & "Ê·„ Ì »ﬁÏ „‰ «·Êﬁ  ··⁄·«ÊÂ" & Chr(10) & Chr(10) & contday & "" & "" & "«Ì«„")
 If result1 = vbOK Then
 Exit Sub
End If
End If
End If
'====================================Õ”«» «·«‘Â—
Dim nummonth As Integer
nummonth = DateDiff("m", TextBox11.Text, TextBox19.Text)
TextBox38.Text = nummonth
TextBox40.Text = " ‘Â— "
 TextBox39.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox11.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)
If TextBox39.value > 3 Then
TextBox38.Text = nummonth + 3
ElseIf TextBox39.Text <= 3 Then
TextBox38.Text = nummonth + TextBox39.value
TextBox40.Text = " ‘Â— "
End If
'===================================== Õ”«» «·”‰Ê« 
Dim nummonth1 As Integer
TextBox28.value = Format(sdate, "mm/dd/yyyy")
nummonth1 = DateDiff("m", TextBox13.Text, TextBox28.Text)
 TextBox43.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox13.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)
If TextBox43.value > 15 Then
TextBox41.Text = (nummonth1 + 15) / 12 & "Ê‘Â—" & (nummonth1 + TextBox43.value) Mod 12
ElseIf val(TextBox41.value) <= 15 Then
TextBox41.Text = (nummonth1 + TextBox43.value) / 12 & "Ê‘Â—" & (nummonth1 + TextBox43.value) Mod 12
TextBox37.Text = "”‰Â "
End If
If TextBox41.value >= 5 Then
 
 CheckBox2.value = True
  CheckBox2.Enabled = True
 CommandButton20.Enabled = True
 Else
 CheckBox2.value = False
  CheckBox2.Enabled = False
  CommandButton20.Enabled = False
End If


'=============================
 
 
 
 



End If
Next







'============================
'=COUNTIFS(L2:L15;">=" &S4;B2:B15;R4;E2:E15;Q4)

Dim fondnam As Long
If fondnam = Application.WorksheetFunction.CountIf(Sheets(8).Range("b2:b100000"), Me.TextBox21.value) Then
fondnam = 0
MsgBox "Â–« «·ﬁÌœ €Ì— „ÊÃÊœ Õ«Ê· «· «ﬂœ „‰Â «Ê ﬁœ ÌﬂÊ‰ Œÿ√ ›Ì ﬂ «» Â", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "—”«·Â  Õ–Ì—ÌÂ ⁄„«œ «·‰⁄Ì„Ì"

Exit Sub
End If

 '========================================
 '=======================================
TextBox26.Text = Format(TextBox26.Text, "yyyy/mm/dd")
TextBox27.Text = Format(TextBox27.Text, "yyyy/mm/dd")
TextBox11.Text = Format(TextBox11.Text, "yyyy/mm/dd")
TextBox12.Text = Format(TextBox12.Text, "yyyy/mm/dd")
TextBox13.Text = Format(TextBox13.Text, "yyyy/mm/dd")
TextBox14.Text = Format(TextBox14.Text, "yyyy/mm/dd")
TextBox16.Text = Format(TextBox16.Text, "yyyy/mm/dd")
TextBox17.Text = Format(TextBox17.Text, "yyyy/mm/dd")
TextBox24.Text = Format(TextBox24.Text, "yyyy/mm/dd")
TextBox25.Text = Format(TextBox25.Text, "yyyy/mm/dd")


 
End Sub

Private Sub CommandButton13_Click()
On Error Resume Next
'=INDEX(A2:O12,MATCH(G18,A2:A12,0),MATCH(F18,A2:O2,0))
 Dim sdat As Date
Dim numday As Integer
Dim nummonth As Integer

sdate = Application.WorksheetFunction.EDate(Format(Date, "yyyy/mm/dd"), 0)
TextBox19.value = Format(sdate, "yyyy/mm/dd")




nummonth = DateDiff("m", TextBox11.Text, TextBox19.Text)
TextBox38.Text = nummonth
TextBox40.Text = " ‘Â— "
 TextBox39.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox11.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)
If TextBox39.value > 3 Then
TextBox38.Text = nummonth + 3
ElseIf TextBox39.Text <= 3 Then
TextBox38.Text = nummonth + TextBox39.value
TextBox40.Text = " ‘Â— "
End If




'================================================


On Error Resume Next
Sheets(9).Activate
Dim seler As Integer
Dim seler1 As Integer
Dim rowrng As Integer
Dim colrng As Integer
Dim resultmsg As VbMsgBoxResult
seler1 = TextBox9.Text
If val(TextBox38.value) < 12 Or TextBox44.Text >= 2 Then

MsgBox "·« Ì„ﬂ‰ «÷«›… «” Õﬁ«ﬁ «·⁄·«ÊÂ ·«‰Â ·„ Ì Ã«Ê“ «·“„‰ «·›⁄·Ì ··«” Õﬁ«ﬁ", vbOKCancel, "—”«·…  ‰»ÌÂ"

Exit Sub
ElseIf TextBox38.value >= 12 Then


TextBox9.Text = val(seler1 + 1)
rowrng = Application.WorksheetFunction.Match(CLng(TextBox8.value), Sheets(9).Range("A2:A12"), 0)
colrng = Application.WorksheetFunction.Match(CLng(TextBox9.value), Sheets(9).Range("A2:O2"), 0)
TextBox33.Text = rowrng
TextBox34.Text = colrng
TextBox10.Text = Application.WorksheetFunction.Index(Sheets(9).Range("A2:O12"), rowrng, colrng)
'=====================================================
 For y1 = 2 To 2000
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
    Sheets(8).Cells(y1, 17) = TextBox17.Text
    Sheets(8).Cells(y1, 18) = TextBox18.Text
    Sheets(8).Cells(y1, 19) = TextBox19.Text
    Sheets(8).Cells(y1, 20) = TextBox44.Text
    Sheets(8).Cells(y1, 21) = TextBox38.Text
    Sheets(8).Cells(y1, 22) = TextBox39.Text
    Sheets(8).Cells(y1, 23) = TextBox42.Text
    Sheets(8).Cells(y1, 24) = TextBox41.Text
    Sheets(8).Cells(y1, 25) = TextBox25.Text
    Sheets(8).Cells(y1, 26) = TextBox28.Text

  resultmsg = MsgBox(" „  Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbCritical + vbOKOnly, "—”«·…  ÊÃÌÂ")

If resultmsg = vbOK Then
MsgBox "Â·  —Ìœ  —ÕÌ· «·⁄·«Ê…  «·Ï «·‰Ÿ«„ «·„«·Ì ··„ÊŸ› " & textbx5.Text, vbQuestion + vbOKCancel, "—”«·…  ÊÃÌÂ"
CommandButton16.Enabled = True
Else
CommandButton16.Enabled = False

End If





'=============================================
CheckBox1.value = False
CheckBox1.Enabled = False
CommandButton13.Enabled = False
Exit Sub
End If

'==========================    —ÕÌ· «·»Ì«‰« 
  
  
 

End Sub

Private Sub CommandButton14_Click()
For Y = 9 To 1500
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text And Sheets(1).Cells(Y, 2) = TextBox2.Text Then
    Exit For
    End If
    Next Y
    Sheets(1).Cells(Y, 1) = TextBox1.Text
    Sheets(1).Cells(Y, 2) = TextBox2.Text
    Sheets(1).Cells(Y, 3) = TextBox3.Text
    Sheets(1).Cells(Y, 4) = TextBox4.Text
    Sheets(1).Cells(Y, 5) = TextBox5.Text
    Sheets(1).Cells(Y, 6) = TextBox6.Text
    Sheets(1).Cells(Y, 7) = TextBox7.Text
    Sheets(1).Cells(Y, 8) = TextBox8.Text
    Sheets(1).Cells(Y, 9) = TextBox9.Text
    Sheets(1).Cells(Y, 10) = TextBox10.Text
    Sheets(1).Cells(Y, 54) = TextBox19.Text
   ' Sheets(1).Cells(Y, 12) = TextBox12.Text
  '   Sheets(1).Cells(Y, 58) = TextBox13.Text
MsgBox " „  —ÕÌ·Â« „«·Ì«"

End Sub

Private Sub CommandButton16_Click()
For Y = 9 To 1500
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text And Sheets(1).Cells(Y, 2) = TextBox2.Text Then
    Exit For
    End If
    Next Y
    Sheets(1).Cells(Y, 1) = TextBox1.Text
    Sheets(1).Cells(Y, 2) = TextBox2.Text
    Sheets(1).Cells(Y, 3) = TextBox3.Text
    Sheets(1).Cells(Y, 4) = TextBox4.Text
    Sheets(1).Cells(Y, 5) = TextBox5.Text
    Sheets(1).Cells(Y, 6) = TextBox6.Text
    Sheets(1).Cells(Y, 7) = TextBox7.Text
    Sheets(1).Cells(Y, 8) = TextBox8.Text
    Sheets(1).Cells(Y, 9) = TextBox9.Text
    Sheets(1).Cells(Y, 10) = TextBox10.Text
    Sheets(1).Cells(Y, 54) = TextBox19.Text
   ' Sheets(1).Cells(Y, 12) = TextBox12.Text
  '   Sheets(1).Cells(Y, 58) = TextBox13.Text
MsgBox " „  —ÕÌ·Â« „«·Ì«"

End Sub

Private Sub CommandButton18_Click()
On Error Resume Next
ListBox1.Clear
Sheets(6).Activate
   

    
   
    ss = Sheets(6).Cells(Rows.count, 2).End(xlUp).row
    For i = 2 To ss
    Dim daydate As Date
    daydate = Date
  '  If daydate >= Application.WorksheetFunction.EDate(Sheets(6).Cells(i, 13).Value, 1 - 1.2) Then
    If Sheets(6).Cells(i, "m").value >= Application.WorksheetFunction.EDate(CDate(TextBox24.Text), 0) And Sheets(6).Cells(i, "m").value <= Application.WorksheetFunction.EDate(CDate(TextBox24.Text), 1) Then
   'CDate(TextBox23.Text) Then
    ListBox1.ColumnHeads = True
    ListBox1.ColumnCount = 15
    'ListBox1.ColumnHeads = Array("«· ”·”·", "«·—ﬁ„ «·ÊŸÌ›Ì", "", "", "", "", "", "", "", "", "", "")
    Me.RightToLeft = True
    ListBox1.ColumnWidths = " 20;100;130;130;130;40;30;100;100;100;100;100;100;100"
        ListBox1.AddItem Sheets(6).Cells(i, 1).value
        ListBox1.List(ListBox1.ListCount - 1, 1) = Sheets(6).Cells(i, 2).value
        ListBox1.List(ListBox1.ListCount - 1, 2) = Sheets(6).Cells(i, 5).value
        ListBox1.List(ListBox1.ListCount - 1, 3) = Sheets(6).Cells(i, 6).value
        ListBox1.List(ListBox1.ListCount - 1, 4) = Sheets(6).Cells(i, 7).value
        ListBox1.List(ListBox1.ListCount - 1, 5) = Sheets(6).Cells(i, 8).value
        ListBox1.List(ListBox1.ListCount - 1, 6) = Sheets(6).Cells(i, 9).value
        ListBox1.List(ListBox1.ListCount - 1, 7) = Format(Sheets(6).Cells(i, 13).value, "yyyy/mm/dd")
        ListBox1.List(ListBox1.ListCount - 1, 8) = Sheets(6).Cells(i, 14).value
        ListBox1.List(ListBox1.ListCount - 1, 9) = Sheets(6).Cells(i, 15).value
        'ListBox1.List(ListBox1.ListCount - 1, 10) = Sheets(6).Cells(i, 11).Value
       ' ListBox1.List(ListBox1.ListCount - 1, 11) = Sheets(6).Cells(i, 12).Value
       ' ListBox1.List(ListBox1.ListCount - 1, 12) = Format(Sheets(6).Cells(i, 13).Value, "yyyy/mm/dd")
      '  ListBox1.List(ListBox1.ListCount - 1, 13) = Format(Sheets(6).Cells(i, 14).Value, "yyyy/mm/dd")
      '  ListBox1.List(ListBox1.ListCount - 1, 14) = Sheets(6).Cells(i, 15).Value
      '  ListBox1.List(ListBox1.ListCount - 1, 15) = Sheets(6).Cells(i, 16).Value
       ' ListBox1.List(ListBox1.ListCount - 1, 16) = Sheets(6).Cells(i, 17).Value
       ' ListBox1.List(ListBox1.ListCount - 1, 17) = Sheets(6).Cells(i, 18).Value
            
        
    End If
    
Next i

End Sub

Private Sub CommandButton19_Click()
'On Error Resume Next
Sheets(6).Activate
Dim rang As Range
Dim Data As Variant
Dim i As Long
Dim CNT, CLN As Integer
Me.RightToLeft = True
ListBox1.ColumnWidths = "50,80,80,40,140,100,30,30,100,100,100,100,100,100,100,100"
Set rang = Sheets(6).Range("a1").CurrentRegion
Me.ListBox1.Clear
ReDim Data(1 To rang.Rows.count, 1 To rang.Columns.count)
For i = 1 To rang.Rows.count
On Error Resume Next
If Sheets(6).Cells(i, "m") >= CDate(TextBox24.Text) And Sheets(6).Cells(i, "m") <= CDate(TextBox25.Text) Then
CNT = CNT + 1
For CLN = 1 To rang.Columns.count
Data(CNT + 1, CLN) = Sheets(6).Cells(i, CLN)
Next CLN
End If
Next i
Me.ListBox1.List = Data
Me.ListBox1.ColumnCount = CLN
Me.ListBox1.AddItem
For X = 1 To CLN
Me.ListBox1.List(0, X - 1) = Sheets(6).Cells(1, X)
Next X
Me.ListBox1.Selected(0) = True








End Sub

Private Sub CommandButton20_Click()
On Error Resume Next
'=INDEX(A2:O12,MATCH(G18,A2:A12,0),MATCH(F18,A2:O2,0))
 Dim sdat As Date
Dim numday As Integer
Dim nummonth As Integer
sdate = Application.WorksheetFunction.EDate(Format(Date, "yyyy/mm/dd"), 0)

 

TextBox28.value = Format(sdate, "yyyy/mm/dd")
nummonth = DateDiff("m", TextBox13.Text, TextBox28.Text)




TextBox37.Text = " ”‰Â "
'=================================================

 TextBox43.value = Application.WorksheetFunction.CountIfs(Sheets(6).Range("$l$2:$l$10000"), ">=" & Me.TextBox13.value, Sheets(6).Range("$b2:$b10000"), _
Me.TextBox2, Sheets(6).Range("$e2:$e10000"), Me.TextBox5.Text)
If TextBox43.value > 15 Then
TextBox41.Text = (nummonth + 15) / 12 & "Ê‘Â—" & (nummonth + TextBox43.value) Mod 12
ElseIf val(TextBox41.value) <= 15 Then
TextBox41.Text = (nummonth + TextBox43.value) / 12 & "Ê‘Â—" & (nummonth + TextBox43.value) Mod 12
TextBox37.Text = "”‰Â "
End If




'================================================


'On Error Resume Next
Sheets(9).Activate
Dim seler As Integer
Dim seler1 As Integer
Dim rowrng As Integer
Dim colrng As Integer
Dim rowrng1 As Integer
Dim colrng1 As Integer

seler1 = TextBox8.Text
If TextBox41.value < 5 Or TextBox44.Text >= 2 Then

MsgBox "·« Ì„ﬂ‰ «÷«›… «” Õﬁ«ﬁ «· —›Ì⁄ ·«‰Â ·„ Ì Ã«Ê“ «·“„‰ «·›⁄·Ì ··«” Õﬁ«ﬁ", vbOKCancel, "—”«·…  ‰»ÌÂ"


ElseIf val(TextBox41.value) >= 5 And TextBox8.Text >= 2 Then


TextBox8.Text = val(seler1 - 1)
TextBox9.Text = 1
rowrng = Application.WorksheetFunction.Match(CLng(TextBox8.value), Sheets(9).Range("A2:A12"), 0)
colrng = Application.WorksheetFunction.Match(CLng(TextBox9.value), Sheets(9).Range("A2:O2"), 0)

TextBox10.Text = Application.WorksheetFunction.Index(Sheets(9).Range("A2:O12"), rowrng, colrng)
'==========================================œ«·… «·Ê’› «·ÊŸÌ›Ì Ê  ‘€Ì· «·„’›Ê›Â
'=INDEX($A$1:L62,MATCH(D68,$A$1:$A$62,0),MATCH(C68,$A$1:$L$1,0))
rowrng1 = Application.WorksheetFunction.Match(TextBox18.Text, Sheets(10).Range("A1:A62"), 0)
colrng1 = Application.WorksheetFunction.Match(CLng(TextBox8.value), Sheets(10).Range("A1:L1"), 0)
TextBox6.Text = Application.WorksheetFunction.Index(Sheets(10).Range("A1:L62"), rowrng1, colrng1)
TextBox33.Text = rowrng1
TextBox34.Text = colrng1
CheckBox2.value = False
CheckBox2.Enabled = False
CommandButton20.Enabled = False

CommandButton14.Enabled = True
Else
MsgBox "«·ÕœÊœ «·⁄·Ì« ·· —›Ì⁄ ·«  ﬁ· ⁄‰ «·œ—Ã… 1"
TextBox8.Text = 1

End If
'=====================================================
 For y1 = 2 To 2000
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
    Sheets(8).Cells(y1, 17) = TextBox17.Text
    Sheets(8).Cells(y1, 18) = TextBox18.Text
    Sheets(8).Cells(y1, 19) = TextBox19.Text
    Sheets(8).Cells(y1, 20) = TextBox44.Text
    Sheets(8).Cells(y1, 21) = TextBox38.Text
    Sheets(8).Cells(y1, 22) = TextBox39.Text
    Sheets(8).Cells(y1, 23) = TextBox42.Text
    Sheets(8).Cells(y1, 24) = TextBox41.Text
    Sheets(8).Cells(y1, 25) = TextBox43.Text
    Sheets(8).Cells(y1, 26) = TextBox28.Text

MsgBox " „  ⁄„·Ì… Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbCritical + vbOKOnly, "—”«·…  ÊÃÌÂ"











'=============================================

Exit Sub


'==========================    —ÕÌ· «·»Ì«‰« 
  
  
End Sub

Private Sub CommandButton21_Click()
'On Error Resume Next
Frame4.Visible = True
 TextBox2.Text = ""
     TextBox3.Text = ""
     TextBox4.Text = ""
     TextBox5.Text = ""
     TextBox7.Text = ""
   TextBox6.Text = ""
  TextBox7.Text = ""
   TextBox8.Text = ""
  TextBox9.Text = ""
  TextBox10.Text = ""
   TextBox11.Text = ""
    TextBox12.Text = ""
      TextBox13.Text = ""
     TextBox14.Text = ""
     TextBox15.Text = ""
   TextBox16.Text = ""
   TextBox17.Text = ""
   TextBox18.Text = ""
    TextBox19.Text = ""
     TextBox21.Text = ""
  TextBox22.Text = ""
  TextBox24.Text = ""
   TextBox25.Text = ""
   TextBox26.Text = ""
   TextBox27.Text = ""
   TextBox28.Text = ""
  TextBox29.Text = ""
    TextBox31.Text = ""
  TextBox32.Text = ""
    TextBox44.Text = ""
   TextBox38.Text = ""
  TextBox39.Text = ""
    TextBox42.Text = ""
  TextBox41.Text = ""
   TextBox43.Text = ""
  
  
  
  
   SetFocus = TextBox32.Text
End Sub



Private Sub CommandButton22_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub Frame2_AddControl(ByVal Control As MSForms.Control)

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub TextBox11_AfterUpdate()
TextBox11.Text = Format(TextBox11.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox11_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
TextBox26.Text = Format(TextBox26.Text, "yyyy/mm/dd")
TextBox27.Text = Format(TextBox27.Text, "yyyy/mm/dd")
TextBox11.Text = Format(TextBox11.Text, "yyyy/mm/dd")
TextBox12.Text = Format(TextBox12.Text, "yyyy/mm/dd")
TextBox13.Text = Format(TextBox13.Text, "yyyy/mm/dd")
TextBox14.Text = Format(TextBox14.Text, "yyyy/mm/dd")
TextBox16.Text = Format(TextBox16.Text, "yyyy/mm/dd")
TextBox17.Text = Format(TextBox17.Text, "yyyy/mm/dd")
TextBox24.Text = Format(TextBox24.Text, "yyyy/mm/dd")
TextBox25.Text = Format(TextBox25.Text, "yyyy/mm/dd")


End Sub

Private Sub TextBox12_AfterUpdate()
TextBox12.Text = Format(TextBox12.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox13_AfterUpdate()
TextBox13.Text = Format(TextBox13.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox14_AfterUpdate()
TextBox14.Text = Format(TextBox14.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox16_AfterUpdate()
TextBox16.Text = Format(TextBox16.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox17_AfterUpdate()
TextBox17.Text = Format(TextBox17.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox22_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
TextBox21.Text = ""
End Sub

Private Sub TextBox22_Change()

Dim last1, frw1 As Integer

Sheets(8).Activate
last1 = Sheets(8).Range("A" & Rows.count).End(xlUp).row
     For frw1 = 2 To last1
 If Sheets(8).Cells(frw1, "e") Like "*" & TextBox22.Text & "*" Then
 TextBox21.Text = Sheets(8).Cells(frw1, "b")
End If
Next frw1
If last1 = 0 Then
TextBox21.Text = ""
MsgBox "Â–« «·ﬁÌœ €Ì— „ÊÃÊœ Õ«Ê· «· «ﬂœ „‰Â «Ê ﬁœ ÌﬂÊ‰ Œÿ√ ›Ì ﬂ «» Â", vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "—”«·Â  Õ–Ì—ÌÂ ⁄„«œ «·‰⁄Ì„Ì"
Exit Sub
End If
If TextBox22.Text = "" Then
TextBox21.Text = ""
End If
End Sub




Private Sub TextBox23_AfterUpdate()
TextBox23.Text = Format(TextBox23.Text, "yyyy/mm/dd")
End Sub

Private Sub TextBox23_Change()

End Sub



Private Sub TextBox22_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
CheckBox1.value = False
CheckBox2.value = False

End Sub

Private Sub TextBox24_Change()

End Sub

Private Sub TextBox31_AfterUpdate()
On Error Resume Next

Sheets(1).Activate
       If TextBox31 = "" Then Exit Sub
    Sheets(1).Activate

    ss = Sheets(1).Cells(Rows.count, 5).End(xlUp).row
     k = 0
     
     
For Each c In Range("E9:E" & ss)
    If c Like "*" & TextBox31.value & "*" Then
        ComboBox1.AddItem
      ComboBox1.List(k, 0) = Cells(c.row, 5).value
        ComboBox1.List(k, 1) = c.row
        k = k + 1
    End If
    
Next c
End Sub

Private Sub TextBox38_Change()
If TextBox38.value >= 12 Then
TextBox44.Text = 1
Else
Exit Sub
End If
End Sub

Private Sub TextBox9_AfterUpdate()
'=INDEX($A$2:O12,MATCH(G18,$A$2:$A$12,0),MATCH(F18,$A$2:$O$2,0))

End Sub

Private Sub UserForm_Activate()

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

TextBox26.Text = Format(TextBox26.Text, "yyyy/mm/dd")
TextBox27.Text = Format(TextBox27.Text, "yyyy/mm/dd")
TextBox11.Text = Format(TextBox11.Text, "yyyy/mm/dd")
TextBox12.Text = Format(TextBox12.Text, "yyyy/mm/dd")
TextBox13.Text = Format(TextBox13.Text, "yyyy/mm/dd")
TextBox14.Text = Format(TextBox14.Text, "yyyy/mm/dd")
TextBox16.Text = Format(TextBox16.Text, "yyyy/mm/dd")
TextBox17.Text = Format(TextBox17.Text, "yyyy/mm/dd")
TextBox24.Text = Format(TextBox24.Text, "yyyy/mm/dd")
TextBox25.Text = Format(TextBox25.Text, "yyyy/mm/dd")



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



Private Sub UserForm_Initialize()
On Error GoTo emad






On Error Resume Next
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
     For frw2 = 2 To last2
Me.ComboBox1.List = Sheets(1).Range("e9:e" & frw2).value
TextBox30.value = Format(Date, "yyyy/mm/dd")
'TextBox24.Visible = False

Next
'======================================
Sheets(8).Activate
'==================================
 With ListView1
 'Call rightToLeftListView
 
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

.ColumnHeaders.Add , , " «—ÌŒ «Œ— ⁄·«Ê…", 120

.ColumnHeaders.Add , , "  «—ÌŒ «” Õﬁ«ﬁ «·⁄·«Ê…", 120

.ColumnHeaders.Add , , "  «—ÌŒ «Œ—  —›Ì⁄", 100

.ColumnHeaders.Add , , " «—ÌŒ «” Õﬁ«ﬁ «· —›Ì⁄", 140

.ColumnHeaders.Add , , "⁄œœ «·”‰Ê«  «· Ì ﬁ÷«Â« ›Ì «· —›Ì⁄", 160

.ColumnHeaders.Add , , "  «—ÌŒ «· Ê’ÌÂ ›Ì «·⁄·«Ê… «·”‰ÊÌÂ", 160

.ColumnHeaders.Add , , " «—ÌŒ «· Ê’ÌÂ »«· —›Ì⁄", 120

.ColumnHeaders.Add , , "«·Ê’› «·ÊŸÌ›Ì", 90

.ColumnHeaders.Add , , " «—ÌŒ «Œ– «” Õﬁ«ﬁÂ", 120

.ColumnHeaders.Add , , "⁄œœ „ﬂ—«  «· ﬂ—«— ", 140

.ColumnHeaders.Add , , "⁄œœ «·«‘Â— «· Ì ﬁ÷«Â« ›Ì «·⁄·«Ê…", 190

.ColumnHeaders.Add , , "⁄œœ ﬂ » «·‘ﬂ— Ê«· ﬁœÌ—", 120

.ColumnHeaders.Add , , "⁄œœ „—«  «· ﬂ—«— ›Ì «· —›Ì⁄", 120

.ColumnHeaders.Add , , "⁄⁄œ ”‰Ê«  «· Ì ﬁ÷«Â« ›Ì «· —›Ì⁄ «·«Œ»—", 160

.ColumnHeaders.Add , , "⁄œœ «·‘ﬂ— Ê«· ﬁœÌ— „‰ «· —›Ì⁄ «·«ŒÌ—", 160

.ColumnHeaders.Add , , " «—ÌŒ «Œ—  —›Ì⁄", 120



'Setting  ??????? ??????
.FullRowSelect = True '?????? ???? ????
.Gridlines = True
.View = lvwReport  '??? ???? ??????
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ??? ????
.Font.Size = 9 '??? ????
.Font.Bold = False  '????
'Font.Name = "Calibri"


Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(8).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(8).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(8).Cells(frw1, "b")
item1.SubItems(2) = Sheets(8).Cells(frw1, "c")
item1.SubItems(3) = Sheets(8).Cells(frw1, "d")
item1.SubItems(4) = Sheets(8).Cells(frw1, "e")
item1.SubItems(5) = Sheets(8).Cells(frw1, "f")
item1.SubItems(6) = Sheets(8).Cells(frw1, "g")
item1.SubItems(7) = Sheets(8).Cells(frw1, "h")
item1.SubItems(8) = Sheets(8).Cells(frw1, "i")
item1.SubItems(9) = Sheets(8).Cells(frw1, "j")
item1.SubItems(10) = Sheets(8).Cells(frw1, "k")
item1.SubItems(11) = Sheets(8).Cells(frw1, "l")
item1.SubItems(12) = Sheets(8).Cells(frw1, "m")
item1.SubItems(13) = Sheets(8).Cells(frw1, "n")
item1.SubItems(14) = Sheets(8).Cells(frw1, "o")
item1.SubItems(15) = Sheets(8).Cells(frw1, "p")
item1.SubItems(16) = Sheets(8).Cells(frw1, "q")
item1.SubItems(17) = Sheets(8).Cells(frw1, "r")
item1.SubItems(18) = Sheets(8).Cells(frw1, "s")
item1.SubItems(19) = Sheets(8).Cells(frw1, "t")
item1.SubItems(20) = Sheets(8).Cells(frw1, "u")
item1.SubItems(21) = Sheets(8).Cells(frw1, "v")
item1.SubItems(22) = Sheets(8).Cells(frw1, "w")
item1.SubItems(23) = Sheets(8).Cells(frw1, "x")
item1.SubItems(24) = Sheets(8).Cells(frw1, "y")
item1.SubItems(25) = Sheets(8).Cells(frw1, "z")
item1.SubItems(26) = Sheets(8).Cells(frw1, "aa")
   
Next
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
End With


End With


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

        
TextBox26.Text = Format(TextBox26.Text, "yyyy/mm/dd")
TextBox27.Text = Format(TextBox27.Text, "yyyy/mm/dd")
TextBox11.Text = Format(TextBox11.Text, "yyyy/mm/dd")
TextBox12.Text = Format(TextBox12.Text, "yyyy/mm/dd")
TextBox13.Text = Format(TextBox13.Text, "yyyy/mm/dd")
TextBox14.Text = Format(TextBox14.Text, "yyyy/mm/dd")
TextBox16.Text = Format(TextBox16.Text, "yyyy/mm/dd")
TextBox17.Text = Format(TextBox17.Text, "yyyy/mm/dd")
TextBox24.Text = Format(TextBox24.Text, "yyyy/mm/dd")
TextBox25.Text = Format(TextBox25.Text, "yyyy/mm/dd")














emad:






End Sub
