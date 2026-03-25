VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm13 
   Caption         =   "«·‰Ÿ«„ «·«œ«—Ì"
   ClientHeight    =   11220
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   20760
   OleObjectBlob   =   "UserForm13.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm13"
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


Private Sub ComboBox1_Change()
TextBox5.Text = ComboBox1.value
ComboBox1.Visible = False

End Sub

Private Sub CommandButton1_Click()
Call emad
Application.DisplayAlerts = False
ThisWorkbook.Save
MsgBox " „ «·Õ›Ÿ Ê⁄„· ‰”Œ «Õ Ì«ÿÌ »‰Ã«Õ"

Application.Quit
End Sub

Private Sub CommandButton16_Click()
On Error GoTo emad
'==============================⁄œœ ”‰Ê«  «·Œœ„Â «·„” „—Ì‰ ›Ì «·Œœ„Â
If TextBox5.Text = "„” „—" Then
If TextBox6.Text = "" Then
MsgBox " «—ÌŒ «⁄«œ… «· ⁄Ì‰ €Ì— „ÊÃÊœ Õ«Ê· «œŒ«·  «—ÌŒ  ⁄Ì‰Â ·Ì ”‰Ï «·«Õ ”«»"
Exit Sub
End If
Dim xday3, xmonth3, xyear3 As Integer
Dim X3, y3 As Date
y3 = TextBox6.Text
X3 = Format(TextBox81.Text, "dd/mm/yyyy")
xyear3 = DateDiff("yyyy", y3, X3)
xmonth3 = DateDiff("m", y3, X3) Mod 12
Dim a3, b3 As Integer
b3 = Day(X3) - Day(y3)
xday3 = b3
If b3 < 0 Then
xmonth3 = xmonth3 - 1
xday3 = xday3 + 30
End If
a3 = Month(X3) - Month(y3)
If a3 < 0 Then
xyear3 = xyear3 - 1
End If
Me.TextBox22.value = xyear3 & "”‰…/" & xmonth3 & "‘Â—/" & xday3 & "ÌÊ„"

If b3 >= 15 Then
Me.TextBox22.value = xyear3 & "”‰…/" & xmonth3 & "‘Â—/" & xday3 & "ÌÊ„"
End If
If xmonth3 = -1 Then
Me.TextBox22.value = xyear3 - 1 & "”‰…/" & 12 + xmonth3 & "‘Â—/" & xday3 & "ÌÊ„"
End If
'======================================„›’Ê· ”Ì«”Ì

ElseIf TextBox5.Text = "„›’Ê· ”Ì«”Ì" Then
Dim xday5, xmonth5, xyear5 As Integer
Dim X5, y5 As Date
y5 = TextBox6.Text
X5 = TextBox80.Text
xyear5 = DateDiff("yyyy", y5, X5)
xmonth5 = DateDiff("m", y5, X5) Mod 12
Dim a5, b5 As Integer
b5 = Day(X5) - Day(y5)
xday5 = b5
If b5 < 0 Then
xmonth5 = xmonth5 - 1
xday5 = xday5 + 30
End If
a5 = Month(X5) - Month(y5)
If a5 < 0 Then
xyear5 = xyear5 - 1
End If
Me.TextBox18.value = xyear5 & "”‰…/" & xmonth5 & "‘Â—/" & xday5 & "ÌÊ„"
'---------------------
Dim xday6, xmonth6, xyear6 As Integer
Dim X6, y6 As Date
y6 = Format(TextBox10.Text, "dd/mm/yyyy")
X6 = Format(TextBox81.Text, "dd/mm/yyyy")
xyear6 = DateDiff("yyyy", y6, X6)
xmonth6 = DateDiff("m", y6, X6) Mod 12
Dim a6, b6 As Integer
b6 = Day(X6) - Day(y6)
xday6 = b6
If b6 < 0 Then

xmonth6 = xmonth6 - 1
xday6 = xday6 + 30
End If
a6 = Month(X6) - Month(y6)
If a6 < 0 Then
xyear6 = xyear6 - 1

End If
Me.TextBox17.value = xyear6 & "”‰…/" & xmonth6 & "‘Â—/" & xday6 & "ÌÊ„"

'-------------------------------
Dim xday7, xmonth7, xyear7 As Integer
Dim X7, y7 As Date
y7 = Sheets(19).Cells(2, 24)
X7 = Format(TextBox10.Text, "dd/mm/yyyy")
xyear7 = DateDiff("yyyy", y7, X7)
xmonth7 = DateDiff("m", y7, X7) Mod 12
Dim a7, b7 As Integer
b7 = Day(X7) - Day(y7)
xday7 = b7
If b7 < 0 Then
xmonth7 = xmonth7 - 1
xday7 = xday7 + 30
End If
a7 = Month(X7) - Month(y7)
If a7 < 0 Then
xyear7 = xyear7 - 1
End If
Me.TextBox16.value = xyear7 & "”‰…/" & xmonth7 & "‘Â—/" & xday7 & "ÌÊ„"

'------------------------------------------«Õ ”«» «·Œœ„… «·ﬂ·Ì… ··„›’Ê· «·”Ì«”Ì
Dim sumday3, summonth3, sumyear3 As Integer
Dim sumday4, summonth4, sumyear4 As Integer
Dim sumday5, summonth5, sumyear5 As Integer

sumday3 = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xday5, xday6), 30)
sumday4 = Application.WorksheetFunction.Sum(xday5, xday6) - sumday3
sumday5 = Application.WorksheetFunction.Sum(xday5, xday6)


summonth5 = Application.WorksheetFunction.Sum(xmonth5, xmonth6)
summonth3 = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xmonth5, xmonth6, (sumday5 / 30)), 12)
summonth4 = Application.WorksheetFunction.Round(Application.WorksheetFunction.Sum(summonth5, sumday5 / 30) - summonth3, 0)





sumyear3 = Application.WorksheetFunction.Sum(xyear5, xyear6, summonth3 / 12)
TextBox22.Text = sumyear3 & "”‰Â" & " / " & summonth4 & "‘Â—" & " / " & sumday4 & "ÌÊ„"
If sumday4 >= 15 Then
TextBox22.Text = sumyear3 & "”‰Â" & " / " & summonth4 - 1 & "‘Â—" & " / " & sumday4 & "ÌÊ„"
End If





'===============================«·„⁄«œÌ‰ ··Œœ„Â
ElseIf TextBox5.Text = "„⁄«œ" Then
'==============================⁄œœ ”‰Ê«  «·Œœ„Â ﬁ»· «·«‰ﬁÿ«⁄
Dim xday, xmonth, xyear As Integer
Dim X, Y As Date
Y = Format(TextBox6.Text, "dd/mm/yyyy")
X = Format(TextBox8.Text, "dd/mm/yyyy")
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
'==================================⁄œœ ”‰Ê«  «·«‰ﬁÿ«⁄
Dim xday1, xmonth1, xyear1 As Integer
Dim X1, y1 As Date
y1 = Format(TextBox8.Text, "dd/mm/yyyy")
X1 = Format(TextBox10.Text, "dd/mm/yyyy")
xyear1 = DateDiff("yyyy", y1, X1)
xmonth1 = DateDiff("m", y1, X1) Mod 12
Dim a1, b1 As Integer
b1 = Day(X1) - Day(y1)
xday1 = b1
If b1 < 0 Then
xmonth1 = xmonth1 - 1
xday1 = xday1 + 30
End If
a1 = Month(X1) - Month(y1)
If a1 < 0 Then
xyear1 = xyear1 - 1
End If
Me.TextBox16.value = xyear1 & "”‰…/" & xmonth1 & "‘Â—/" & xday1 & "ÌÊ„"
'==================================⁄œœ ”‰Ê«  «·Œœ„Â »⁄œ «⁄«œ… «· ⁄Ì‰
Dim xday2, xmonth2, xyear2 As Integer
Dim X2, y2 As Date
y2 = Format(TextBox10.Text, "dd/mm/yyyy")
X2 = Format(TextBox81.Text, "dd/mm/yyyy")
xyear2 = DateDiff("yyyy", y2, X2)
xmonth2 = DateDiff("m", y2, X2) Mod 12
Dim a2, b2 As Integer
b2 = Day(X2) - Day(y2)
xday2 = b2
If b2 < 0 Then
xmonth2 = xmonth2 - 1
xday2 = xday2 + 30
End If
a2 = Month(X2) - Month(y2)
If a2 < 0 Then
xyear2 = xyear2 - 1
End If
Me.TextBox17.value = xyear2 & "”‰…/" & xmonth2 & "‘Â—/" & xday2 & "ÌÊ„"
'================================= ··„⁄«œÌ‰ ⁄œœ ”‰Ê«  «·Œœ„Â
Dim sumday, summonth, sumyear As Integer
Dim sumday1, summonth1, sumyear1 As Integer
Dim sumday2, summonth2, sumyear2 As Integer

sumday = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xday, xday2), 30)
sumday1 = Application.WorksheetFunction.Sum(xday, xday2) - sumday
sumday2 = Application.WorksheetFunction.Sum(xday, xday2)


summonth2 = Application.WorksheetFunction.Sum(xmonth, xmonth2)
summonth = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xmonth, xmonth2, (sumday2 / 30)), 12)
summonth1 = Application.WorksheetFunction.Round(Application.WorksheetFunction.Sum(summonth2, sumday2 / 30) - summonth, 0)





sumyear = Application.WorksheetFunction.Sum(xyear, xyear2, summonth / 12)
TextBox22.Text = sumyear & "”‰Â" & " / " & summonth1 & "‘Â—" & " / " & sumday1 & "ÌÊ„"
If sumday1 >= 15 Then
TextBox22.Text = sumyear & "”‰Â" & " / " & summonth1 - 1 & "‘Â—" & " / " & sumday1 & "ÌÊ„"


End If

'=================================⁄œœ ”‰Ê«  «·Œœ„Â «·ﬂ·ÌÂ

End If


ThisWorkbook.Save



emad:
End Sub

Private Sub CommandButton17_Click()
Call emad1
End Sub

Private Sub CommandButton2_Click()
On Error GoTo emad
 TextBox81.Enabled = False
  TextBox16.Enabled = False
  TextBox18.Enabled = False
  TextBox20.Enabled = False
 
 
    
    
    Dim iRow As Long, i As Long

    Sheets(19).Activate

    iRow = Range("A" & Rows.count).End(xlUp).row
'=======================================================
  ' Format(textbox81.Text, "dd/mm/yyyy") = 0                                              '
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
  '==============================
  If TextBox5.Text = "„” „—" Then
If TextBox6.Text = "" Then
MsgBox " «—ÌŒ «⁄«œ… «· ⁄Ì‰ €Ì— „ÊÃÊœ Õ«Ê· «œŒ«·  «—ÌŒ  ⁄Ì‰Â ·Ì ”‰Ï «·«Õ ”«»"
Exit Sub
End If
Dim xday3, xmonth3, xyear3 As Integer
Dim X3, y3 As Date
y3 = TextBox6.Text
X3 = Format(TextBox81.Text, "dd/mm/yyyy")
xyear3 = DateDiff("yyyy", y3, X3)
xmonth3 = DateDiff("m", y3, X3) Mod 12
Dim a3, b3 As Integer
b3 = Day(X3) - Day(y3)
xday3 = b3
If b3 < 0 Then
xmonth3 = xmonth3 - 1
xday3 = xday3 + 30
End If
a3 = Month(X3) - Month(y3)
If a3 < 0 Then
xyear3 = xyear3 - 1
End If
Me.TextBox22.value = xyear3 & "”‰…/" & xmonth3 & "‘Â—/" & xday3 & "ÌÊ„"
'======================================„›’Ê· ”Ì«”Ì

ElseIf TextBox5.Text = "„›’Ê· ”Ì«”Ì" Then
Dim xday5, xmonth5, xyear5 As Integer
Dim X5, y5 As Date
y5 = TextBox6.Text
X5 = #8/4/2003#

xyear5 = DateDiff("yyyy", y5, X5)
xmonth5 = DateDiff("m", y5, X5) Mod 12
Dim a5, b5 As Integer
b5 = Day(X5) - Day(y5)
xday5 = b5
If b5 < 0 Then
xmonth5 = xmonth5 - 1
xday5 = xday5 + 30
End If
a5 = Month(X5) - Month(y5)
If a5 < 0 Then
xyear5 = xyear5 - 1
End If
Me.TextBox18.value = xyear5 & "”‰…/" & xmonth5 & "‘Â—/" & xday5 & "ÌÊ„"
'---------------------
Dim xday6, xmonth6, xyear6 As Integer
Dim X6, y6 As Date
y6 = TextBox10.Text
X6 = Format(TextBox81.Text, "dd/mm/yyyy")
xyear6 = DateDiff("yyyy", y6, X6)
xmonth6 = DateDiff("m", y6, X6) Mod 12
Dim a6, b6 As Integer
b6 = Day(X6) - Day(y6)
xday6 = b6
If b6 < 0 Then
xmonth6 = xmonth6 - 1
xday6 = xday6 + 30
End If
a6 = Month(X6) - Month(y6)
If a6 < 0 Then
xyear6 = xyear6 - 1
End If
Me.TextBox17.value = xyear6 & "”‰…/" & xmonth6 & "‘Â—/" & xday6 & "ÌÊ„"
'-------------------------------
Dim xday7, xmonth7, xyear7 As Integer
Dim X7, y7 As Date
y7 = #8/4/2003#
X7 = TextBox10.Text
xyear7 = DateDiff("yyyy", y7, X7)
xmonth7 = DateDiff("m", y7, X7) Mod 12
Dim a7, b7 As Integer
b7 = Day(X7) - Day(y7)
xday7 = b7
If b7 < 0 Then
xmonth7 = xmonth7 - 1
xday7 = xday7 + 30
End If
a7 = Month(X7) - Month(y7)
If a7 < 0 Then
xyear7 = xyear7 - 1
End If
Me.TextBox16.value = xyear7 & "”‰…/" & xmonth7 & "‘Â—/" & xday7 & "ÌÊ„"


'------------------------------------------«Õ ”«» «·Œœ„… «·ﬂ·Ì… ··„›’Ê· «·”Ì«”Ì
Dim sumday3, summonth3, sumyear3 As Integer
Dim sumday4, summonth4, sumyear4 As Integer
Dim sumday5, summonth5, sumyear5 As Integer

sumday3 = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xday5, xday6), 30)
sumday4 = Application.WorksheetFunction.Sum(xday5, xday6) - sumday3
sumday5 = Application.WorksheetFunction.Sum(xday5, xday6)


summonth5 = Application.WorksheetFunction.Sum(xmonth5, xmonth6)
summonth3 = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xmonth5, xmonth6, (sumday5 / 30)), 12)
summonth4 = Application.WorksheetFunction.Round(Application.WorksheetFunction.Sum(summonth5, sumday5 / 30) - summonth3, 0)





sumyear3 = Application.WorksheetFunction.Sum(xyear5, xyear6, summonth3 / 12)
TextBox22.Text = sumyear3 & "”‰Â" & " / " & summonth4 & "‘Â—" & " / " & sumday4 & "ÌÊ„"








'===============================«·„⁄«œÌ‰ ··Œœ„Â
Else
'==============================⁄œœ ”‰Ê«  «·Œœ„Â ﬁ»· «·«‰ﬁÿ«⁄
Dim xday, xmonth, xyear As Integer
Dim X, Y As Date
Y = TextBox6.Text
X = TextBox8.Text
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
'==================================⁄œœ ”‰Ê«  «·«‰ﬁÿ«⁄
Dim xday1, xmonth1, xyear1 As Integer
Dim X1, y1 As Date
y1 = TextBox8.Text
X1 = TextBox10.Text
xyear1 = DateDiff("yyyy", y1, X1)
xmonth1 = DateDiff("m", y1, X1) Mod 12
Dim a1, b1 As Integer
b1 = Day(X1) - Day(y1)
xday1 = b1
If b1 < 0 Then
xmonth1 = xmonth1 - 1
xday1 = xday1 + 30
End If
a1 = Month(X1) - Month(y1)
If a1 < 0 Then
xyear1 = xyear1 - 1
End If
Me.TextBox16.value = xyear1 & "”‰…/" & xmonth1 & "‘Â—/" & xday1 & "ÌÊ„"
'==================================⁄œœ ”‰Ê«  «·Œœ„Â »⁄œ «⁄«œ… «· ⁄Ì‰
Dim xday2, xmonth2, xyear2 As Integer
Dim X2, y2 As Date
y2 = TextBox10.Text
X2 = Format(TextBox81.Text, "dd/mm/yyyy")
xyear2 = DateDiff("yyyy", y2, X2)
xmonth2 = DateDiff("m", y2, X2) Mod 12
Dim a2, b2 As Integer
b2 = Day(X2) - Day(y2)
xday2 = b2
If b2 < 0 Then
xmonth2 = xmonth2 - 1
xday2 = xday2 + 30
End If
a2 = Month(X2) - Month(y2)
If a2 < 0 Then
xyear2 = xyear2 - 1
End If
Me.TextBox17.value = xyear2 & "”‰…/" & xmonth2 & "‘Â—/" & xday2 & "ÌÊ„"
'================================= ··„⁄«œÌ‰ ⁄œœ ”‰Ê«  «·Œœ„Â
Dim sumday, summonth, sumyear As Integer
Dim sumday1, summonth1, sumyear1 As Integer
Dim sumday2, summonth2, sumyear2 As Integer

sumday = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xday, xday2), 30)
sumday1 = Application.WorksheetFunction.Sum(xday, xday2) - sumday
sumday2 = Application.WorksheetFunction.Sum(xday, xday2)


summonth2 = Application.WorksheetFunction.Sum(xmonth, xmonth2)
summonth = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xmonth, xmonth2, (sumday2 / 30)), 12)
summonth1 = Application.WorksheetFunction.Round(Application.WorksheetFunction.Sum(summonth2, sumday2 / 30) - summonth, 0)





sumyear = Application.WorksheetFunction.Sum(xyear, xyear2, summonth / 12)
TextBox22.Text = sumyear & "”‰Â" & " / " & summonth1 & "‘Â—" & " / " & sumday1 & "ÌÊ„"


'=================================⁄œœ ”‰Ê«  «·Œœ„Â «·ﬂ·ÌÂ

End If

Sheets(19).Activate
'=========================================================
    Range("A" & iRow + 1).Offset(0, 0).value = TextBox1.value
    Range("A" & iRow + 1).Offset(0, 1).value = TextBox2.value
    Range("A" & iRow + 1).Offset(0, 2).value = TextBox3.value
    Range("A" & iRow + 1).Offset(0, 3).value = TextBox4.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox5.value
    Range("A" & iRow + 1).Offset(0, 5).value = TextBox6.value
    Range("A" & iRow + 1).Offset(0, 6).value = TextBox7.value
    Range("A" & iRow + 1).Offset(0, 7).value = TextBox8.value
    Range("A" & iRow + 1).Offset(0, 8).value = TextBox9.value
    Range("A" & iRow + 1).Offset(0, 9).value = TextBox10.value
    Range("A" & iRow + 1).Offset(0, 10).value = TextBox11.value
    Range("A" & iRow + 1).Offset(0, 11).value = TextBox12.value
    Range("A" & iRow + 1).Offset(0, 12).value = TextBox13.value
    Range("A" & iRow + 1).Offset(0, 13).value = TextBox81.value
    Range("A" & iRow + 1).Offset(0, 14).value = TextBox15.value
    Range("A" & iRow + 1).Offset(0, 15).value = TextBox16.value
    Range("A" & iRow + 1).Offset(0, 16).value = TextBox17.value
    Range("A" & iRow + 1).Offset(0, 17).value = TextBox18.value
    Range("A" & iRow + 1).Offset(0, 18).value = TextBox19.value
    Range("A" & iRow + 1).Offset(0, 19).value = TextBox20.value
    Range("A" & iRow + 1).Offset(0, 20).value = TextBox21.value
    Range("A" & iRow + 1).Offset(0, 21).value = TextBox22.value
   ' Range("A" & iRow + 1).Offset(0, 22).Value = TextBox23.Value
    'Range("A" & iRow + 1).Offset(0, 23).Value = TextBox24.Value
   ' Range("A" & iRow + 1).Offset(0, 24).Value = TextBox25.Value
  
   
   
        TextBox1.value = ""
   TextBox2.value = ""
    TextBox3.value = ""
    TextBox4.value = ""
    TextBox5.value = ""
   TextBox6.value = ""
     TextBox7.value = ""
    TextBox8.value = ""
 TextBox9.value = ""
     TextBox10.value = ""
    TextBox11.value = ""
    TextBox12.value = ""
  TextBox13.value = ""
    TextBox81.value = ""
 TextBox15.value = ""
   TextBox16.value = ""
    TextBox17.value = ""
   TextBox18.value = ""
  TextBox19.value = ""
 TextBox20.value = ""
  TextBox21.value = ""
   TextBox22.value = ""
   TextBox77.value = ""
emad:
End Sub

Private Sub CommandButton20_Click()
On Error GoTo emad






'=======================================================
Dim lastR As Integer
lastR = Sheets(19).Cells(Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(19).Cells(Y, 1) = TextBox1.Text And Sheets(19).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(19).Cells(Y, 1) = TextBox1.Text
    Sheets(19).Cells(Y, 2) = TextBox2.Text
    Sheets(19).Cells(Y, 3) = TextBox3.Text
    Sheets(19).Cells(Y, 4) = TextBox4.Text
    Sheets(19).Cells(Y, 5) = TextBox5.Text
    Sheets(19).Cells(Y, 6) = Format(TextBox6.Text, "yyyy/mm/dd")
    Sheets(19).Cells(Y, 7) = TextBox7.Text
    Sheets(19).Cells(Y, 8) = TextBox8.Text
    Sheets(19).Cells(Y, 9) = TextBox9.Text
    Sheets(19).Cells(Y, 10) = Format(TextBox10.Text, "yyyy/mm/dd")
    Sheets(19).Cells(Y, 11) = TextBox11.Text
    Sheets(19).Cells(Y, 12) = TextBox12.Text
     Sheets(19).Cells(Y, 13) = TextBox13.Text
    Sheets(19).Cells(Y, 14) = Format(TextBox81.Text, "dd/mm/yyyy")
    Sheets(19).Cells(Y, 15) = TextBox15.Text
    Sheets(19).Cells(Y, 16) = TextBox16.Text
    Sheets(19).Cells(Y, 17) = TextBox17.Text
    Sheets(19).Cells(Y, 18) = TextBox18.Text
    Sheets(19).Cells(Y, 19) = TextBox19.Text
    Sheets(19).Cells(Y, 20) = TextBox20.Text
    Sheets(19).Cells(Y, 21) = TextBox21.Text
    Sheets(19).Cells(Y, 22) = TextBox22.Text
   
   ThisWorkbook.Save
   
        
emad:
End Sub

Private Sub CommandButton23_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton25_Click()
On Error GoTo emad
If TextBox5.Text = "„” „—" Then
If TextBox6.Text = "" Then
MsgBox " «—ÌŒ «⁄«œ… «· ⁄Ì‰ €Ì— „ÊÃÊœ Õ«Ê· «œŒ«·  «—ÌŒ  ⁄Ì‰Â ·Ì ”‰Ï «·«Õ ”«»"
Exit Sub
End If
Dim xday3, xmonth3, xyear3 As Integer
Dim X3, y3 As Date
y3 = TextBox6.Text
X3 = TextBox78.Text
xyear3 = DateDiff("yyyy", y3, X3)
xmonth3 = DateDiff("m", y3, X3) Mod 12
Dim a3, b3 As Integer
b3 = Day(X3) - Day(y3)
xday3 = b3
If b3 < 0 Then
xmonth3 = xmonth3 - 1
xday3 = xday3 + 30
End If
a3 = Month(X3) - Month(y3)
If a3 < 0 Then
xyear3 = xyear3 - 1
End If
Me.TextBox19.value = xyear3 & "”‰…/" & xmonth3 & "‘Â—/" & xday3 & "ÌÊ„"
'======================================„›’Ê· ”Ì«”Ì

ElseIf TextBox5.Text = "„›’Ê· ”Ì«”Ì" Then
Dim xday5, xmonth5, xyear5 As Integer
Dim X5, y5 As Date
y5 = TextBox6.Text
X5 = Sheets(19).Cells(2, 24)

xyear5 = DateDiff("yyyy", y5, X5)
xmonth5 = DateDiff("m", y5, X5) Mod 12
Dim a5, b5 As Integer
b5 = Day(X5) - Day(y5)
xday5 = b5
If b5 < 0 Then
xmonth5 = xmonth5 - 1
xday5 = xday5 + 30
End If
a5 = Month(X5) - Month(y5)
If a5 < 0 Then
xyear5 = xyear5 - 1
End If
Me.TextBox18.value = xyear5 & "”‰…/" & xmonth5 & "‘Â—/" & xday5 & "ÌÊ„"
'---------------------
Dim xday6, xmonth6, xyear6 As Integer
Dim X6, y6 As Date
y6 = TextBox10.Text
X6 = TextBox78.Text
xyear6 = DateDiff("yyyy", y6, X6)
xmonth6 = DateDiff("m", y6, X6) Mod 12
Dim a6, b6 As Integer
b6 = Day(X6) - Day(y6)
xday6 = b6
If b6 < 0 Then
xmonth6 = xmonth6 - 1
xday6 = xday6 + 30
End If
a6 = Month(X6) - Month(y6)
If a6 < 0 Then
xyear6 = xyear6 - 1
End If
Me.TextBox17.value = xyear6 & "”‰…/" & xmonth6 & "‘Â—/" & xday6 & "ÌÊ„"
'-------------------------------
Dim xday7, xmonth7, xyear7 As Integer
Dim X7, y7 As Date
y7 = #8/4/2003#
X7 = TextBox10.Text
xyear7 = DateDiff("yyyy", y7, X7)
xmonth7 = DateDiff("m", y7, X7) Mod 12
Dim a7, b7 As Integer
b7 = Day(X7) - Day(y7)
xday7 = b7
If b7 < 0 Then
xmonth7 = xmonth7 - 1
xday7 = xday7 + 30
End If
a7 = Month(X7) - Month(y7)
If a7 < 0 Then
xyear7 = xyear7 - 1
End If
Me.TextBox16.value = xyear7 & "”‰…/" & xmonth7 & "‘Â—/" & xday7 & "ÌÊ„"


'------------------------------------------«Õ ”«» «·Œœ„… «·ﬂ·Ì… ··„›’Ê· «·”Ì«”Ì
Dim sumday3, summonth3, sumyear3 As Integer
Dim sumday4, summonth4, sumyear4 As Integer
Dim sumday5, summonth5, sumyear5 As Integer

sumday3 = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xday5, xday6), 30)
sumday4 = Application.WorksheetFunction.Sum(xday5, xday6) - sumday3
sumday5 = Application.WorksheetFunction.Sum(xday5, xday6)


summonth5 = Application.WorksheetFunction.Sum(xmonth5, xmonth6)
summonth3 = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xmonth5, xmonth6, (sumday5 / 30)), 12)
summonth4 = Application.WorksheetFunction.Round(Application.WorksheetFunction.Sum(summonth5, sumday5 / 30) - summonth3, 0)





sumyear3 = Application.WorksheetFunction.Sum(xyear5, xyear6, summonth3 / 12)
TextBox19.Text = sumyear3 & "”‰Â" & " / " & summonth4 & "‘Â—" & " / " & sumday4 & "ÌÊ„"








'===============================«·„⁄«œÌ‰ ··Œœ„Â
Else
'==============================⁄œœ ”‰Ê«  «·Œœ„Â ﬁ»· «·«‰ﬁÿ«⁄
Dim xday, xmonth, xyear As Integer
Dim X, Y As Date
Y = TextBox6.Text
X = TextBox8.Text
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
'==================================⁄œœ ”‰Ê«  «·«‰ﬁÿ«⁄
Dim xday1, xmonth1, xyear1 As Integer
Dim X1, y1 As Date
y1 = TextBox8.Text
X1 = TextBox10.Text
xyear1 = DateDiff("yyyy", y1, X1)
xmonth1 = DateDiff("m", y1, X1) Mod 12
Dim a1, b1 As Integer
b1 = Day(X1) - Day(y1)
xday1 = b1
If b1 < 0 Then
xmonth1 = xmonth1 - 1
xday1 = xday1 + 30
End If
a1 = Month(X1) - Month(y1)
If a1 < 0 Then
xyear1 = xyear1 - 1
End If
Me.TextBox16.value = xyear1 & "”‰…/" & xmonth1 & "‘Â—/" & xday1 & "ÌÊ„"
'==================================⁄œœ ”‰Ê«  «·Œœ„Â »⁄œ «⁄«œ… «· ⁄Ì‰
Dim xday2, xmonth2, xyear2 As Integer
Dim X2, y2 As Date
y2 = TextBox10.Text
X2 = TextBox78.Text
xyear2 = DateDiff("yyyy", y2, X2)
xmonth2 = DateDiff("m", y2, X2) Mod 12
Dim a2, b2 As Integer
b2 = Day(X2) - Day(y2)
xday2 = b2
If b2 < 0 Then
xmonth2 = xmonth2 - 1
xday2 = xday2 + 30
End If
a2 = Month(X2) - Month(y2)
If a2 < 0 Then
xyear2 = xyear2 - 1
End If
Me.TextBox17.value = xyear2 & "”‰…/" & xmonth2 & "‘Â—/" & xday2 & "ÌÊ„"
'================================= ··„⁄«œÌ‰ ⁄œœ ”‰Ê«  «·Œœ„Â
Dim sumday, summonth, sumyear As Integer
Dim sumday1, summonth1, sumyear1 As Integer
Dim sumday2, summonth2, sumyear2 As Integer

sumday = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xday, xday2), 30)
sumday1 = Application.WorksheetFunction.Sum(xday, xday2) - sumday
sumday2 = Application.WorksheetFunction.Sum(xday, xday2)


summonth2 = Application.WorksheetFunction.Sum(xmonth, xmonth2)
summonth = Application.WorksheetFunction.Floor(Application.WorksheetFunction.Sum(xmonth, xmonth2, (sumday2 / 30)), 12)
summonth1 = Application.WorksheetFunction.Round(Application.WorksheetFunction.Sum(summonth2, sumday2 / 30) - summonth, 0)





sumyear = Application.WorksheetFunction.Sum(xyear, xyear2, summonth / 12)
TextBox19.Text = sumyear & "”‰Â" & " / " & summonth1 & "‘Â—" & " / " & sumday1 & "ÌÊ„"


'=================================⁄œœ ”‰Ê«  «·Œœ„Â «·ﬂ·ÌÂ

End If





emad:
End Sub

Private Sub CommandButton26_Click()
Sheets(19).Activate
Dim lastR, r As Integer
lastR = Sheets(19).Cells(Rows.count, 1).End(xlUp).row
For r = 2 To lastR
If TextBox1.Text = Sheets(19).Cells(r, 1).value And TextBox2.Text = Sheets(19).Cells(r, 2).value Then

Sheets(19).Cells(r, 1).Select
Exit For
End If
Next
ActiveCell.EntireRow.Resize(, 24).Delete
MsgBox " „ Õ–› «·”Ã· »‰Ã«Õ", vbMsgBoxRight + vbOKOnly, "⁄„«œ «·‰⁄Ì„Ì"

End Sub

Private Sub CommandButton3_Click()
UserForm14.Show
End Sub

Private Sub CommandButton4_Click()
'========================================================
'Data migration
 
  
    
    
    Dim iRow As Long, i As Long

    Sheets(24).Activate

    iRow = Range("A" & Rows.count).End(xlUp).row
'=======================================================
  ' Format(textbox81.Text, "dd/mm/yyyy") = 0                                              '
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
    Range("A" & iRow + 1).Offset(0, 3).value = TextBox4.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox5.value
    Range("A" & iRow + 1).Offset(0, 5).value = TextBox6.value
    Range("A" & iRow + 1).Offset(0, 6).value = TextBox7.value
    Range("A" & iRow + 1).Offset(0, 7).value = TextBox8.value
    Range("A" & iRow + 1).Offset(0, 8).value = TextBox9.value
    Range("A" & iRow + 1).Offset(0, 9).value = TextBox10.value
    Range("A" & iRow + 1).Offset(0, 10).value = TextBox11.value
    Range("A" & iRow + 1).Offset(0, 11).value = TextBox12.value
    Range("A" & iRow + 1).Offset(0, 12).value = TextBox13.value
    Range("A" & iRow + 1).Offset(0, 13).value = TextBox81.value
    Range("A" & iRow + 1).Offset(0, 14).value = TextBox15.value
    Range("A" & iRow + 1).Offset(0, 15).value = TextBox16.value
    Range("A" & iRow + 1).Offset(0, 16).value = TextBox17.value
    Range("A" & iRow + 1).Offset(0, 17).value = TextBox18.value
    Range("A" & iRow + 1).Offset(0, 18).value = TextBox19.value
    Range("A" & iRow + 1).Offset(0, 19).value = TextBox20.value
    Range("A" & iRow + 1).Offset(0, 20).value = TextBox21.value
    Range("A" & iRow + 1).Offset(0, 21).value = TextBox22.value
   Range("A" & iRow + 1).Offset(0, 23).value = TextBox79.value
    
    
        Sheets(19).Activate
'============================================================================


Sheets(19).Activate
Dim lastR, r As Integer
lastR = Sheets(19).Cells(Rows.count, 1).End(xlUp).row
For r = 2 To lastR
If TextBox1.Text = Sheets(19).Cells(r, 1).value And TextBox2.Text = Sheets(19).Cells(r, 2).value Then

Sheets(19).Cells(r, 1).Select
Exit For
End If
Next
ActiveCell.EntireRow.Resize(, 24).Delete
MsgBox " „ Õ–› «·”Ã· »‰Ã«Õ", vbMsgBoxRight + vbOKOnly, "⁄„«œ «·‰⁄Ì„Ì"
ThisWorkbook.Save
End Sub

Private Sub CommandButton6_Click()
'On Error GoTo emad
 
'========================================================

     TextBox2.Text = 0
     TextBox3.Text = 0
     TextBox4.Text = ""
     TextBox5.Text = ""
     TextBox7.Text = ""
   TextBox6.Text = ""
  TextBox7.Text = ""
   TextBox8.Text = 0
  TextBox9.Text = 0
  TextBox10.Text = 0
   TextBox11.Text = 0
    TextBox12.Text = 0
      TextBox13.Text = 0
     Format(TextBox81.Text, "dd/mm/yyyy") = ""
     TextBox15.Text = 0
   TextBox16.Text = ""
   TextBox17.Text = 0
   TextBox18.Text = ""
    TextBox19.Text = 0
   TextBox20.Text = ""
    TextBox21.Text = 0
  TextBox22.Text = 0
  

  
 
'============================================================
   

  Dim iRow As Long

    Sheets(19).Activate
    iRow = Sheets(19).Cells(Rows.count, 2).End(xlUp).row
    TextBox1.value = iRow
   
'emad:
End Sub

Private Sub CommandButton9_Click()
'On Error GoTo emad
Dim item1 As ListItem
Dim last1, frw1 As Integer

 With Me.ListView1
.ListItems.Clear













'====================================
Sheets(19).Activate
last1 = Sheets(19).Range("A" & Rows.count).End(xlUp).row



'=============================================

   Sheets(19).Activate
    
    
    For frw1 = 2 To last1
 
 If Sheets(19).Cells(frw1, "e").Text = ComboBox2.Text Then
     
'==================================================

Set item1 = ListView1.ListItems.Add(, , Sheets(19).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(19).Cells(frw1, "b")
item1.SubItems(2) = Sheets(19).Cells(frw1, "c")
item1.SubItems(3) = Sheets(19).Cells(frw1, "d")
item1.SubItems(4) = Sheets(19).Cells(frw1, "e")
item1.SubItems(5) = Sheets(19).Cells(frw1, "f")
item1.SubItems(6) = Sheets(19).Cells(frw1, "g")
item1.SubItems(7) = Sheets(19).Cells(frw1, "h")
item1.SubItems(8) = Sheets(19).Cells(frw1, "i")
item1.SubItems(9) = Sheets(19).Cells(frw1, "j")
item1.SubItems(10) = Sheets(19).Cells(frw1, "k")
item1.SubItems(11) = Sheets(19).Cells(frw1, "l")
item1.SubItems(12) = Sheets(19).Cells(frw1, "m")
item1.SubItems(13) = Sheets(19).Cells(frw1, "n")
item1.SubItems(14) = Sheets(19).Cells(frw1, "o")
item1.SubItems(15) = Sheets(19).Cells(frw1, "p")
item1.SubItems(16) = Sheets(19).Cells(frw1, "q")
item1.SubItems(17) = Sheets(19).Cells(frw1, "r")
item1.SubItems(18) = Sheets(19).Cells(frw1, "s")
item1.SubItems(19) = Sheets(19).Cells(frw1, "t")
item1.SubItems(20) = Sheets(19).Cells(frw1, "u")
item1.SubItems(21) = Sheets(19).Cells(frw1, "v")
'item1.SubItems(22) = Sheets(19).Cells(frw1, "w")


TextBox75.Text = Application.WorksheetFunction.CountIf(Sheets(19).Range("e2:e" & last1), "*" & Me.ComboBox2.Text & "*")







End If
Next frw1
End With







emad:

End Sub

Private Sub ListBox1_Click()
On Error Resume Next
Sheets(19).Activate
If TextBox2.Text = "" And TextBox3.Text = "" Then
Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
End If
For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) = True Then
        For j = 1 To 64
        Controls("TextBox" & j).Text = Cells(ListBox1.List(i, 1), j)
        Next j
        TextBox77.Text = TextBox22.Text
    End If
Next i
ComboBox1.Visible = False
'===================================
Dim last1, frw1 As Integer
last1 = Sheets(19).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
If Sheets(19).Cells(frw1, "b") = TextBox2.Text And Sheets(19).Cells(frw1, "c") = TextBox3.Text Then
Image1.Picture = LoadPicture(Sheets(19).Cells(frw1, 64))
Image2.Picture = LoadPicture(Sheets(19).Cells(frw1, 65))

'TextBox58Text = Sheets(19).Cells(frw1, "bg")

End If
Next frw1
If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If
'=========================
' Dim colwidth, widtharray, field As Integer

'Sheets(19).Range("c:c").EntireColumn.AutoFit
'For field = 1 To 3
'colwidth = Sheets(19).Cells(1, field).Width
'If widtharray = Empty Then
'widtharray = colwidth
'Else
'widtharray = widtharray & "," & colwidth
'End If
'Next
'ListBox1.ColumnWidths = widtharray
TextBox14.Text = TextBox81.Text
End Sub

Private Sub TextBox17_Change()

End Sub

Private Sub TextBox18_Change()

End Sub

Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ComboBox1.Visible = True
End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox65_AfterUpdate()
If TextBox72.Text <> "*" Then
Beep
 Beep
Application.Speech.Speak " The employee                 number            you          entered is not              found  "
End If
End Sub

Private Sub TextBox65_Change()
On Error GoTo emad
Sheets(19).Cells(2, "ag") = TextBox65.Text
TextBox72.Text = Sheets(19).Cells(2, "ah").Text
If TextBox72.Text <> "" Then

 Beep
Application.Speech.Speak " The employee              number               you entered           is     already          in   the           administrative            system"
End If



emad:
End Sub

Private Sub TextBox65_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox72.Text = ""                                              '
   TextBox65.Text = ""
  ListBox1.value = ""

End Sub

Private Sub TextBox72_Change()
On Error Resume Next
ListBox1.Clear

Sheets(19).Activate
    For i = 1 To 63
            Controls("TextBox" & i).Text = ""
    Next i

    If TextBox72 = "" Then Exit Sub
    Sheets(19).Activate

    ss = Sheets(19).Cells(Rows.count, 3).End(xlUp).row
     k = 0
     
     
For Each c In Range("c2:c" & ss)
    If c Like TextBox72.value & "*" Then
        ListBox1.AddItem
        ListBox1.List(k, 0) = Cells(c.row, 3).value
        ListBox1.List(k, 1) = c.row
        k = k + 1
    End If
    
Next c
End Sub

Private Sub TextBox72_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets(19).Activate
TextBox72.Text = ""
'TextBox50.SetFocus
For i = 2 To Sheets(19).Cells(Rows.count, 1).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 2, 0) = Cells(i, 3).value
  ListBox1.List(i - 2, 1) = i
      Next i
      
      
      
     
End Sub

Private Sub TextBox77_Change()

End Sub

Private Sub UserForm_Activate()
TextBox80.Text = Format(Sheets(19).Cells(2, 24), "yyyy/mm/dd")
TextBox81.Text = Format(Sheets(19).Cells(2, 25), "yyyy/mm/dd")
Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)


'================================================
On Error Resume Next

'=========================

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True

.ColumnHeaders.Add , , "«·—ﬁ„", 40
.ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
.ColumnHeaders.Add , , "«”„ «·„ÊŸ›", 180
.ColumnHeaders.Add , , "«·⁄‰Ê«‰ «·ÊŸÌ›Ì ", 120
.ColumnHeaders.Add , , "Õ«·… «·Œœ„… «·ÊŸÌ›ÌÂ", 120
.ColumnHeaders.Add , , "  «—ÌŒ  ⁄Ì‰Â «·√Ê·", 120
.ColumnHeaders.Add , , "—ﬁ„ «·ﬂ «» ", 70
.ColumnHeaders.Add , , " «—ÌŒ «·«‰ﬁÿ«⁄ Ê«·›’·", 100
.ColumnHeaders.Add , , " —ﬁ„ «·ﬂ «» ", 70
.ColumnHeaders.Add , , " «—ÌŒ «⁄«œ Â ··ÊŸÌ›Â", 100
.ColumnHeaders.Add , , " —ﬁ„ «·ﬂ «» ", 120
.ColumnHeaders.Add , , "  «—ÌŒ «·„»«‘—Â «·À«‰Ì… ", 120
.ColumnHeaders.Add , , "«· «—ÌŒ «·Œ«·Ì ", 0
.ColumnHeaders.Add , , "«· «—ÌŒ «·Œ«·Ì ", 120
.ColumnHeaders.Add , , " ⁄œœ ”‰Ê«  «·Œœ„Â ﬁ»· «·«‰ﬁÿ«⁄ ", 140
.ColumnHeaders.Add , , " ⁄œœ ”‰Ê«  «·«‰ﬁÿ«⁄ ", 120
.ColumnHeaders.Add , , " ⁄œœ ”‰Ê«  «·Œœ„Â »⁄œ «⁄«œ Â ·· ⁄Ì‰ ", 120
.ColumnHeaders.Add , , " ⁄œœ ”‰Ê«  «·Œœ„Â «·„Õ ”»Â ··›’· «·”Ì«”Ì ", 120
.ColumnHeaders.Add , , " ⁄œœ ”‰Ê«  «·Œœ„Â · «—ÌŒ „⁄Ì‰ ", 120
.ColumnHeaders.Add , , " «· Õ’Ì· «·œ—«”Ì «·ÃœÌœ  ", 140
.ColumnHeaders.Add , , " «· Õ’Ì· «·œ—«”Ì ⁄‰œ «· ⁄Ì‰", 120
.ColumnHeaders.Add , , " ⁄œœ ”‰Ê«  «·Œœ„Â «·ﬂ·Ì… ", 120
.ColumnHeaders.Add , , "„”«— «·’Ê— «·‘Œ’Ì…  ", 120





ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"

End With


With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = &HFFFFFF     ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(19).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(19).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(19).Cells(frw1, "b")
item1.SubItems(2) = Sheets(19).Cells(frw1, "c")
item1.SubItems(3) = Sheets(19).Cells(frw1, "d")
item1.SubItems(4) = Sheets(19).Cells(frw1, "e")
item1.SubItems(5) = Sheets(19).Cells(frw1, "f")
item1.SubItems(6) = Sheets(19).Cells(frw1, "g")
item1.SubItems(7) = Sheets(19).Cells(frw1, "h")
item1.SubItems(8) = Sheets(19).Cells(frw1, "i")
item1.SubItems(9) = Sheets(19).Cells(frw1, "j")
item1.SubItems(10) = Sheets(19).Cells(frw1, "k")
item1.SubItems(11) = Sheets(19).Cells(frw1, "l")
item1.SubItems(12) = Sheets(19).Cells(frw1, "m")
item1.SubItems(13) = Sheets(19).Cells(frw1, "n")
item1.SubItems(14) = Sheets(19).Cells(frw1, "o")
item1.SubItems(15) = Sheets(19).Cells(frw1, "p")
item1.SubItems(16) = Sheets(19).Cells(frw1, "q")
item1.SubItems(17) = Sheets(19).Cells(frw1, "r")
item1.SubItems(18) = Sheets(19).Cells(frw1, "s")
item1.SubItems(19) = Sheets(19).Cells(frw1, "t")
item1.SubItems(20) = Sheets(19).Cells(frw1, "u")
item1.SubItems(21) = Sheets(19).Cells(frw1, "v")
item1.SubItems(22) = Sheets(19).Cells(frw1, "w")
item1.SubItems(23) = Sheets(19).Cells(frw1, "x")
item1.SubItems(24) = Sheets(19).Cells(frw1, "y")






Next frw1


'=============================================
   
   
   
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
ComboBox1.List = Sheets(19).Range("AI2:AI7")
'emad:
'Else
'CommandButton5.Enabled = False
'End If
End Sub
Private Sub UserForm_Initialize()
'On Error GoTo emad
TextBox80.Text = Format(Ê—ﬁ…19.Cells(2, "x"), "dd/mm/yyyy")
ComboBox1.List = Sheets(19).Range("AI2:AI7").value
ComboBox2.List = Sheets(19).Range("AI2:AI7").value

emad:
End Sub
