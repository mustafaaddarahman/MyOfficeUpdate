VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm47 
   Caption         =   "UserForm47"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   22140
   OleObjectBlob   =   "UserForm47.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm47"
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

Private Sub ComboBox1_Change()
TextBox23.value = ComboBox1.value
ComboBox1.Visible = False
End Sub

Private Sub ComboBox1_DropButtonClick()

TextBox23.value = ComboBox1.value

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub ComboBox3_Change()
TextBox58.Text = ComboBox3.value
End Sub

Private Sub CommandButton1_Click()
Unload Me







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

Private Sub CommandButton15_Click()


 
End Sub

Private Sub CommandButton16_Click()

End Sub

Private Sub CommandButton17_Click()
On Error GoTo emad
Application.ThisWorkbook.Save
ThisWorkbook.SaveCopyAs fileName:="d:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
ThisWorkbook.SaveCopyAs fileName:="e:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
MsgBox "·ﬁœ  „  ⁄„·Ì… «·‰”Œ «·«Œ Ì«ÿÌ »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
emad:
End Sub





Private Sub CommandButton20_Click()
Dim lastR As Integer
lastR = Sheets(51).Cells(Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(51).Cells(Y, 5) = TextBox5.Text And Sheets(51).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(51).Cells(Y, 1) = TextBox1.Text
    Sheets(51).Cells(Y, 2) = TextBox2.Text
    Sheets(51).Cells(Y, 3) = TextBox3.Text
    Sheets(51).Cells(Y, 4) = TextBox4.value
    Sheets(51).Cells(Y, 5) = TextBox5.Text
    Sheets(51).Cells(Y, 6) = TextBox6.value
    Sheets(51).Cells(Y, 7) = TextBox7.Text
    Sheets(51).Cells(Y, 8) = TextBox8.Text
    Sheets(51).Cells(Y, 9) = TextBox9.Text
    Sheets(51).Cells(Y, 10) = TextBox10.Text
    Sheets(51).Cells(Y, 11) = TextBox11.Text
    Sheets(51).Cells(Y, 12) = TextBox12.Text
     Sheets(51).Cells(Y, 13) = TextBox13.value
    Sheets(51).Cells(Y, 14) = TextBox14.Text
    Sheets(51).Cells(Y, 15) = TextBox15.value
    Sheets(51).Cells(Y, 16) = TextBox16.Text
   Sheets(51).Cells(Y, 17) = TextBox17.value
    Sheets(51).Cells(Y, 18) = TextBox18.Text
    Sheets(51).Cells(Y, 19) = TextBox19.Text
    Sheets(51).Cells(Y, 20) = TextBox20.Text
    Sheets(51).Cells(Y, 21) = TextBox21.Text
    Sheets(51).Cells(Y, 22) = TextBox22.Text
    Sheets(51).Cells(Y, 23) = TextBox23.Text
    Sheets(51).Cells(Y, 24) = TextBox24.Text
    Sheets(51).Cells(Y, 25) = TextBox25.Text
   Sheets(51).Cells(Y, 26) = TextBox26.value
    Sheets(51).Cells(Y, 27) = TextBox27.Text
   Sheets(51).Cells(Y, 28) = TextBox28.value
  Sheets(51).Cells(Y, 29) = TextBox29.Text
    Sheets(51).Cells(Y, 30) = TextBox30.Text
    Sheets(51).Cells(Y, 31) = TextBox31.Text
  Sheets(51).Cells(Y, 32) = TextBox32.Text
   Sheets(51).Cells(Y, 33) = TextBox33.Text
    Sheets(51).Cells(Y, 34) = TextBox34.Text
    Sheets(51).Cells(Y, 35) = TextBox35.Text
    Sheets(51).Cells(Y, 36) = TextBox36.Text
    Sheets(51).Cells(Y, 37) = TextBox37.Text
    Sheets(51).Cells(Y, 38) = TextBox38.Text
    Sheets(51).Cells(Y, 39) = TextBox39.Text
    Sheets(51).Cells(Y, 40) = TextBox40.Text
    Sheets(51).Cells(Y, 41) = TextBox41.Text
    Sheets(51).Cells(Y, 42) = TextBox42.Text
    Sheets(51).Cells(Y, 43) = TextBox43.Text
    Sheets(51).Cells(Y, 44) = TextBox44.Text
    Sheets(51).Cells(Y, 45) = TextBox45.Text
    Sheets(51).Cells(Y, 46) = TextBox46.Text
    Sheets(51).Cells(Y, 47) = TextBox47.Text
  Sheets(51).Cells(Y, 48) = TextBox48.Text
   Sheets(51).Cells(Y, 49) = TextBox49.Text
   Sheets(51).Cells(Y, 50) = TextBox50.Text
   Sheets(51).Cells(Y, 51) = TextBox51.value
    Sheets(51).Cells(Y, 52) = TextBox52.Text
  Sheets(51).Cells(Y, 53) = TextBox53.Text
   Sheets(51).Cells(Y, 54) = TextBox54.Text
    Sheets(51).Cells(Y, 55) = TextBox55.Text
    Sheets(51).Cells(Y, 56) = TextBox56.Text
    Sheets(51).Cells(Y, 57) = TextBox57.Text
  Sheets(51).Cells(Y, 58) = TextBox58.Text
   Sheets(51).Cells(Y, 59) = TextBox59.Text
    Sheets(51).Cells(Y, 60) = TextBox60.Text
    Sheets(51).Cells(Y, 61) = TextBox61.Text
  Sheets(51).Cells(Y, 62) = TextBox62.Text
   Sheets(51).Cells(Y, 63) = TextBox63.Text
  Sheets(51).Cells(Y, 64) = TextBox64.Text
    
    
'    Sheets(3).Cells(y, 66) = Format(TextBox82.Text, "yyyy/mm/dd")
'  Sheets(3).Cells(y, 67) = TextBox83.Text
   
'   Sheets(3).Cells(y, 68) = TextBox84.Text
' Sheets(3).Cells(y, 69) = TextBox85.Text
'  Sheets(3).Cells(y, 70) = TextBox86.Text
'   Sheets(3).Cells(y, 71) = TextBox87.Text
'   Sheets(3).Cells(y, 72) = TextBox88.Text
   Sheets(51).Cells(Y, 73) = TextBox64.Text
   Sheets(51).Cells(Y, 75) = TextBox81.Text
    
          
     MsgBox " „  ⁄„·Ì…  ÕœÌÀ «·»Ì«‰«  »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
emad:
End Sub

Private Sub CommandButton21_Click()


End Sub

Private Sub CommandButton22_Click()




End Sub

Private Sub CommandButton23_Click()
On Error Resume Next
Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton24_Click()
Dim wb As Workbook, lr1 As Double, lr2 As Double
Dim fil As Variant, dat As Double

Application.ScreenUpdating = False
Application.DisplayAlerts = False


lr1 = Sheets(51).Cells(Rows.count, 2).End(xlUp).row
Sheets(51).Range("A9:bv" & lr1 + 1).ClearContents

INF = ThisWorkbook.path
fil = Dir(INF & "\*.xl??")


Do While fil <> ""
    If fil <> "Financial and administrative system.xlsm" Then
       
        Set wb = Workbooks.Open(INF & "\" & fil)
        lr1 = Workbooks(1).Sheets(51).Cells(Rows.count, 2).End(xlUp).row
        lr2 = ActiveSheet.Cells(Rows.count, 1).End(xlUp).row
        ActiveSheet.Range("A9:bv" & lr2).Copy Workbooks(1).Sheets(51).Range("A" & lr1 + 1)
       ' dep = Left(ActiveWorkbook.Name, Application.Search(".", ActiveWorkbook.Name) - 1)
       ' Workbooks("DATA").Sheets("Temp").Range("g" & lr1 + 1 & ":g" & lr1 + lr2 - 1) = dep
       Application.Visible = True

        ActiveWorkbook.Close True
        Application.Visible = False
    End If
fil = Dir
Loop


Application.DisplayAlerts = True
Application.ScreenUpdating = True

Sheets(51).Activate
MsgBox "·ﬁœ  „ ‰ﬁ· Ê«” Ì—«œ «·»Ì«‰«  «·Œ«—ÃÌÂ »‰Ã«Õ", vbCritical, "—”«·…  ‰»ÌÂ"






End Sub



Private Sub CommandButton25_Click()

End Sub

Private Sub CommandButton3_Click()
On Error GoTo emad

End Sub

Private Sub CommandButton4_Click()
'========================================================
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
      If TextBox56.Text = "" Or TextBox57.Text = "" Or TextBox58.Text = "" Then
    MsgBox "„‰ ›÷·ﬂ «œŒ· «·ÃÂ… «·„‰”Ê» «·ÌÂ« Ê «—ÌŒ  —ÕÌ·Â Ê„Êﬁ›Â «·ÊŸÌ›Ì", vbCritical, "—”«·…  ‰»ÌÂ"
     
        Exit Sub
       Else
    
    Dim iRow As Long, i As Long

    Sheets(1).Activate

    iRow = Range("b" & Rows.count).End(xlUp).row
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
TextBox53.Text = Format(Sheets(3).Cells(1, "cc"), "yyyy/mm,dd")
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
    Range("A" & iRow + 1).Offset(0, 13).value = TextBox14.value
    Range("A" & iRow + 1).Offset(0, 14).value = TextBox15.value
    Range("A" & iRow + 1).Offset(0, 15).value = TextBox16.value
    Range("A" & iRow + 1).Offset(0, 16).value = TextBox17.value
    Range("A" & iRow + 1).Offset(0, 17).value = TextBox18.value
    Range("A" & iRow + 1).Offset(0, 18).value = TextBox19.value
    Range("A" & iRow + 1).Offset(0, 19).value = TextBox20.value
    Range("A" & iRow + 1).Offset(0, 20).value = TextBox21.value
    Range("A" & iRow + 1).Offset(0, 21).value = TextBox22.value
    Range("A" & iRow + 1).Offset(0, 22).value = TextBox23.value
    Range("A" & iRow + 1).Offset(0, 23).value = TextBox24.value
    Range("A" & iRow + 1).Offset(0, 24).value = TextBox25.value
    Range("A" & iRow + 1).Offset(0, 25).value = TextBox26.value
    Range("A" & iRow + 1).Offset(0, 26).value = TextBox27.value
    Range("A" & iRow + 1).Offset(0, 27).value = TextBox28.value
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
     Range("A" & iRow + 1).Offset(0, 51).value = TextBox52.value
     Range("A" & iRow + 1).Offset(0, 52).value = Format(TextBox53.Text, "yyyy/mm/dd")
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
    Range("A" & iRow + 1).Offset(0, 63).value = TextBox64.value
    
      
      
      
      
      
      
        Sheets(1).Activate
'============================================================================


Sheets(51).Activate
Dim lastR As Integer
lastR = Sheets(51).Cells(Rows.count, 1).End(xlUp).row
For X = 2 To lastR
If Sheets(51).Cells(X, 2).value = TextBox2.Text And Sheets(51).Cells(X, 5).value = TextBox5.Text Then
 Sheets(51).Cells(X, 1).Select

End If
Next


End If
Sheets(51).Activate

MsgBox " „  —ÕÌ·Â «·Ï «·‰Ÿ«„ «·—∆Ì”Ì ÊÕ–›Â „‰ »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"

End Sub



Private Sub CommandButton6_Click()


End Sub



Private Sub CommandButton8_Click()
Sheets(51).Activate
For X = 9 To 1500
If Cells(X, 5) = TextBox5.Text Then
Cells(X, 5).Select
Exit For
End If
Next X
'==================================



End Sub

Private Sub CommandButton7_Click()

End Sub

Private Sub CommandButton9_Click()
On Error GoTo emad
 With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer


Sheets(51).Activate
last1 = Sheets(51).Cells(Rows.count, 1).End(xlUp).row



'=============================================

   Sheets(51).Activate
    
    
    For frw1 = 2 To last1
                
  If Sheets(51).Cells(frw1, "h") = TextBox67.Text And Sheets(51).Cells(frw1, "f") = ComboBox2.value And ComboBox3.Text = Sheets(51).Cells(frw1, "bf").value Then
     
'==================================================

Set item1 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(51).Cells(frw1, "b")
item1.SubItems(2) = Sheets(51).Cells(frw1, "c")
item1.SubItems(3) = Sheets(51).Cells(frw1, "d")
item1.SubItems(4) = Sheets(51).Cells(frw1, "e")
item1.SubItems(5) = Sheets(51).Cells(frw1, "f")
item1.SubItems(6) = Sheets(51).Cells(frw1, "g")
item1.SubItems(7) = Sheets(51).Cells(frw1, "h")
item1.SubItems(8) = Sheets(51).Cells(frw1, "i")
item1.SubItems(9) = Sheets(51).Cells(frw1, "j")
item1.SubItems(10) = Sheets(51).Cells(frw1, "k")
item1.SubItems(11) = Sheets(51).Cells(frw1, "l")
item1.SubItems(12) = Sheets(51).Cells(frw1, "m")
item1.SubItems(13) = Sheets(51).Cells(frw1, "n")
item1.SubItems(14) = Sheets(51).Cells(frw1, "o")
item1.SubItems(15) = Sheets(51).Cells(frw1, "p")
item1.SubItems(16) = Sheets(51).Cells(frw1, "q")
item1.SubItems(17) = Sheets(51).Cells(frw1, "r")
item1.SubItems(18) = Sheets(51).Cells(frw1, "s")
item1.SubItems(19) = Sheets(51).Cells(frw1, "t")
item1.SubItems(20) = Sheets(51).Cells(frw1, "u")
item1.SubItems(21) = Sheets(51).Cells(frw1, "v")
item1.SubItems(22) = Sheets(51).Cells(frw1, "w")
item1.SubItems(23) = Sheets(51).Cells(frw1, "x")
item1.SubItems(24) = Sheets(51).Cells(frw1, "y")
item1.SubItems(25) = Sheets(51).Cells(frw1, "z")
item1.SubItems(26) = Sheets(51).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(51).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(51).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(51).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(51).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(51).Cells(frw1, "af")
item1.SubItems(32) = Sheets(51).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(51).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(51).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(51).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(51).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(51).Cells(frw1, "al")
item1.SubItems(38) = Sheets(51).Cells(frw1, "am")
item1.SubItems(39) = Sheets(51).Cells(frw1, "an")
item1.SubItems(40) = Sheets(51).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(51).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(51).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(51).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(51).Cells(frw1, "as")
item1.SubItems(45) = Sheets(51).Cells(frw1, "at")
item1.SubItems(46) = Sheets(51).Cells(frw1, "au")
item1.SubItems(47) = Sheets(51).Cells(frw1, "av")
item1.SubItems(48) = Sheets(51).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(51).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(51).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(51).Cells(frw1, "az")
item1.SubItems(52) = Sheets(51).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(51).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(51).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(51).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(51).Cells(frw1, "be")
item1.SubItems(57) = Sheets(51).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(51).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(51).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(51).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(51).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(51).Cells(frw1, "bk")
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("j9:j800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ad9:ad800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ae9:ae800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("av9:av800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("aw9:aw800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("x9:x800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(51).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")






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
last2 = Sheets(51).Range("A" & Rows.count).End(xlUp).row
For frw2 = 2 To last2



If ComboBox2.value = "«·ﬂ·" And Sheets(51).Cells(frw2, "bf") = ComboBox3.value Then

Set item2 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw2, "A"))




item2.SubItems(1) = Sheets(51).Cells(frw2, "b")
item2.SubItems(2) = Sheets(51).Cells(frw2, "c")
item2.SubItems(3) = Sheets(51).Cells(frw2, "d")
item2.SubItems(4) = Sheets(51).Cells(frw2, "e")
item2.SubItems(5) = Sheets(51).Cells(frw2, "f")
item2.SubItems(6) = Sheets(51).Cells(frw2, "g")
item2.SubItems(7) = Sheets(51).Cells(frw2, "h")
item2.SubItems(8) = Sheets(51).Cells(frw2, "i")
item2.SubItems(9) = Sheets(51).Cells(frw2, "j")
item2.SubItems(10) = Sheets(51).Cells(frw2, "k")
item2.SubItems(11) = Sheets(51).Cells(frw2, "l")
item2.SubItems(12) = Sheets(51).Cells(frw2, "m")
item2.SubItems(13) = Sheets(51).Cells(frw2, "n")
item2.SubItems(14) = Sheets(51).Cells(frw2, "o")
item2.SubItems(15) = Sheets(51).Cells(frw2, "p")
item2.SubItems(16) = Sheets(51).Cells(frw2, "q")
item2.SubItems(17) = Sheets(51).Cells(frw2, "r")
item2.SubItems(18) = Sheets(51).Cells(frw2, "s")
item2.SubItems(19) = Sheets(51).Cells(frw2, "t")
item2.SubItems(20) = Sheets(51).Cells(frw2, "u")
item2.SubItems(21) = Sheets(51).Cells(frw2, "v")
item2.SubItems(22) = Sheets(51).Cells(frw2, "w")
item2.SubItems(23) = Sheets(51).Cells(frw2, "x")
item2.SubItems(24) = Sheets(51).Cells(frw2, "y")
item2.SubItems(25) = Sheets(51).Cells(frw2, "z")
item2.SubItems(26) = Sheets(51).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(51).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(51).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(51).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(51).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(51).Cells(frw2, "af")
item2.SubItems(32) = Sheets(51).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(51).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(51).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(51).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(51).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(51).Cells(frw2, "al")
item2.SubItems(38) = Sheets(51).Cells(frw2, "am")
item2.SubItems(39) = Sheets(51).Cells(frw2, "an")
item2.SubItems(40) = Sheets(51).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(51).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(51).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(51).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(51).Cells(frw2, "as")
item2.SubItems(45) = Sheets(51).Cells(frw2, "at")
item2.SubItems(46) = Sheets(51).Cells(frw2, "au")
item2.SubItems(47) = Sheets(51).Cells(frw2, "av")
item2.SubItems(48) = Sheets(51).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(51).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(51).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(51).Cells(frw2, "az")
item2.SubItems(52) = Sheets(51).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(51).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(51).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(51).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(51).Cells(frw2, "be")
item2.SubItems(57) = Sheets(51).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(51).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(51).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(51).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(51).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(51).Cells(frw2, "bk")
'item2.SubItems(63) = Sheets(51).Cells(frw2, "bl")
simadd = Application.WorksheetFunction.Sum(Sheets(51).Range("bh9:bh" & frw2))
simsubtract = Application.WorksheetFunction.Sum(Sheets(51).Range("bj9:bj" & frw2))
TextBox73.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("j9:j" & frw2))
TextBox74.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("ad9:ad" & frw2))
TextBox79.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("ae9:ae" & frw2))
TextBox80.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("av9:av" & frw2))
TextBox66.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("aw9:aw" & frw2))
TextBox76.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("X9:X" & frw1))
TextBox67.Text = ""

TextBox66.Text = (TextBox66.Text + simadd) - simsubtract
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(51).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578;BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")




ElseIf Sheets(51).Cells(frw1, "h") Like "*" & TextBox67.Text & "*" And Sheets(51).Cells(frw1, "f") Like "*" & ComboBox2.Text & "*" And ComboBox3.value = Sheets(51).Cells(frw1, "bf").value Then
  item2.SubItems(1) = Sheets(51).Cells(frw2, "b")
item2.SubItems(2) = Sheets(51).Cells(frw2, "c")
item2.SubItems(3) = Sheets(51).Cells(frw2, "d")
item2.SubItems(4) = Sheets(51).Cells(frw2, "e")
item2.SubItems(5) = Sheets(51).Cells(frw2, "f")
item2.SubItems(6) = Sheets(51).Cells(frw2, "g")
item2.SubItems(7) = Sheets(51).Cells(frw2, "h")
item2.SubItems(8) = Sheets(51).Cells(frw2, "i")
item2.SubItems(9) = Sheets(51).Cells(frw2, "j")
item2.SubItems(10) = Sheets(51).Cells(frw2, "k")
item2.SubItems(11) = Sheets(51).Cells(frw2, "l")
item2.SubItems(12) = Sheets(51).Cells(frw2, "m")
item2.SubItems(13) = Sheets(51).Cells(frw2, "n")
item2.SubItems(14) = Sheets(51).Cells(frw2, "o")
item2.SubItems(15) = Sheets(51).Cells(frw2, "p")
item2.SubItems(16) = Sheets(51).Cells(frw2, "q")
item2.SubItems(17) = Sheets(51).Cells(frw2, "r")
item2.SubItems(18) = Sheets(51).Cells(frw2, "s")
item2.SubItems(19) = Sheets(51).Cells(frw2, "t")
item2.SubItems(20) = Sheets(51).Cells(frw2, "u")
item2.SubItems(21) = Sheets(51).Cells(frw2, "v")
item2.SubItems(22) = Sheets(51).Cells(frw2, "w")
item2.SubItems(23) = Sheets(51).Cells(frw2, "x")
item2.SubItems(24) = Sheets(51).Cells(frw2, "y")
item2.SubItems(25) = Sheets(51).Cells(frw2, "z")
item2.SubItems(26) = Sheets(51).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(51).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(51).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(51).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(51).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(51).Cells(frw2, "af")
item2.SubItems(32) = Sheets(51).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(51).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(51).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(51).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(51).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(51).Cells(frw2, "al")
item2.SubItems(38) = Sheets(51).Cells(frw2, "am")
item2.SubItems(39) = Sheets(51).Cells(frw2, "an")
item2.SubItems(40) = Sheets(51).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(51).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(51).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(51).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(51).Cells(frw2, "as")
item2.SubItems(45) = Sheets(51).Cells(frw2, "at")
item2.SubItems(46) = Sheets(51).Cells(frw2, "au")
item2.SubItems(47) = Sheets(51).Cells(frw2, "av")
item2.SubItems(48) = Sheets(51).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(51).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(51).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(51).Cells(frw2, "az")
item2.SubItems(52) = Sheets(51).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(51).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(51).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(51).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(51).Cells(frw2, "be")
item2.SubItems(57) = Sheets(51).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(51).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(51).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(51).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(51).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(51).Cells(frw2, "bk")
item2.SubItems(63) = Sheets(51).Cells(frw2, "bl")





'==================================================
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("j9:j800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ad9:ad800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ae9:ae800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("av9:av800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("aw9:aw800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("x9:x800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(51).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")






UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")
End If
Next






'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(51).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("BF9:BF571"), "=" & Me.ComboBox3.Value)


'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("BF9:BF" & frw2), "=" & Me.ComboBox3.Text)
TextBox75.Text = ListView1.ListItems.count




emad:



End Sub



Private Sub Frame_Click()

End Sub

Private Sub ListBox1_Click()
'On Error Resume Next
Sheets(51).Activate
If TextBox2.Text = "" And TextBox5.Text = "" Then
Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
End If
For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) = True Then
        For j = 1 To 64
        Controls("TextBox" & j).Text = Cells(ListBox1.List(i, 1), j)
        Next j
    End If
Next i
ComboBox1.Visible = False
'===================================
Dim last1, frw1 As Integer
last1 = Sheets(51).Range("A" & Rows.count).End(xlUp).row
For frw1 = 1 To last1

If Sheets(51).Cells(frw1, "b") = TextBox2.Text And Sheets(51).Cells(frw1, "e") = TextBox5.Text Then
TextBox81.Text = Sheets(51).Cells(frw1, "bw")
Image1.Picture = LoadPicture(Sheets(51).Cells(frw1, 64))

'TextBox58Text = Sheets(51).Cells(frw1, "bg")

End If
Next frw1
If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If

TextBox53.value = Format(Date, "yyyy/mm/dd")
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ComboBox1.Visible = False
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

 
End Sub



Private Sub TextBox10_Change()


End Sub

Private Sub TextBox11_Change()


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

Private Sub TextBox21_Change()

End Sub

Private Sub TextBox22_Change()

End Sub

Private Sub TextBox23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next

ComboBox1.Visible = True
ComboBox1.List = Sheets(4).Range("f1:f11").value

End Sub

Private Sub TextBox25_Change()

End Sub

Private Sub TextBox26_Change()

End Sub

Private Sub TextBox28_Change()

End Sub

Private Sub TextBox3_Change()



End Sub

Private Sub TextBox34_Change()

End Sub

Private Sub TextBox35_Change()

End Sub

Private Sub TextBox36_Change()

End Sub

Private Sub TextBox37_Change()

End Sub

Private Sub TextBox38_Change()

End Sub

Private Sub TextBox4_Change()
arb



End Sub

Private Sub TextBox40_Change()

End Sub

Private Sub TextBox41_Change()

End Sub

Private Sub TextBox42_Change()

End Sub

Private Sub TextBox43_Change()

End Sub

Private Sub TextBox44_Change()

End Sub

Private Sub TextBox45_Change()

End Sub

Private Sub TextBox46_Change()

End Sub

Private Sub TextBox5_Change()
arb
End Sub

Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


TextBox27.value = ""
ListBox1.Clear
End Sub



Private Sub TextBox6_Change()
arb

End Sub

Private Sub TextBox60_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets(51).Activate
CommandButton5.Locked = True
TextBox50.SetFocus
For i = 9 To Sheets(51).Cells(Rows.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 9, 0) = Cells(i, 5).value
  ListBox1.List(i - 9, 1) = i
      Next i
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
Sheets(51).Cells(3, 64) = TextBox65.Text
TextBox72.Text = Sheets(51).Cells(3, 65).Text
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

Private Sub TextBox67_Change()

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
last1 = Sheets(51).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(51).Cells(frw1, "b")
item1.SubItems(2) = Sheets(51).Cells(frw1, "c")
item1.SubItems(3) = Sheets(51).Cells(frw1, "d")
item1.SubItems(4) = Sheets(51).Cells(frw1, "e")
item1.SubItems(5) = Sheets(51).Cells(frw1, "f")
item1.SubItems(6) = Sheets(51).Cells(frw1, "g")
item1.SubItems(7) = Sheets(51).Cells(frw1, "h")
item1.SubItems(8) = Sheets(51).Cells(frw1, "i")
item1.SubItems(9) = Sheets(51).Cells(frw1, "j")
item1.SubItems(10) = Sheets(51).Cells(frw1, "k")
item1.SubItems(11) = Sheets(51).Cells(frw1, "l")
item1.SubItems(12) = Sheets(51).Cells(frw1, "m")
item1.SubItems(13) = Sheets(51).Cells(frw1, "n")
item1.SubItems(14) = Sheets(51).Cells(frw1, "o")
item1.SubItems(15) = Sheets(51).Cells(frw1, "p")
item1.SubItems(16) = Sheets(51).Cells(frw1, "q")
item1.SubItems(17) = Sheets(51).Cells(frw1, "r")
item1.SubItems(18) = Sheets(51).Cells(frw1, "s")
item1.SubItems(19) = Sheets(51).Cells(frw1, "t")
item1.SubItems(20) = Sheets(51).Cells(frw1, "u")
item1.SubItems(21) = Sheets(51).Cells(frw1, "v")
item1.SubItems(22) = Sheets(51).Cells(frw1, "w")
item1.SubItems(23) = Sheets(51).Cells(frw1, "x")
item1.SubItems(24) = Sheets(51).Cells(frw1, "y")
item1.SubItems(25) = Sheets(51).Cells(frw1, "z")
item1.SubItems(26) = Sheets(51).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(51).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(51).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(51).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(51).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(51).Cells(frw1, "af")
item1.SubItems(32) = Sheets(51).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(51).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(51).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(51).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(51).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(51).Cells(frw1, "al")
item1.SubItems(38) = Sheets(51).Cells(frw1, "am")
item1.SubItems(39) = Sheets(51).Cells(frw1, "an")
item1.SubItems(40) = Sheets(51).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(51).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(51).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(51).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(51).Cells(frw1, "as")
item1.SubItems(45) = Sheets(51).Cells(frw1, "at")
item1.SubItems(46) = Sheets(51).Cells(frw1, "au")
item1.SubItems(47) = Sheets(51).Cells(frw1, "av")
item1.SubItems(48) = Sheets(51).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(51).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(51).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(51).Cells(frw1, "az")
item1.SubItems(52) = Sheets(51).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(51).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(51).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(51).Cells(frw1, "bf")




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
arb
ListBox1.Clear

Sheets(51).Activate
    For i = 1 To 63
            Controls("TextBox" & i).Text = ""
    Next i

    If TextBox72 = "" Then Exit Sub
    Sheets(51).Activate

    ss = Sheets(51).Cells(Rows.count, 5).End(xlUp).row
     k = 0
     
     
For Each c In Range("E2:E" & ss)
    If c Like TextBox72.value & "*" Then
        ListBox1.AddItem
        ListBox1.List(k, 0) = Cells(c.row, 5).value
        ListBox1.List(k, 1) = c.row
        k = k + 1
    End If
    
Next c
End Sub

Private Sub TextBox72_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets(51).Activate
TextBox72.Text = ""
'TextBox50.SetFocus
For i = 2 To Sheets(51).Cells(Rows.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 2, 0) = Cells(i, 5).value
  ListBox1.List(i - 2, 1) = i
      Next i
End Sub

Private Sub TextBox8_Change()


End Sub

Private Sub TextBox81_Change()

End Sub

Private Sub TextBox9_Change()


End Sub
Private Sub UserForm_Activate()

On Error GoTo emad

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

ComboBox2.List = Sheets(5).Range("a1:a373").value
ComboBox3.List = Sheets(4).Range("g1:g13").value
Me.TextBox67.Text = ""
'========================================================== Õ„Ì· »‰ﬂ «·»Ì«‰« 





















emad:

End Sub





Private Sub UserForm_Initialize()
'===============================
On Error Resume Next

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
ColumnHeaders.Add , , " «—ÌŒ Õ’Ê·Â ⁄·Ï ﬂ «» ‘ﬂ— Ê ﬁœÌ— ", 140
.ColumnHeaders.Add , , "  «—ÌŒ Õ–›Â Ê —ÕÌ·Â ", 0
.ColumnHeaders.Add , , " «·≈÷«›… «·„«·ÌÂ «·„ƒﬁ Â", 140
.ColumnHeaders.Add , , " «·„·«ÕŸ«  ", 140
.ColumnHeaders.Add , , " «·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ  ", 140
.ColumnHeaders.Add , , " —„“ «·«” Õﬁ«ﬁ ", 140




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

last1 = Sheets(51).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(51).Cells(frw1, "b")
item1.SubItems(2) = Sheets(51).Cells(frw1, "c")
item1.SubItems(3) = Sheets(51).Cells(frw1, "d")
item1.SubItems(4) = Sheets(51).Cells(frw1, "e")
item1.SubItems(5) = Sheets(51).Cells(frw1, "f")
item1.SubItems(6) = Sheets(51).Cells(frw1, "g")
item1.SubItems(7) = Sheets(51).Cells(frw1, "h")
item1.SubItems(8) = Sheets(51).Cells(frw1, "i")
item1.SubItems(9) = Sheets(51).Cells(frw1, "j")
item1.SubItems(10) = Sheets(51).Cells(frw1, "k")
item1.SubItems(11) = Sheets(51).Cells(frw1, "l")
item1.SubItems(12) = Sheets(51).Cells(frw1, "m")
item1.SubItems(13) = Sheets(51).Cells(frw1, "n")
item1.SubItems(14) = Sheets(51).Cells(frw1, "o")
item1.SubItems(15) = Sheets(51).Cells(frw1, "p")
item1.SubItems(16) = Sheets(51).Cells(frw1, "q")
item1.SubItems(17) = Sheets(51).Cells(frw1, "r")
item1.SubItems(18) = Sheets(51).Cells(frw1, "s")
item1.SubItems(19) = Sheets(51).Cells(frw1, "t")
item1.SubItems(20) = Sheets(51).Cells(frw1, "u")
item1.SubItems(21) = Sheets(51).Cells(frw1, "v")
item1.SubItems(22) = Sheets(51).Cells(frw1, "w")
item1.SubItems(23) = Sheets(51).Cells(frw1, "x")
item1.SubItems(24) = Sheets(51).Cells(frw1, "y")
item1.SubItems(25) = Sheets(51).Cells(frw1, "z")
item1.SubItems(26) = Sheets(51).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(51).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(51).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(51).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(51).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(51).Cells(frw1, "af")
item1.SubItems(32) = Sheets(51).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(51).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(51).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(51).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(51).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(51).Cells(frw1, "al")
item1.SubItems(38) = Sheets(51).Cells(frw1, "am")
item1.SubItems(39) = Sheets(51).Cells(frw1, "an")
item1.SubItems(40) = Sheets(51).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(51).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(51).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(51).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(51).Cells(frw1, "as")
item1.SubItems(45) = Sheets(51).Cells(frw1, "at")
item1.SubItems(46) = Sheets(51).Cells(frw1, "au")
item1.SubItems(47) = Sheets(51).Cells(frw1, "av")
item1.SubItems(48) = Sheets(51).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(51).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(51).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(51).Cells(frw1, "az")
item1.SubItems(52) = Sheets(51).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(51).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(51).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(51).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(51).Cells(frw1, "be")
item1.SubItems(57) = Sheets(51).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(51).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(51).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(51).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(51).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(51).Cells(frw1, "bk")
item1.SubItems(74) = Sheets(51).Cells(frw1, "bw")




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
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
 If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If

End Sub



