VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm33 
   Caption         =   "«·‰Ÿ«„ «·„«·Ì"
   ClientHeight    =   11880
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   20748
   OleObjectBlob   =   "UserForm33.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm33"
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

Private Sub ComboBox10_Change()
If TextBox10.value = "" Then
TextBox10.value = 0

Else
TextBox27.value = TextBox10.value * ComboBox10.value
End If
End Sub

Private Sub ComboBox11_Change()
If TextBox10.value = "" Then
TextBox10.value = 0
Else
TextBox29.value = TextBox10.value * ComboBox11.value
End If
End Sub

Private Sub ComboBox2_Change()
ListView1.ListItems.Clear
End Sub

Private Sub ComboBox3_AfterUpdate()
On Error Resume Next
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
End Sub

Private Sub ComboBox3_Change()
On Error Resume Next
TextBox58.Text = ComboBox3.value
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
End Sub

Private Sub ComboBox6_Change()
On Error GoTo emad
TextBox14.value = TextBox10.value * ComboBox6.value
emad:
End Sub

Private Sub ComboBox7_Change()
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


Call bdoon_ratb1

'======================================== «·«” Õﬁ«ﬁ« 





'=========================================== «·«” ﬁÿ«⁄« 

If TextBox58.Text = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Then



Call five_snwat
'Call update_data_acount
End If



'==================================================«·’«›Ì





If TextBox58.Text = "„” „— »«·Œœ„…" Then

   Call mastmer_5dma
 'Call update_data_acount
End If
   '=========================================
Call update_data_acount
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

Private Sub CommandButton2_Click()
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
    Range("A" & iRow + 1).Offset(0, 25).value = TextBox26.value
   ' Range("A" & iRow + 1).Offset(0, 26).Value = TextBox27.Value
   Range("A" & iRow + 1).Offset(0, 27).value = TextBox28.value
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
    Range("A" & iRow + 1).Offset(0, 50).value = TextBox51.value
   ' Range("A" & iRow + 1).Offset(0, 51).Value = TextBox52.Value
   ' Range("A" & iRow + 1).Offset(0, 52).Value = TextBox53.Value
   ' Range("A" & iRow + 1).Offset(0, 53).Value = TextBox54.Value
    Range("A" & iRow + 1).Offset(0, 54).value = TextBox55.value
    Range("A" & iRow + 1).Offset(0, 55).value = TextBox56.value
    Range("A" & iRow + 1).Offset(0, 56).value = ComboBox9.value
   ' Range("A" & iRow + 1).Offset(0, 57).Value = TextBox58.Value
   ' Range("A" & iRow + 1).Offset(0, 58).Value = TextBox59.Value
    Range("A" & iRow + 1).Offset(0, 59).value = TextBox60.value
   ' Range("A" & iRow + 1).Offset(0, 64).Value = TextBox64.Value
   ' Range("A" & iRow + 1).Offset(0, 61).Value = TextBox62.Value
   ' Range("A" & iRow + 1).Offset(0, 62).Value = TextBox63.Value
   
   
        For i = 1 To 63
        Controls("TextBox" & i).value = ""
    Next i

End Sub



Private Sub CommandButton20_Click()
On Error GoTo emad








'=======================================================
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 5) = TextBox5.Text And Sheets(1).Cells(Y, 2) = TextBox2.Text Then
   
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

Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton24_Click()
'On Error GoTo emad
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
'emad:

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
Application.ScreenUpdating = False
Sheets(15).Range("a8:L8341").ClearContents
Dim lastRow, lastrow1 As Integer

lastRow = Sheets(13).Cells(Rows.count, 2).End(xlUp).row

Dim lr, i As Integer
For i = 1 To lastRow
lr = Sheets(15).Cells(Rows.count, 1).End(xlUp).row
Sheets(15).Range("a1:l7").Copy
Sheets(15).Range("a" & lr + 1 - 1 + 3).PasteSpecial
Application.CutCopyMode = False
Sheets(15).Range("b" & lr + i - i + 3) = Sheets(13).Range("e" & i + 2)
Next i
If MsgBox(" „ «‰‘«¡ «‘—ÿ… «·—Ê« » ·„’‰⁄ «·ﬂ—«„Â ÊÂÌ Ã«Â“Â ··ÿ»«⁄… ", vbDefaultButton2, "") = vbOK Then
Application.ScreenUpdating = True
Application.Visible = True
Me.Hide

Sheets(15).Range("a1:L5351").PrintPreview
Application.Visible = False
Me.Show
Else
Exit Sub
End If
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

Ê—ﬁ…1.Range("x9:x650").FillDown
 Sheets(1).Cells(Y, 23) = TextBox23.Text
'=========================================

MsgBox " „  ⁄„Ì„ «·Ÿ—Ì»Â »‰Ã«Õ", vbDefaultButton1
End Sub

Private Sub CommandButton32_Click()
'============================== ‰ﬁ· „‰ ‘Ì  Œ«—ÃÌ «·Ï ﬁ«⁄œ… «·»Ì«‰«  «·Õ«·ÌÂ »ﬂ· «· ›«’Ì·

On Error Resume Next
Application.Visible = True
Call open_file
Application.Visible = True
Windows("rwtab_backup.xlsm").Activate
Sheets(1).Select
ActiveSheet.Range("b9:bv3000").Select
Selection.Copy
ThisWorkbook.Activate
Sheets(1).Select
Range("b9:bv3000").Select
Selection.PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
Windows("rwtab_backup.xlsm").Close
Application.Visible = False

Sheets(1).Range("a9").Formula = "=IF(B9="""","""", SUBTOTAL(3,$B$9:B9))"
Sheets(1).Range("a9:a3000").FillDown




Call copy_from_to_no_userform1

MsgBox "·ﬁœ  „ ‰ﬁ· Ê«” Ì—«œ «·»Ì«‰«  «·Œ«—ÃÌÂ »‰Ã«Õ", vbCritical, "—”«·…  ‰»ÌÂ"
End Sub

Private Sub CommandButton33_Click()
UserForm26.Show
End Sub

Private Sub CommandButton34_Click()

Call count_all
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




If TextBox58.Text = "  «Ã«“… »œÊ‰ —« »" Then
MsgBox " «Ã«“… »œÊ‰ —« »", vbDefaultButton2, "—”«·…  ÊÃÌÂ"



' ’›Ì— ﬂ· «·«” ﬁÿ«⁄« 
TextBox10.value = 0
TextBox24.value = 0
TextBox32.value = 0
TextBox33.value = 0
TextBox34.value = 0
TextBox35.value = 0
'TextBox36.value = 0
TextBox38.value = 0
'TextBox39.value = 0
'TextBox40.value = 0
TextBox42.value = 0
TextBox44.value = 0
TextBox46.value = 0



TextBox14.value = 0
TextBox16.value = 0
TextBox18.value = 0
TextBox19.value = 0
TextBox21.value = 0
TextBox20.value = 0
TextBox25.value = 0
TextBox27.value = 0
TextBox29.value = 0

'„»·€ «·€Ì«»

'TextBox49.Text = TextBox49.Text - TextBox48.Text
'=============================== «·«” ﬁÿ«⁄« 
'TextBox32.Text =
'TextBox33.Text =
'TextBox34.Text
'TextBox35.Text
'TextBox36.Text
'TextBox38.Text
'TextBox39.Text
'TextBox40.Text
'TextBox42.Text
'TextBox44.Text
'TextBox46.Text


'„Ã„Ê⁄ «·«” ﬁÿ«⁄« 
TextBox48.value = 0
TextBox49.value = 0
TextBox31.value = 0
TextBox30.value = 0
ComboBox6.value = 0
ComboBox7.value = 0
ComboBox8.value = 0
ComboBox10.value = 0
ComboBox11.value = 0


End If
'=================================
TextBox14.value = 0
TextBox16.value = 0
TextBox18.value = 0
'==============================„Œ’’«  «·«ÿ›«·




'==================================
TextBox27.value = 0
TextBox29.value = 0
'==========================„Ã„Ê⁄ «·„Œ’’« 
TextBox30.value = 0
'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ
TextBox32.value = 0
'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 

TextBox31.value = val(TextBox10.value) + TextBox30.value

'===============«·«” ﬁÿ«⁄« 
TextBox48.Text = 0

'=======================

'==================== „»·€ «·€Ì«»
TextBox49.Text = 0
TextBox12.Text = 0
TextBox47.Text = 0
'=======================«·—« » «·’«›Ì
TextBox49.Text = 0
TextBox49.Text = 0




'=========================================== «·«” ﬁÿ«⁄« 

If TextBox58.Text = "Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« " Then
On Error Resume Next
TextBox49.value = TextBox10.value - (TextBox32.value + TextBox33.value)

MsgBox "·œÌÂ «Ã«“Â Œ„” ”‰Ê« ", vbDefaultButton2, "—”«·…  ÊÃÌÂ"
' ’›Ì— ﬂ· «·„Œ’’« 
  TextBox11.Text = 0
 TextBox12.Text = 0
 TextBox13.Text = 0
 TextBox14.Text = 0
 TextBox15.Text = 0
 TextBox16.Text = 0
 TextBox15.Text = 0
 TextBox18.Text = 0
 TextBox17.Text = 0
  TextBox19.Text = 0
 TextBox20.Text = 0
  TextBox21.Text = 0
 TextBox22.Text = 0
 ' TextBox24.Text = 0
  TextBox25.Text = 0
   TextBox26.Text = 0
 TextBox27.Text = 0
 TextBox28.Text = 0
 TextBox29.Text = 0
 TextBox30.Text = 0
 TextBox35.Text = 0
 'TextBox38.Text = 0
'TextBox34.Text = 0
 'TextBox35.Text = 0
 'TextBox42.Text = 0
 'TextBox46.Text = 0

'==========================„Ã„Ê⁄ «·„Œ’’« 
TextBox30.value = val(TextBox14.value) + val(TextBox16.value) + val(TextBox18.value) + val(TextBox19.value) + val(TextBox21.value) + val(TextBox20.value) + val(TextBox25.value) + val(TextBox27.value) + val(TextBox29.value)
'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 

TextBox31.value = val(TextBox10.value) + val(TextBox14.value) + val(TextBox16.value) + val(TextBox18.value) + val(TextBox19.value) _
+ val(TextBox21.value) + val(TextBox20.value) + val(TextBox25.value) + val(TextBox27.value) + val(TextBox29.value)
TextBox32.value = CDbl(TextBox10.value * 0.1)
TextBox33.value = CDbl(TextBox10.value * 0.15)

'===============«·«” ﬁÿ«⁄« 
TextBox48.Text = val(TextBox24.Text) + val(TextBox32.Text) + val(TextBox33.Text) + val(TextBox34.Text) + val(TextBox35.Text) + val(TextBox36.Text) _
+ val(TextBox38.Text) + val(TextBox39.Text) + val(TextBox40.Text) + val(TextBox42.Text) + val(TextBox44.Text) + val(TextBox46.Text) + val(TextBox37.Text) + val(TextBox62.value)

'=======================«·—« » «·’«›Ì
TextBox49.Text = CDbl(TextBox31.Text) - CDbl(TextBox48.Text)
'=======================

'==================== „»·€ «·€Ì«»
TextBox47.Text = CDbl(TextBox49.Text) / 30 * CDbl(TextBox11.Text)







'========================================

'TextBox49.Text = TextBox31.Text - TextBox48.Text


End If



'==================================================«·’«›Ì
Dim vlookup1 As Long
Dim vlookup2 As Long
Dim vlookup3 As Long

Dim textvlue1 As Long
Dim textvlue2 As Long
Dim combvlue As Long
Dim rang1 As Range
Dim rang2 As Range




If TextBox58.Text = "„” „— »«·Œœ„…" Then
'=IF(AND(VLOOKUP(W9,“ÊÃÌ…4,2,0)<4,V9>0),"Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—",VLOOKUP(J9,«” ﬁÿ«⁄« ,VLOOKUP(W9,“ÊÃÌ…4,2,0)+V9,1))

Call state_tax


'=================================================


TextBox14.value = val(ComboBox6.value) * val(TextBox10.value)
TextBox16.value = val(ComboBox7.value) * val(TextBox10.value)
TextBox18.value = val(ComboBox8.value) * val(TextBox10.value)
TextBox20.value = val(TextBox22.value) * 10000
TextBox27.value = ComboBox10.value * val(TextBox10.value)
TextBox29.value = ComboBox11.value * val(TextBox10.value)
'==========================„Ã„Ê⁄ «·„Œ’’« 
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



TextBox30.value = val(TextBox14.value) + val(TextBox16.value) + val(TextBox18.value) + val(TextBox19.value) + val(TextBox21.value) + val(TextBox20.value) + val(TextBox25.value) + val(TextBox27.value) + val(TextBox29.value)

'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ
TextBox32.value = val(TextBox10.value) * 0.1
TextBox33.value = 0
'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 

TextBox31.value = val(TextBox10.value) + val(TextBox14.value) + val(TextBox16.value) + val(TextBox18.value) + val(TextBox19.value) _
+ val(TextBox21.value) + val(TextBox20.value) + val(TextBox25.value) + val(TextBox27.value) + val(TextBox29.value) + val(TextBox60.value)


'===============«·«” ﬁÿ«⁄« 
TextBox48.Text = val(TextBox24.Text) + val(TextBox32.Text) + val(TextBox33.Text) + val(TextBox34.Text) + val(TextBox35.Text) + val(TextBox36.Text) _
+ val(TextBox38.Text) + val(TextBox39.Text) + val(TextBox40.Text) + val(TextBox42.Text) + val(TextBox44.Text) + val(TextBox46.Text) _
+ val(TextBox37.Text) + val(TextBox62.value) + val(TextBox83.value + val(TextBox85.value))



'=======================

'==================== „»·€ «·€Ì«»
TextBox12.Text = 30 - val(TextBox11.Text)
TextBox47.Text = Format(val(TextBox49.Text / 30) * val(TextBox11.Text), "#.#")

'=======================«·—« » «·’«›Ì
TextBox49.Text = val(TextBox31.Text) - val(TextBox48.Text)
TextBox49.Text = val(TextBox49.Text) + val(TextBox84.Text) + val(TextBox86.Text) - val(TextBox47.Text)



End If








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
UserForm35.Show
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

End Sub

Private Sub CommandButton42_Click()

End Sub



Private Sub CommandButton6_Click()

On Error GoTo emad
 
'========================================================
If MsgBox("   ***Â·  —€» »«œŒ«· «·»Ì«‰«  ⁄»— „‰’… Ê«·Ê«ÃÂ… «·Õ«·ÌÂ  øø    **** " & vbCrLf & " Ê«–« ﬂ«‰ ·œÌﬂ ”Ã·«  „ Ê›—Â ›Ì „’‰› Œ«—ÃÌ „‰ «·„„ﬂ‰ «” Ì—«œÂ« ⁄»— »Ê«»… «·« ’«· «·ﬂ —Ê‰Ì «÷€ÿ ⁄·Ï “— NO    ", vbYesNo, "—”«·…  ÊÃÌÂ") = vbYes Then
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
     TextBox14.Text = ""
     TextBox15.Text = 0
   TextBox16.Text = ""
   TextBox17.Text = 0
   TextBox18.Text = ""
    TextBox19.Text = 0
   TextBox20.Text = ""
    TextBox21.Text = 0
  TextBox22.Text = 0
   TextBox23.Text = ""
TextBox24.Text = ""
   TextBox25.Text = 0
   TextBox26.Text = 0
   TextBox27.Text = ""
   TextBox28.Text = 0
  TextBox29.Text = ""
   TextBox30.Text = ""
 TextBox31.Text = ""
  TextBox32.Text = ""
   'TextBox33.Text = ""
   TextBox34.Text = 0
   TextBox35.Text = 0
   TextBox36.Text = 0
  TextBox37.Text = ""
 TextBox38.Text = 0
     TextBox39.Text = 0
    TextBox40.Text = 0
     TextBox41.Text = ""
   TextBox42.Text = 0
    TextBox43.Text = ""
 TextBox44.Text = 0
   TextBox45.Text = ""
    TextBox46.Text = 0
    TextBox47.Text = ""
   TextBox48.Text = ""
   TextBox49.Text = ""
      TextBox50.Text = ""
   TextBox52.Text = ""
   TextBox56.Text = ""
    TextBox51.Text = ""
   TextBox54.Text = ""
    TextBox55.Text = ""
   TextBox57.Text = ""
   
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
  Dim iRow As Long, i As Long

    Sheets(1).Activate
    iRow = Range("b" & Rows.count).End(xlUp).row
    TextBox1.Text = Application.WorksheetFunction.CountA(Sheets(1).Range("b9:b1000")) + 1
    ComboBox1.Visible = True
For i = 2 To Sheets(1).Cells(Rows.count, 63).End(xlUp).row
ComboBox1.AddItem
 ComboBox1.List(i - 2, 0) = Cells(i, "cb").value
 ComboBox1.List(i - 2, 1) = i
      Next i
   
   Else
   UserForm19.Show
   End If
emad:
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
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ad9:ad800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ae9:ae800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("av9:av800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aw9:aw800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("x9:x800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(1).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")






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
simsubtract = Application.WorksheetFunction.Sum(Sheets(1).Range("bj9:bj" & frw2))
TextBox73.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("j9:j" & frw2))
TextBox74.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("ad9:ad" & frw2))
TextBox79.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("ae9:ae" & frw2))
TextBox80.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("av9:av" & frw2))
TextBox66.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("aw9:aw" & frw2))
TextBox76.Text = Application.WorksheetFunction.Sum(Sheets(1).Range("X9:X" & frw1))
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
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ad9:ad800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ae9:ae800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("av9:av800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aw9:aw800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("x9:x800"), Sheets(1).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(1).Range("h9:h800"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(1).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")






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

Private Sub Label108_Click()


End Sub

Private Sub ListBox1_Click()
On Error Resume Next


Sheets(1).Activate
If TextBox2.Text = "" And TextBox5.Text = "" Then
Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
End If
For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) = True Then
        For j = 1 To 63
        Controls("TextBox" & j).Text = Cells(ListBox1.List(i, 1), j)
       
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

TextBox82.Text = Format(Sheets(1).Cells(frw1, "bn"), "yyyy/mm/dd")

TextBox83.Text = Sheets(1).Cells(frw1, "bo")
TextBox84.Text = Sheets(1).Cells(frw1, "bp")
TextBox85.Text = Sheets(1).Cells(frw1, "bq")
TextBox86.Text = Sheets(1).Cells(frw1, "br")
TextBox87.Text = Sheets(1).Cells(frw1, "bs")
TextBox88.Text = Sheets(1).Cells(frw1, "bt")
TextBox91.Text = Sheets(1).Cells(frw1, "bv")

'TextBox58Text = Sheets(1).Cells(frw1, "bg")

End If
Next frw1
If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If
If TextBox7.Text = "„Ã«“Ì‰" Then
MsgBox "«·„ÊŸ› " & " " & TextBox5.Text & "  " & "«·„Êﬁ› «·ÊŸÌ›Ì ·Â „‰ «·„Ã«“Ì‰ Õ«Ê· «⁄«œ… «Õ ”«» «·„Êﬁ› «·„«·Ì ·Â ·Ì ÿ«»ﬁ „⁄ ﬁ«‰Ê‰ «·«Ã«“«  ", vbInformation + vbOKOnly, "—”«·…  ‰»ÌÂ"


End If
TextBox89.value = Sheets(1).Cells(2, "bw").value
TextBox90.value = Sheets(1).Cells(2, "bx").value

If TextBox58.Text = "—« » Ã“∆Ì" Then
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

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox93.value = Format(UserForm1.TextBox93.value, "#,## IQD")
UserForm1.TextBox94.value = Format(UserForm1.TextBox94.value, "#,## IQD")
UserForm1.TextBox97.value = Format(UserForm1.TextBox97.value, "#,## IQD")
UserForm1.TextBox98.value = Format(UserForm1.TextBox98.value, "#,## IQD")
UserForm1.TextBox95.value = Format(UserForm1.TextBox95.value, "#,## IQD")
UserForm1.TextBox96.value = Format(UserForm1.TextBox96.value, "#,## IQD")

Label118.Caption = " ›«’Ì· «·Õ”«»«  Õ”» «·„Êﬁ› «·Ê÷Ì›Ì:" & "  " & TextBox58.Text







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

Private Sub TextBox11_Change()
TextBox12.value = 30 - val(TextBox11)

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

Private Sub TextBox59_Change()

End Sub

Private Sub TextBox60_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets(1).Activate
CommandButton5.Locked = True

For i = 9 To Sheets(1).Cells(Rows.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 9, 0) = Cells(i, 5).value
  ListBox1.List(i - 9, 1) = i
      Next i
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
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("e9:e571"), "*" & Me.TextBox67.Text & "*", Sheets(1).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")

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

Private Sub TextBox72_AfterUpdate()

End Sub

Private Sub TextBox72_Change()
On Error Resume Next
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

End Sub

Private Sub TextBox72_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Sheets(1).Activate

TextBox72.Text = ""


For i = 9 To Sheets(1).Cells(Rows.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 9, 0) = Cells(i, 5).value
  ListBox1.List(i - 9, 1) = i
      Next i
      If TextBox22.value = "" Then
TextBox20.value = ""
End If
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

Private Sub TextBox9_Change()
On Error Resume Next
Sheets(9).Cells(18, "g") = TextBox8.Text
Sheets(9).Cells(18, "f") = TextBox9.Text
TextBox10.Text = Sheets(9).Cells(18, "h")


Sheets(9).Cells(17, "q") = TextBox8.Text
Sheets(9).Cells(17, "r") = TextBox9.Text
TextBox91.Text = Sheets(9).Cells(17, "s")








If TextBox9.Text = "" Then
Exit Sub
End If
End Sub
Private Sub UserForm_Activate()

On Error GoTo emad

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)
ComboBox1.List = Sheets(4).Range("f1:f11").value
ComboBox2.List = Sheets(5).Range("a1:a373").value
ComboBox3.List = Sheets(4).Range("g1:g13").value

ComboBox4.List = Array("–ﬂ—", "«‰ÀÏ")
ComboBox5.List = Sheets(5).Range("a1:a373").value
ComboBox6.List = Array("0", "0.15", "0.25", "0.35", "0.45", "0.75", "1.25")
ComboBox7.List = Array("0", "0.35", "0.5")
ComboBox8.List = Array("0", "0.3", "0.25", "20", 15)
ComboBox9.List = Array("„«” —", "ﬂ«‘")
ComboBox10.List = Array("0", "0.3", "0.25", "0.20", 0.15)
ComboBox11.List = Array("0", "0.3", "0.25", "20", 15)






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


'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(1).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578,BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")









emad:
'Else
'CommandButton5.Enabled = False
'End If
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
Label104.Caption = Sheets("title_factory").Cells(2, 1)

Label105.Caption = Sheets("title_factory").Cells(2, 3)

End Sub


