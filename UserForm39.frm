VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm39 
   Caption         =   "UserForm39"
   ClientHeight    =   11505
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   23400
   OleObjectBlob   =   "UserForm39.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox24_Click()

End Sub

Private Sub ComboBox1_Change()
On Error Resume Next
ListBox1.Selected(ListBox1.ListIndex) = False
If TextBox5.Text <> "" And TextBox6.Text <> "" And TextBox7.Text <> "" And TextBox8.Text <> "" _
And TextBox9.Text <> "" And TextBox10.Text <> "" And TextBox11.Text <> "" And TextBox12.Text <> "" _
And TextBox13.Text <> "" And TextBox14.Text <> "" And TextBox15.Text <> "" And TextBox16.Text <> "" _
 And TextBox17.Text <> "" And TextBox18.Text <> "" And TextBox19.Text <> "" And TextBox20.Text <> "" _
And TextBox21.Text <> "" And TextBox22.Text <> "" And TextBox23.Text <> "" And TextBox24.Text <> "" Then
MsgBox "«Ê·« ﬁ„ »« ›—Ì€ «·»Ì«‰«  ·Ì ”‰Ï ··‰Ÿ«„ «⁄«œ…  — Ì» ﬁÌ„ «· ﬁ—Ì—", vbCritical, "—”«·…  ‰»ÌÂ"
Exit Sub
CommandButton6.SetFocus

Else


Dim criterion
Dim c As Integer
Dim column_headers As Variant
column_headers = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "Aj", "Ak", "Al", "Am", "An", "Ao", "Ap" _
, "Aq", "Ar", "As", "At", "Au", "Av", "Aw", "Ax", "Ay", "Az", "ba", "bb", "bc", "bd", "be", "bf", "bg", "bh", "bi", "bj", "bk", "bl", "bm", "bn", "bo", "bp", "bq")
' , "br", "bs", "bt", "bu", "bv", "bw", "bx", "by", "bz" _
' , "ca", "cb", "cc", "cd")
For c = 1 To 13
If Sheets(43).Cells(8, c).value = Me.ComboBox1.value Then
criterion = column_headers(c - 1)
End If
 Next
 Sheets(4).Cells(1, "b").value = criterion
TextBox1.Text = ComboBox1.ListIndex + 1
TextBox25.Text = Sheets(4).Cells(1, "b")
'==============================================













'=========================================

If TextBox5.Text = "" Then
TextBox5.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox5.Text









 
'========================
ElseIf TextBox6.Text = "" Then
TextBox6.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox6.Text







ElseIf TextBox7.Text = "" Then
TextBox7.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox7.Text








ElseIf TextBox8.Text = "" Then
TextBox8.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox8.Text










ElseIf TextBox9.Text = "" Then
TextBox9.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox9.Text







ElseIf TextBox10.Text = "" Then
TextBox10.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox10.Text








ElseIf TextBox11.Text = "" Then
TextBox11.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox11.Text





ElseIf TextBox12.Text = "" Then
TextBox12.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox12.Text










ElseIf TextBox13.Text = "" Then
TextBox13.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox13.Text









ElseIf TextBox14.Text = "" Then
TextBox14.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox14.Text









ElseIf TextBox15.Text = "" Then
TextBox15.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox15.Text








ElseIf TextBox16.Text = "" Then
TextBox16.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox16.Text








ElseIf TextBox17.Text = "" Then
TextBox17.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox17.Text









ElseIf TextBox18.Text = "" Then
TextBox18.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox18.Text










ElseIf TextBox19.Text = "" Then
TextBox19.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox19.Text









ElseIf TextBox20.Text = "" Then
TextBox20.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox20.Text








ElseIf TextBox21.Text = "" Then
TextBox21.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox21.Text











ElseIf TextBox22.Text = "" Then
TextBox22.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox22.Text










ElseIf TextBox23.Text = "" Then
TextBox23.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox23.Text











ElseIf TextBox24.Text = "" Then
TextBox24.Text = TextBox25.Text
ListBox1.AddItem ComboBox1.value
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox1.Text
ListBox1.List(ListBox1.ListCount - 1, 2) = TextBox24.Text










End If
TextBox26.value = ListBox1.ListCount




 

End If
CommandButton4.Enabled = True
CommandButton21.Enabled = True
CommandButton8.Enabled = True
CommandButton1.Enabled = True

CommandButton20.Enabled = True
CommandButton7.Enabled = True
CommandButton11.Enabled = True
CommandButton6.Enabled = True
CommandButton3.Enabled = True
emad:
End Sub



Private Sub CommandButton1_Click()


Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Dim lastR As Long
lastR = Sheets(43).Cells(Rows.count, 1).End(xlUp).row

Sheets(43).Range("a3:m" & lastR).PrintPreview
ThisWorkbook.Application.Visible = False
UserForm32.Show
End Sub

Private Sub CommandButton10_Click()
Call Showit
Application.Visible = True
End Sub

Private Sub CommandButton11_Click()

Call vba_border_sheet_44_1


UserForm40.Show
End Sub

Private Sub CommandButton2_Click()
Unload Me
UserForm10.Show

End Sub

Private Sub CommandButton20_Click()
Call all_data_user39
Call sum_report_44_1
Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Dim lastR As Long
lastR = Sheets(43).Cells(Rows.count, 1).End(xlUp).row

Sheets(43).Range("a3:m" & lastR).PrintPreview
ThisWorkbook.Application.Visible = False
UserForm39.Show





End Sub

Private Sub CommandButton21_Click()


Call vba_border_sheet_44_1

Call case_filter_form_39_1




' Ê—ﬁ…42.Range("a9:m36").AutoFilter Field:=1, Criteria1:=Sheets("Sheet44").Cells(1, "r"), Operator:=xlOr, Criteria2:=Sheets("Sheet44").Cells(2, "r", Operator:=xlOr, Criteria3:=Sheets("Sheet44").Cells(3, "r"), Operator:=xlOr, Criteria4:=Sheets("Sheet44").Cells(4, "r"), Operator:=xlOr, _
' Criteria5:=Sheets("Sheet44").Cells(5, "r"), Operator:=xlOr, Criteria6:=Sheets("Sheet44").Cells(6, "r"), Operator:=xlOr, _
'Criteria7:=Sheets("Sheet44").Cells(7, "r"), Operator:=xlOr, Criteria8:=Sheets("Sheet44").Cells(8, "r"), Operator:=xlOr, Criteria9:=Sheets("Sheet44").Cells(9, "r"), Operator:=xlOr, Criteria10:=Sheets("Sheet44").Cells(10, "r"), Operator:=xlOr, Criteria11:=Sheets("Sheet44").Cells(11, "r"), Operator:=xlOr, _
'Criteria12:=Sheets("Sheet44").Cells(12, "r"), Operator:=xlOr, Criteria13:=Sheets("Sheet44").Cells(13, "r", Operator:=xlOr))

































'«·—« » «·«”„Ì
'„Œ’’«  ‘Â«œ…
'„Œ’’«  Â‰œ”Ì…
'„Œ’’«  „‰’»
'„Œ’’«  „Êﬁ⁄ Ã€—«›Ì
'„Œ’’«  √ÿ›«·
'„Œ’’«  “ÊÃÌ…
'„Œ’’«  ŒÿÊ—… „»·€
'„Œ’’«  „Â‰ÌÂ
'„Œ’’«  ŒÿÊ—Â ‰”»…
'«Ÿ«›… „«·ÌÂ „ƒﬁ Â
'«÷«›Â „«·ÌÂ ⁄«„…
'«÷«›Â „«·ÌÂ ‰”»…
'«·«” ﬁÿ«⁄ «·÷—Ì»Ì
'«· Êﬁ›«  «· ﬁ«⁄œÌ… 10%
'«· Êﬁ›«  «· ﬁ«⁄œÌ… 15%
'«· «„Ì‰ «·’ÕÌ
'√ÃÊ— «·‰ﬁ·
'«·‰›ﬁ…
'—”Ê„ «·ÿ«»⁄
'«” ﬁÿ«⁄ 1
'‰«œÌ «·ÂÌ∆…
'«” ﬁÿ«⁄ „ ›—ﬁ
'«” ﬁÿ«⁄ —⁄«Ì… «Ã „«⁄Ì…
'„»·€ «·€Ì«»« 
'«” ﬁÿ«⁄ „ƒﬁ 
'«” ﬁÿ«⁄ ⁄«„
'«” ﬁÿ«⁄ ⁄«„ ‰”»…





End Sub



Private Sub CommandButton3_Click()
Sheets(43).Range("a:m").EntireColumn.Hidden = True
End Sub

Private Sub CommandButton4_Click()
On Error Resume Next
Sheets(43).Range(TextBox25.Text & ":" & TextBox25.Text).EntireColumn.Hidden = False

End Sub

Private Sub CommandButton5_Click()
On Error Resume Next
Sheets(43).Range(TextBox25.Text & ":" & TextBox25.Text).EntireColumn.Hidden = True
End Sub

Private Sub CommandButton6_Click()
Sheets(43).Range(TextBox25.Text & ":" & TextBox25.Text).EntireColumn.Hidden = True
End Sub

Private Sub CommandButton7_Click()
On Error Resume Next


Sheets(44).Range("a:m").EntireColumn.Hidden = True
Sheets(44).Range(TextBox5.Text & ":" & TextBox5.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox6.Text & ":" & TextBox6.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox7.Text & ":" & TextBox7.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox8.Text & ":" & TextBox8.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox9.Text & ":" & TextBox9.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox10.Text & ":" & TextBox10.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox11.Text & ":" & TextBox11.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox12.Text & ":" & TextBox12.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox13.Text & ":" & TextBox13.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox14.Text & ":" & TextBox14.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox15.Text & ":" & TextBox15.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox16.Text & ":" & TextBox16.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox17.Text & ":" & TextBox17.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox18.Text & ":" & TextBox18.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox19.Text & ":" & TextBox19.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox20.Text & ":" & TextBox20.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox21.Text & ":" & TextBox21.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox22.Text & ":" & TextBox22.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox23.Text & ":" & TextBox23.Text).EntireColumn.Hidden = False
Sheets(44).Range(TextBox24.Text & ":" & TextBox24.Text).EntireColumn.Hidden = False






End Sub

Private Sub CommandButton8_Click()
On Error Resume Next
Dim lastR As Integer

'lastr = Sheets(39).Cells(Rows.Count, 1).End(xlUp).Row
 
'Range("a7:c" & lastr) = ""

ListBox1.Clear
TextBox5.Text = ""
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
TextBox20.Text = ""
TextBox21.Text = ""
TextBox22.Text = ""
TextBox23.Text = ""
TextBox24.Text = ""
TextBox1.Text = ""
TextBox25.Text = ""
TextBox26.Text = ""
TextBox27.Text = ""
ComboBox1.value = ""


End Sub

Private Sub CommandButton9_Click()
On Error Resume Next

Call all_data_user39
Call vba_border_sheet_44_1
Sheets(44).Range("a:m").EntireColumn.Hidden = False
Call all_data_user39

ListBox1.Clear
TextBox5.Text = ""
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
TextBox20.Text = ""
TextBox21.Text = ""
TextBox22.Text = ""
TextBox23.Text = ""
TextBox24.Text = ""
TextBox1.Text = ""
TextBox25.Text = ""
TextBox26.Text = ""
TextBox27.Text = ""
ComboBox1.value = ""




MsgBox " „  ⁄„·Ì… «Õ÷«— «·»Ì«‰«  «·ﬂ·Ì… «·⁄«„Â ··‰Ÿ«„", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
Unload Me
UserForm39.Show

End Sub

Private Sub Frame3_Click()

End Sub

Private Sub Frame4_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub ListBox1_Click()
On Error Resume Next
ListView1.ListItems.Clear
TextBox26.value = ListBox1.ListCount
Dim i As Integer
For i = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(i) Then
ListView1.ListItems.Clear
TextBox27.Text = ListBox1.List(i, 0)
TextBox1.Text = ListBox1.List(i, 1)
TextBox25.Text = ListBox1.List(i, 2)
Sheets(29).Cells(8 + i, 1) = ListBox1.List(i, 0)
 Sheets(29).Cells(8 + i, 2) = TextBox1.Text
 Sheets(29).Cells(8 + i, 3) = TextBox25.Text

End If

Next





'===================================
 
 



 
 
 
 
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim r As Integer
For r = 0 To ListBox1.ListCount
If ListBox1.Selected(r) = True Then
ListBox1.RemoveItem (r)
End If
Next
End Sub

Private Sub ListBox1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
ListBox1.Selected(ListBox1.ListIndex) = False
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox10_Change()
On Error Resume Next
Sheets(29).Cells(2, 2) = TextBox10.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox11_Change()
On Error Resume Next
Sheets(29).Cells(3, 2) = TextBox11.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox12_Change()
On Error Resume Next
Sheets(29).Cells(4, 2) = TextBox12.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox13_Change()
On Error Resume Next
Dim r As Integer
Sheets(29).Cells(1, 3) = TextBox13.Text
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox14_Change()
On Error Resume Next
Sheets(29).Cells(2, 3) = TextBox14.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox15_Change()
On Error Resume Next
Sheets(29).Cells(3, 3) = TextBox15.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox16_Change()
On Error Resume Next
Sheets(29).Cells(4, 3) = TextBox16.Text
Dim r As Integer

For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox17_Change()
On Error Resume Next
Sheets(29).Cells(1, 4) = TextBox17.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox18_Change()
On Error Resume Next
Sheets(29).Cells(2, 4) = TextBox18.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox19_Change()
On Error Resume Next
Sheets(29).Cells(3, 4) = TextBox19.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
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



Private Sub TextBox20_Change()
On Error Resume Next
Sheets(29).Cells(4, 4) = TextBox20.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox21_Change()
On Error Resume Next
Sheets(29).Cells(1, 5) = TextBox21.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox22_Change()
On Error Resume Next
Sheets(29).Cells(2, 5) = TextBox22.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox23_Change()
On Error Resume Next
Sheets(29).Cells(3, 5) = TextBox23.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox24_Change()
On Error Resume Next
Sheets(29).Cells(4, 5) = TextBox24.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox26_Change()

End Sub

Private Sub TextBox27_Change()

End Sub

Private Sub TextBox28_Change()
Call delete_amployee_month
End Sub

Private Sub TextBox5_AfterUpdate()

End Sub

Private Sub TextBox5_Change()
On Error Resume Next
Sheets(29).Cells(1, 1) = TextBox5.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox5_Enter()

End Sub



Private Sub TextBox6_Change()
On Error Resume Next
Sheets(29).Cells(2, 1) = TextBox6.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub TextBox7_Change()
On Error Resume Next
Sheets(29).Cells(3, 1) = TextBox7.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox7_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

End Sub

Private Sub TextBox8_Change()
On Error Resume Next
Sheets(29).Cells(4, 1) = TextBox8.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub TextBox9_Change()
On Error Resume Next
Sheets(29).Cells(1, 2) = TextBox9.Text
Dim r As Integer
For r = 0 To ListBox1.ListCount - 1
If ListBox1.Selected(r) = True Then
 ListBox1.RemoveItem (r)
 End If
 Next
End Sub

Private Sub UserForm_Activate()
On Error Resume Next
TextBox3.Text = Sheets(44).Cells(Rows.count, 1).End(xlUp).row
TextBox4.Text = Sheets(44).Cells(Rows.count, 1).End(xlUp).row - 8

'ListBox1.RowSource = "a8:c30"
TextBox26.value = ListBox1.ListCount

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True


.ColumnHeaders.Add , , " ›«’Ì· «·—« »", 150
.ColumnHeaders.Add , , "Ì‰«Ì—", 100
.ColumnHeaders.Add , , "›»—«Ì—", 100
.ColumnHeaders.Add , , "„«—”", 100
.ColumnHeaders.Add , , "√»—Ì·", 100
.ColumnHeaders.Add , , "„«ÌÊ", 100
.ColumnHeaders.Add , , " ÌÊ‰ÌÊ", 100
.ColumnHeaders.Add , , " ÌÊ·ÌÊ", 100
.ColumnHeaders.Add , , "√€”ÿ”", 100
.ColumnHeaders.Add , , " ”» „»— ", 100
.ColumnHeaders.Add , , "√ﬂ Ê»—", 100
.ColumnHeaders.Add , , " ‰Ê›„»— ", 100
.ColumnHeaders.Add , , "œÌ”„»—  ", 100

'   Ì‰«Ì—
'   ›»—«Ì—
 ' „«—”
'  √»—Ì·
'  „«ÌÊ
'  ÌÊ‰ÌÊ
'  ÌÊ·ÌÊ
'    √€”ÿ”
'  ”» „»—
'  √ﬂ Ê»—
'  ‰Ê›„»—
'  œÌ”„»—



ListView1.Font = Bold
ListView1.Font.Size = 11
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
.Font.Size = 12 'ÕÃ„ «·Œÿ
.Font.Bold = True  '€«„ﬁ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(43).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(43).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(43).Cells(frw1, "b")
item1.SubItems(2) = Sheets(43).Cells(frw1, "c")
item1.SubItems(3) = Sheets(43).Cells(frw1, "d")
item1.SubItems(4) = Sheets(43).Cells(frw1, "e")
item1.SubItems(5) = Sheets(43).Cells(frw1, "f")
item1.SubItems(6) = Sheets(43).Cells(frw1, "g")
item1.SubItems(7) = Sheets(43).Cells(frw1, "h")
item1.SubItems(8) = Sheets(43).Cells(frw1, "i")
item1.SubItems(9) = Sheets(43).Cells(frw1, "j")
item1.SubItems(10) = Sheets(43).Cells(frw1, "k")
item1.SubItems(11) = Sheets(43).Cells(frw1, "l")
item1.SubItems(12) = Sheets(43).Cells(frw1, "m")






















   


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
   








'=================================================




End With
End With



ListView1.Font = Bold
ListView1.Font.Size = 11
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
 
CommandButton4.Enabled = False
CommandButton21.Enabled = False
CommandButton8.Enabled = False
CommandButton1.Enabled = False

CommandButton20.Enabled = False
CommandButton7.Enabled = False
CommandButton11.Enabled = False
CommandButton6.Enabled = False
CommandButton3.Enabled = False


End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()
On Error GoTo emad

 TextBox5.Text = Sheets(29).Cells(1, 1)
TextBox6.Text = Sheets(29).Cells(2, 1)
TextBox7.Text = Sheets(29).Cells(3, 1)
TextBox8.Text = Sheets(29).Cells(4, 1)

TextBox9.Text = Sheets(29).Cells(1, 2)
TextBox10.Text = Sheets(29).Cells(2, 2)
TextBox11.Text = Sheets(29).Cells(3, 2)
TextBox12.Text = Sheets(29).Cells(4, 2)


TextBox13.Text = Sheets(29).Cells(1, 3)
TextBox14.Text = Sheets(29).Cells(2, 3)
TextBox15.Text = Sheets(29).Cells(3, 3)
TextBox16.Text = Sheets(29).Cells(4, 3)

TextBox17.Text = Sheets(29).Cells(1, 4)
TextBox18.Text = Sheets(29).Cells(2, 4)
TextBox19.Text = Sheets(29).Cells(3, 4)
TextBox20.Text = Sheets(29).Cells(4, 4)

TextBox21.Text = Sheets(29).Cells(1, 5)
TextBox22.Text = Sheets(29).Cells(2, 5)
TextBox23.Text = Sheets(29).Cells(3, 5)
TextBox24.Text = Sheets(29).Cells(4, 5)


Dim i, Y As Integer


For i = 1 To 13
Me.ComboBox1.AddItem Sheets(43).Cells(8, i).value


Next i

ComboBox8.RowSource = "col_ltr"
' ComboBox6.List = Array("sheet14", "")


'Dim lastr As Integer

'lastr = Sheets(39).Cells(Rows.Count, 1).End(xlUp).Row
 

'=========================
emad:
End Sub





