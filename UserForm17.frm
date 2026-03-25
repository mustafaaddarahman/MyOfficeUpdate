VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm17 
   Caption         =   "UserForm17"
   ClientHeight    =   9720.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   19152
   OleObjectBlob   =   "UserForm17.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_AfterUpdate()

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
column_headers = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "as", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "Bl", "Bm", "Bn", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV")
For c = 1 To 74
If Sheets(13).Cells(8, c).value = Me.ComboBox1.value Then
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

emad:
End Sub



Private Sub CommandButton1_Click()


Unload Me

ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Dim lastR As Long
lastR = Sheets(28).Cells(Rows.count, 1).End(xlUp).row + 10

Sheets(28).Range("a8:bv" & lastR).PrintPreview
ThisWorkbook.Application.Visible = False
UserForm17.Show
End Sub

Private Sub CommandButton10_Click()
Call all_data_userform17

TextBox3.value = Sheets(28).Cells(Rows.count, 2).End(xlUp).row + 2

Sheets(28).Cells(TextBox3.Text, "a") = "«·„Ã„Ê⁄ «·ﬂ·Ì"
Sheets(28).Cells(TextBox3.value, "j") = Application.WorksheetFunction.Sum(Sheets(28).Range("j9" & " : " & "j" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "x") = Application.WorksheetFunction.Sum(Sheets(28).Range("x9" & " : " & "x" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "ad") = Application.WorksheetFunction.Sum(Sheets(28).Range("ad9" & " : " & "ad" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "n") = Application.WorksheetFunction.Sum(Sheets(28).Range("n9" & " : " & "n" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "p") = Application.WorksheetFunction.Sum(Sheets(28).Range("p9" & " : " & "p" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "r") = Application.WorksheetFunction.Sum(Sheets(28).Range("r9" & " : " & "r" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "s") = Application.WorksheetFunction.Sum(Sheets(28).Range("s9" & " : " & "s" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "t") = Application.WorksheetFunction.Sum(Sheets(28).Range("t9" & " : " & "t" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "u") = Application.WorksheetFunction.Sum(Sheets(28).Range("u9" & " : " & "u" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "y") = Application.WorksheetFunction.Sum(Sheets(28).Range("y9" & " : " & "y" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "aa") = Application.WorksheetFunction.Sum(Sheets(28).Range("aa9" & " : " & "aa" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "ac") = Application.WorksheetFunction.Sum(Sheets(28).Range("ac9" & " : " & "ac" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "ae") = Application.WorksheetFunction.Sum(Sheets(28).Range("ae9" & " : " & "ae" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "aw") = Application.WorksheetFunction.Sum(Sheets(28).Range("aw9" & " : " & "aw" & TextBox3.value))



Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Dim lastR As Long
lastR = Sheets(28).Cells(Rows.count, 2).End(xlUp).row

'==================================================== «·€«¡ œ„Ã «·Œ·«Ì«
TextBox3.value = Sheets(28).Cells(Rows.count, 2).End(xlUp).row + 2

Sheets(28).Cells(TextBox3.Text, "a") = "«·„Ã„Ê⁄ «·ﬂ·Ì"
Sheets(28).Cells(TextBox3.value, "j") = Application.WorksheetFunction.Sum(Sheets(28).Range("j9" & " : " & "j" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "x") = Application.WorksheetFunction.Sum(Sheets(28).Range("x9" & " : " & "x" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "ad") = Application.WorksheetFunction.Sum(Sheets(28).Range("ad9" & " : " & "ad" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "n") = Application.WorksheetFunction.Sum(Sheets(28).Range("n9" & " : " & "n" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "p") = Application.WorksheetFunction.Sum(Sheets(28).Range("p9" & " : " & "p" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "r") = Application.WorksheetFunction.Sum(Sheets(28).Range("r9" & " : " & "r" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "s") = Application.WorksheetFunction.Sum(Sheets(28).Range("s9" & " : " & "s" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "t") = Application.WorksheetFunction.Sum(Sheets(28).Range("t9" & " : " & "t" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "u") = Application.WorksheetFunction.Sum(Sheets(28).Range("u9" & " : " & "u" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "y") = Application.WorksheetFunction.Sum(Sheets(28).Range("y9" & " : " & "y" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "aa") = Application.WorksheetFunction.Sum(Sheets(28).Range("aa9" & " : " & "aa" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "ac") = Application.WorksheetFunction.Sum(Sheets(28).Range("ac9" & " : " & "ac" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "ae") = Application.WorksheetFunction.Sum(Sheets(28).Range("ae9" & " : " & "ae" & TextBox3.value))
Sheets(28).Cells(TextBox3.value, "aw") = Application.WorksheetFunction.Sum(Sheets(28).Range("aw9" & " : " & "aw" & TextBox3.value))














'===============================================


Sheets(28).Range("a8:bv" & lastR + 4).PrintPreview



ThisWorkbook.Application.Visible = False
UserForm17.Show
End Sub

Private Sub CommandButton11_Click()
Call copy_from_to_no_userform1
End Sub

Private Sub CommandButton13_Click()

End Sub

Private Sub CommandButton14_Click()

End Sub

Private Sub CommandButton18_Click()

End Sub

Private Sub CommandButton19_Click()
Call vba_border_sheet_28_1

UserForm36.Show

End Sub

Private Sub CommandButton2_Click()
Unload Me
UserForm10.Show
End Sub



Private Sub CommandButton21_Click()

End Sub

Private Sub CommandButton3_Click()
On Error Resume Next


Sheets(28).Range(TextBox5.Text & ":" & TextBox5.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox6.Text & ":" & TextBox6.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox7.Text & ":" & TextBox7.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox8.Text & ":" & TextBox8.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox9.Text & ":" & TextBox9.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox10.Text & ":" & TextBox10.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox11.Text & ":" & TextBox11.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox12.Text & ":" & TextBox12.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox13.Text & ":" & TextBox13.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox14.Text & ":" & TextBox14.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox15.Text & ":" & TextBox15.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox16.Text & ":" & TextBox16.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox17.Text & ":" & TextBox17.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox18.Text & ":" & TextBox18.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox19.Text & ":" & TextBox19.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox20.Text & ":" & TextBox20.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox21.Text & ":" & TextBox21.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox22.Text & ":" & TextBox22.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox23.Text & ":" & TextBox23.Text).EntireColumn.Hidden = True
Sheets(28).Range(TextBox24.Text & ":" & TextBox24.Text).EntireColumn.Hidden = True








End Sub

Private Sub CommandButton4_Click()
On Error Resume Next
Sheets(28).Range(TextBox25.Text & ":" & TextBox25.Text).EntireColumn.Hidden = False

End Sub

Private Sub CommandButton5_Click()
On Error Resume Next
Sheets(28).Range(TextBox25.Text & ":" & TextBox25.Text).EntireColumn.Hidden = True
End Sub

Private Sub CommandButton6_Click()
Sheets(28).Range(TextBox25.Text & ":" & TextBox25.Text).EntireColumn.Hidden = True
End Sub

Private Sub CommandButton7_Click()
On Error Resume Next

Sheets(28).Range("a:bv").EntireColumn.Hidden = True
Sheets(28).Range(TextBox5.Text & ":" & TextBox5.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox6.Text & ":" & TextBox6.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox7.Text & ":" & TextBox7.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox8.Text & ":" & TextBox8.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox9.Text & ":" & TextBox9.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox10.Text & ":" & TextBox10.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox11.Text & ":" & TextBox11.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox12.Text & ":" & TextBox12.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox13.Text & ":" & TextBox13.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox14.Text & ":" & TextBox14.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox15.Text & ":" & TextBox15.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox16.Text & ":" & TextBox16.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox17.Text & ":" & TextBox17.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox18.Text & ":" & TextBox18.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox19.Text & ":" & TextBox19.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox20.Text & ":" & TextBox20.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox21.Text & ":" & TextBox21.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox22.Text & ":" & TextBox22.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox23.Text & ":" & TextBox23.Text).EntireColumn.Hidden = False
Sheets(28).Range(TextBox24.Text & ":" & TextBox24.Text).EntireColumn.Hidden = False
MsgBox "ﬁœ  „ «Õ÷«— «·»Ì«‰«  Õ”» «·„⁄ÿÌ«  «·„ÿ·Ê»Â", vbDefaultButton1, "—”«·…  ‰»ÌÂ"







End Sub

Private Sub CommandButton8_Click()
On Error Resume Next

Dim lastR As Integer

'lastr = Sheets(29).Cells(Rows.Count, 1).End(xlUp).Row
 
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
Label11.Caption = ""

TextBox28.Text = ""
TextBox28.Text = Now

'Call start_time


TextBox3.Text = Sheets(28).Cells(Rows.count, 1).End(xlUp).row + 2
TextBox4.Text = Sheets(28).Cells(Rows.count, 1).End(xlUp).row - 8
Sheets(28).Activate
Range("a:bz").EntireColumn.Hidden = False

Call all_data_userform17
Call sum_filter_form17_1
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

TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)



MsgBox " „  ⁄„·Ì… «Õ÷«— «·»Ì«‰«  «·ﬂ·Ì… «·⁄«„Â ··‰Ÿ«„", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label13_Click()

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

Private Sub TextBox10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox10.Text = ""
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

Private Sub TextBox11_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox11.Text = ""
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

Private Sub TextBox12_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox12.Text = ""
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

Private Sub TextBox13_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox13.Text = ""
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

Private Sub TextBox14_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox14.Text = ""
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

Private Sub TextBox15_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox15.Text = ""
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

Private Sub TextBox16_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox16.Text = ""
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

Private Sub TextBox17_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox17.Text = ""
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

Private Sub TextBox18_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox18.Text = ""
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

Private Sub TextBox19_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox19.Text = ""
End Sub

Private Sub TextBox2_Change()
'On Error GoTo emad

Select Case TextBox1.value

Case 2 '«·—ﬁ„ «·ÊŸÌ›Ì

Call case_2

'=============================



Case 4  '«·Ã‰”


Call case_4



Case 5  '«”„ «·„ÊŸ›
On Error Resume Next

Call case_5


Case 6  '«·⁄‰Ê«‰ «·ÊŸÌ›Ì


Call case_6



Case 7  '«·ﬁ”„


Call case_7



Case 8 '«·œ—Ã…

On Error Resume Next

Call case_8





Case 9  '«·„—Õ·…
On Error Resume Next

Call case_9




Case 10  '«·—« » «·«”„Ì

On Error Resume Next
Call case_10


Case 11  '⁄œœ «Ì«„ «·€Ì«»

On Error Resume Next
Call case_11




Case 13
    '‰”»… «·‘Â«œ…
On Error Resume Next

Call case_13




Case 15
     '‰”»… «·Â‰œ”ÌÂ

On Error Resume Next

Call case_15





Case 17
     '‰”»… «·„‰’»

Call case_17





Case 19
       '„Œ’’«  „Êﬁ⁄ Ã€—«›Ì


Call case_19







Case 20
      '„Œ’’«  «ÿ›«·
Call case_20




Case 21
   '„Œ’’«  “ÊÃÌÂ

Call case_21





Case 22
    '⁄œœ «·«ÿ›«·

Call case_22



Case 23
   '«·Õ«·… «·«Ã „«⁄Ì…

Call case_23



Case 24    '«·«” ﬁÿ«⁄ «·÷—»Ì

Call case_24






Case 25    '„Œ’’«  ŒÿÊ—… „»·€


Call case_25




Case 26    '‰”»… „Â‰ÌÂ

Call case_26





Case 28    '‰”»… «·ŒÿÊ—…

Call case_28




Case 32     '«· Êﬁ›«  «· ﬁ«⁄œÌ… 10%

Call case_32



Case 33   '«· Êﬁ›«  «· ﬁ«⁄œÌ… 15


Call case_33




Case 34    '«· «„Ì‰ «·’ÕÌ

Call case_34





Case 35    '«ÃÊ— «·‰ﬁ·
Call case_35





Case 36   '«·‰›ﬁ…
Call case_36






Case 40   '«·«” ﬁÿ«⁄ «·«Ê·

Call case_40





Case 44  '«·«” ﬁÿ«⁄ «·„ ›—ﬁ


Call case_44






Case 47   '„»·€ «·€Ì«»« 
Call case_47


Case 49   ' «·—« » «·ﬂ·Ì
Call case_49



Case 51   'ÿ—Ìﬁ… «·œ›⁄

Call case_51



Case 53   '«·«÷«›…

Call case53_add1



Case 57   ' ’›Ì… Õ”» «·ÃÂ…
Call filter_aljha57





Case 58    '«·„Êﬁ› «·ÊŸÌ›Ì

Call case_58





Case 60  '«·«÷«›… «·„«·ÌÂ «·„ƒﬁ Â
Call case_60




Case 61  '«· ⁄œÌ·
Call case61_modyfay1





Case 62  '«·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ 

Call case_62

Case 66  ' «—ÌŒ «Œ—  ÕœÌÀ

Call case_66




Case 67         '«” —Ã«⁄ ⁄«„
Call Case_67

Case 68  '«” ﬁÿ«⁄ ⁄«„
Call Case_68

Case 69     '«” ﬁÿ«⁄ ‰”»…
Call Case_69

Case 70     '«” —Ã«⁄ ‰”»…
Call Case_70


Case 73     '„·«Õ÷« 
Call Case_73




 Case 74     '„ﬁœ«— «·“Ì«œÂ ··⁄·«Ê… «Ê «· —›Ì⁄
Call Case_74


End Select
TextBox30.value = ListView1.ListItems.count

emad:

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

Private Sub TextBox20_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox20.Text = ""
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

Private Sub TextBox21_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox21.Text = ""
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

Private Sub TextBox22_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox22.Text = ""
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

Private Sub TextBox23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox23.Text = ""
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

Private Sub TextBox24_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox24.Text = ""
End Sub

Private Sub TextBox27_Change()

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

Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox5.Text = ""
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

Private Sub TextBox6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox6.Text = ""
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

Private Sub TextBox7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox7.Text = ""
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

Private Sub TextBox8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox8.Text = ""
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

Private Sub TextBox9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox9.Text = ""
End Sub

Private Sub UserForm_Activate()
On Error Resume Next
Call copy_from_to_t002
Call sum_report_fotm17_1
TextBox3.Text = Sheets(13).Cells(Rows.count, 2).End(xlUp).row + 2
TextBox4.Text = Sheets(13).Cells(Rows.count, 2).End(xlUp).row - 8
'Sheets(29).Activate
'ListBox1.RowSource = "a8:c30"
TextBox26.value = ListBox1.ListCount

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
.ColumnHeaders.Add , , " «· ⁄œÌ·", 140
.ColumnHeaders.Add , , " «·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ  ", 140
.ColumnHeaders.Add , , " —„“ «·«” Õﬁ«ﬁ ", 140
.ColumnHeaders.Add , , "„”«— «·’Ê—…", 140
.ColumnHeaders.Add , , "barcode", 140
.ColumnHeaders.Add , , " «Œ—  ÕœÌÀ ··»Ì«‰«  ", 140
.ColumnHeaders.Add , , " «·«” ﬁÿ«⁄ «·⁄«„ ", 140

.ColumnHeaders.Add , , " «·«” —Ã«⁄ «·⁄«„ ", 140
.ColumnHeaders.Add , , "«·«” ﬁÿ«⁄ ‰”»…", 140
.ColumnHeaders.Add , , "«·«” —Ã«⁄ ‰”»…", 120
.ColumnHeaders.Add , , "”»» «·«” ﬁÿ«⁄", 0
.ColumnHeaders.Add , , "”»» «·«÷«›…", 0
.ColumnHeaders.Add , , "«·„·«Õ÷«  «·⁄«„…", 120
.ColumnHeaders.Add , , " „ﬁœ«— «·“Ì«œÂ «·⁄·«ÊÂ «Ê «· —›Ì⁄ ", 140
.ColumnHeaders.Add , , "«·„·«ÕŸ‹‹‹‹‹‹‹‹‹‹«  ", 140
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

last1 = Sheets(13).Range("b" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(13).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(13).Cells(frw1, "b")
item1.SubItems(2) = Sheets(13).Cells(frw1, "c")
item1.SubItems(3) = Sheets(13).Cells(frw1, "d")
item1.SubItems(4) = Sheets(13).Cells(frw1, "e")
item1.SubItems(5) = Sheets(13).Cells(frw1, "f")
item1.SubItems(6) = Sheets(13).Cells(frw1, "g")
item1.SubItems(7) = Sheets(13).Cells(frw1, "h")
item1.SubItems(8) = Sheets(13).Cells(frw1, "i")
item1.SubItems(9) = Sheets(13).Cells(frw1, "j")
item1.SubItems(10) = Sheets(13).Cells(frw1, "k")
item1.SubItems(11) = Sheets(13).Cells(frw1, "l")
item1.SubItems(12) = Sheets(13).Cells(frw1, "m")
item1.SubItems(13) = Sheets(13).Cells(frw1, "n")
item1.SubItems(14) = Sheets(13).Cells(frw1, "o")
item1.SubItems(15) = Sheets(13).Cells(frw1, "p")
item1.SubItems(16) = Sheets(13).Cells(frw1, "q")
item1.SubItems(17) = Sheets(13).Cells(frw1, "r")
item1.SubItems(18) = Sheets(13).Cells(frw1, "s")
item1.SubItems(19) = Sheets(13).Cells(frw1, "t")
item1.SubItems(20) = Sheets(13).Cells(frw1, "u")
item1.SubItems(21) = Sheets(13).Cells(frw1, "v")
item1.SubItems(22) = Sheets(13).Cells(frw1, "w")
item1.SubItems(23) = Sheets(13).Cells(frw1, "x")
item1.SubItems(24) = Sheets(13).Cells(frw1, "y")
item1.SubItems(25) = Sheets(13).Cells(frw1, "z")
item1.SubItems(26) = Sheets(13).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(13).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(13).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(13).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(13).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(13).Cells(frw1, "af")
item1.SubItems(32) = Sheets(13).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(13).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(13).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(13).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(13).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(13).Cells(frw1, "al")
item1.SubItems(38) = Sheets(13).Cells(frw1, "am")
item1.SubItems(39) = Sheets(13).Cells(frw1, "an")
item1.SubItems(40) = Sheets(13).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(13).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(13).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(13).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(13).Cells(frw1, "as")
item1.SubItems(45) = Sheets(13).Cells(frw1, "at")
item1.SubItems(46) = Sheets(13).Cells(frw1, "au")
item1.SubItems(47) = Sheets(13).Cells(frw1, "av")
item1.SubItems(48) = Sheets(13).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(13).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(13).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(13).Cells(frw1, "az")
item1.SubItems(52) = Sheets(13).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(13).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(13).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(13).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(13).Cells(frw1, "be")
item1.SubItems(57) = Sheets(13).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(13).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(13).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(13).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(13).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(13).Cells(frw1, "bk")
item1.SubItems(63) = Sheets(13).Cells(frw1, "bl")
item1.SubItems(64) = Sheets(13).Cells(frw1, "bm")
item1.SubItems(65) = Sheets(13).Cells(frw1, "bn")
item1.SubItems(66) = Sheets(13).Cells(frw1, "bo")
item1.SubItems(67) = Sheets(13).Cells(frw1, "bp")
item1.SubItems(68) = Sheets(13).Cells(frw1, "bq")
item1.SubItems(69) = Sheets(13).Cells(frw1, "br")
item1.SubItems(70) = Sheets(13).Cells(frw1, "bs")
item1.SubItems(71) = Sheets(13).Cells(frw1, "bt")
item1.SubItems(72) = Sheets(13).Cells(frw1, "bu")
item1.SubItems(73) = Sheets(13).Cells(frw1, "bv")


   


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



ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
 
ComboBox2.List = Sheets(16).Range("a69:a108").value



End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
With Me.ListView1
.ListItems.Clear
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

last1 = Sheets(13).Range("b" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(13).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(13).Cells(frw1, "b")
item1.SubItems(2) = Sheets(13).Cells(frw1, "c")
item1.SubItems(3) = Sheets(13).Cells(frw1, "d")
item1.SubItems(4) = Sheets(13).Cells(frw1, "e")
item1.SubItems(5) = Sheets(13).Cells(frw1, "f")
item1.SubItems(6) = Sheets(13).Cells(frw1, "g")
item1.SubItems(7) = Sheets(13).Cells(frw1, "h")
item1.SubItems(8) = Sheets(13).Cells(frw1, "i")
item1.SubItems(9) = Sheets(13).Cells(frw1, "j")
item1.SubItems(10) = Sheets(13).Cells(frw1, "k")
item1.SubItems(11) = Sheets(13).Cells(frw1, "l")
item1.SubItems(12) = Sheets(13).Cells(frw1, "m")
item1.SubItems(13) = Sheets(13).Cells(frw1, "n")
item1.SubItems(14) = Sheets(13).Cells(frw1, "o")
item1.SubItems(15) = Sheets(13).Cells(frw1, "p")
item1.SubItems(16) = Sheets(13).Cells(frw1, "q")
item1.SubItems(17) = Sheets(13).Cells(frw1, "r")
item1.SubItems(18) = Sheets(13).Cells(frw1, "s")
item1.SubItems(19) = Sheets(13).Cells(frw1, "t")
item1.SubItems(20) = Sheets(13).Cells(frw1, "u")
item1.SubItems(21) = Sheets(13).Cells(frw1, "v")
item1.SubItems(22) = Sheets(13).Cells(frw1, "w")
item1.SubItems(23) = Sheets(13).Cells(frw1, "x")
item1.SubItems(24) = Sheets(13).Cells(frw1, "y")
item1.SubItems(25) = Sheets(13).Cells(frw1, "z")
item1.SubItems(26) = Sheets(13).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(13).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(13).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(13).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(13).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(13).Cells(frw1, "af")
item1.SubItems(32) = Sheets(13).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(13).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(13).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(13).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(13).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(13).Cells(frw1, "al")
item1.SubItems(38) = Sheets(13).Cells(frw1, "am")
item1.SubItems(39) = Sheets(13).Cells(frw1, "an")
item1.SubItems(40) = Sheets(13).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(13).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(13).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(13).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(13).Cells(frw1, "as")
item1.SubItems(45) = Sheets(13).Cells(frw1, "at")
item1.SubItems(46) = Sheets(13).Cells(frw1, "au")
item1.SubItems(47) = Sheets(13).Cells(frw1, "av")
item1.SubItems(48) = Sheets(13).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(13).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(13).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(13).Cells(frw1, "az")
item1.SubItems(52) = Sheets(13).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(13).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(13).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(13).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(13).Cells(frw1, "be")
item1.SubItems(57) = Sheets(13).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(13).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(13).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(13).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(13).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(13).Cells(frw1, "bk")
item1.SubItems(63) = Sheets(13).Cells(frw1, "bl")
item1.SubItems(64) = Sheets(13).Cells(frw1, "bm")
item1.SubItems(65) = Sheets(13).Cells(frw1, "bn")
item1.SubItems(66) = Sheets(13).Cells(frw1, "bo")
item1.SubItems(67) = Sheets(13).Cells(frw1, "bp")
item1.SubItems(68) = Sheets(13).Cells(frw1, "bq")
item1.SubItems(69) = Sheets(13).Cells(frw1, "br")
item1.SubItems(70) = Sheets(13).Cells(frw1, "bs")
item1.SubItems(71) = Sheets(13).Cells(frw1, "bt")
item1.SubItems(72) = Sheets(13).Cells(frw1, "bu")
item1.SubItems(73) = Sheets(13).Cells(frw1, "bv")


   


Next frw1

UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
End With
End With

With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
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


For i = 1 To 74
Sheets(13).Cells(8, "a").value = "«· ”·”·"

Me.ComboBox1.AddItem Sheets(13).Cells(8, i).value

Next i

ComboBox8.RowSource = "col_ltr"
' ComboBox6.List = Array("sheet14", "")


'Dim lastr As Integer

'lastr = Sheets(29).Cells(Rows.Count, 1).End(xlUp).Row
 

'=========================
emad:
End Sub

