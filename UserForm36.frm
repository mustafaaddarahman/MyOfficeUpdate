VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm36 
   Caption         =   "UserForm36"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7020
   OleObjectBlob   =   "UserForm36.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Sheets(28).Activate
Label11.Caption = ""
TextBox28.Text = ""
TextBox28.Text = Now

If ComboBox1.value = " ﬁ—Ì—Õ”» «·œ—Ã…" Then
Call export_data_daraja1
Call export_pdf_draja_1
Else
Select Case ComboBox1.value
'====================================
Case " ﬁ—Ì— «·≈÷«›…"
 
Call Data_appearance1
Call sum_filter_form17_1
'Call vba_border_sheet_28_1
Call marge_down_form17_1
Call export_specificsheet
TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)

MsgBox " „  ⁄„·Ì… ’œÌ— «·»Ì«‰«  «·Ï  ﬁ«—Ì—  »’Ì€… «ﬂ”·", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
'====================================
Case " ﬁ—Ì— «·ﬁ«∆„… «·„›—œ…"

Call Data_appearance1
Call sum_filter_form17_1
'Call vba_border_sheet_28_1

Call marge_down_form17_1
Call export_specificsheet
TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)


MsgBox " „  ⁄„·Ì… ’œÌ— «·»Ì«‰«  «·Ï  ﬁ«—Ì—  »’Ì€… «ﬂ”·", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
'==============================================
'Case " ﬁ—Ì—Õ”» «·œ—Ã…"
' Call export_data_daraja1

'=============================================
Case " ﬁ—Ì— ’—›Ì… —Ê« »"

Call Data_appearance_hiden_1

Call sum_filter_form17_1
Call marge_down_form17_1
'Call vba_border_sheet_28_1
Call export_specificsheet

TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)

MsgBox " „  ⁄„·Ì… ’œÌ— «·»Ì«‰«  «·Ï  ﬁ«—Ì—  »’Ì€… «ﬂ”·", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
'===========================================
Case " ﬁ—Ì—„«” — "


Call Data_appearance1

'Call vba_border_sheet_28_1
Call marge_down_form17_1
Call export_specificsheet
TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)

MsgBox " „  ⁄„·Ì… ’œÌ— «·»Ì«‰«  «·Ï  ﬁ«—Ì—  »’Ì€… «ﬂ”·", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
'==================================================
Case " ﬁ—Ì— „Êﬁ› «· ’«—ÌÕ «·√„‰Ì…"

Call Data_appearance1

'Call vba_border_sheet_28_1
Call marge_down_form17_1
Call export_specificsheet
TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)

MsgBox " „  ⁄„·Ì… ’œÌ— «·»Ì«‰«  «·Ï  ﬁ«—Ì—  »’Ì€… «ﬂ”·", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
'===========================================   ’œÌ— ›Ê—„… «·»Ì«‰«  «·ﬂ·Ì…

Case " ’œÌ— ›Ê—„… «·»Ì«‰«  «·ﬂ·Ì…"
Call copy_from_to53_1
Call marge_down_form53_1
Call export_all_data_1
Call export_pdf_form_53_1
TextBox29.Text = ""
TextBox29.Text = Now


Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)


















'=========================================
Case Is <> ""
Call Data_appearance1
Call sum_filter_form17_1
'Call vba_border_sheet_28_1
Call marge_down_form17_1
Call export_specificsheet
'Call stop_time1
TextBox29.Text = ""
TextBox29.Text = Now

'Call stop_time1
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)

MsgBox " „  ⁄„·Ì… ’œÌ— «·»Ì«‰«  «·Ï  ﬁ«—Ì—  »’Ì€… «ﬂ”·", vbDefaultButton1, "—”«·…  ÊÃÌÂ"









End Select

Call export_pdf_form_17_1
End If
TextBox29.Text = ""
TextBox29.Text = Now
Label11.Caption = DateDiff("n", TextBox28.Text, TextBox29.Text) & ":" & DateDiff("s", TextBox28.Text, TextBox29.Text)
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub Label11_Click()

End Sub

Private Sub UserForm_Activate()
On Error Resume Next
ComboBox1.List = Sheets(4).Range("i2:i8").value
ComboBox2.List = Array("„ﬁ— «·ÂÌ√…", "„ﬁ— ‘—ﬂ… «·’‰«⁄«  «·Õ—»ÌÂ", "„’‰⁄ «·ﬂ—«„… Ê«·Õ«—À", "„’‰⁄ «·—»Ì⁄", "„’‰⁄ «·‰Â—Ê«‰", "„’‰⁄ Õ„Ê—«»Ì", "„’‰⁄ «·Ì—„Êﬂ", "„’‰⁄ ⁄ﬁ»… Ê»œ—", "„’‰⁄ «·ﬁ«œ”ÌÂ", "„’‰⁄ «·—‘Ìœ", "„’‰⁄ ÕÿÌ‰", "„’‰⁄Ã«»— »‰ ÕÌ«‰", "„’‰⁄  »Êﬂ", "", "")
ComboBox3.List = Sheets(4).Range("j2:j14").value
ComboBox4.List = Array(Year(Date))

End Sub


