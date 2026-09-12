VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm46 
   Caption         =   "UserForm46"
   ClientHeight    =   12300
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7992
   OleObjectBlob   =   "UserForm46.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm46"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton27_Click()
Unload Me
UserForm10.Show
End Sub

Private Sub CommandButton28_Click()

Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
 Sheets(52).Cells(6, "a").value = Int(TextBox1.Text)
Sheets(52).Cells(6, "c").value = Int(TextBox2.Text)
Sheets(52).Cells(6, "e").value = Int(TextBox3.Text)
Sheets(52).Cells(6, "g").value = Int(TextBox4.Text)
Sheets(52).Cells(6, "i").value = Int(TextBox5.Text)
 Sheets(52).Cells(8, "a").value = Int(TextBox6.Text)
Sheets(52).Cells(8, "c").value = Int(TextBox7.Text)
Sheets(52).Cells(8, "e").value = Int(TextBox8.Text)
Sheets(52).Cells(8, "g").value = Int(TextBox9.Text)
Sheets(52).Cells(8, "i").value = Int(TextBox10.Text)
 Sheets(52).Cells(10, "a").value = Int(TextBox11.Text)
Sheets(52).Cells(10, "c").value = Int(TextBox12.Text)
Sheets(52).Cells(10, "e").value = Int(TextBox13.Text)
Sheets(52).Cells(10, "g").value = Int(TextBox15.Text)
Sheets(52).Cells(13, "e").value = Int(TextBox14.Text)


Sheets(52).Cells(23, "b").value = Int(TextBox16.Text)
Sheets(52).Cells(24, "b").value = Int(TextBox17.Text)
Sheets(52).Cells(25, "b").value = Int(TextBox18.Text)
Sheets(52).Cells(26, "b").value = Int(TextBox20.Text)
 Sheets(52).Cells(27, "b").value = Int(TextBox19.Text)
Sheets(52).Cells(23, "e").value = Int(TextBox21.Text)
Sheets(52).Cells(24, "e").value = Int(TextBox22.Text)
Sheets(52).Cells(25, "e").value = Int(TextBox23.Text)
Sheets(52).Cells(26, "e").value = Int(TextBox24.Text)
Sheets(52).Cells(22, "h").value = Int(TextBox25.Text)
Sheets(52).Cells(23, "g").value = Int(TextBox26.Text)
Sheets(52).Cells(23, "j").value = Int(TextBox27.Text)
Sheets(52).Cells(24, "g").value = Int(TextBox29.Text)
Sheets(52).Cells(24, "j").value = Int(TextBox28.Text)
Call ExportSalarySummary_J27







Sheets(52).Range("a1:j30").PrintPreview
ThisWorkbook.Application.Visible = False
UserForm46.Show

End Sub

Private Sub Label22_Click()

End Sub

Private Sub Label33_Click()

End Sub

Private Sub TextBox27_Change()

End Sub

Private Sub UserForm_Activate()

TextBox1.Text = Int(Sheets(1).Cells(3, "cs"))
TextBox2.Text = Int(Sheets(1).Cells(3, "ct"))
TextBox3.Text = Int(Sheets(1).Cells(3, "cu"))
   
TextBox4.Text = Int(Sheets(1).Cells(3, "cv"))
TextBox5.Text = Int(Sheets(1).Cells(3, "cw"))
TextBox6.Text = Int(Sheets(1).Cells(3, "cx"))
TextBox7.Text = Int(Sheets(1).Cells(3, "cy"))
TextBox8.Text = Int(Sheets(1).Cells(3, "cz"))
TextBox9.Text = Int(Sheets(1).Cells(5, "cs"))
TextBox10.Text = Int(Sheets(1).Cells(5, "ct"))
TextBox11.Text = Int(Sheets(1).Cells(5, "cu"))
TextBox12.Text = Int(Sheets(1).Cells(5, "cv"))
TextBox13.Text = Int(Sheets(1).Cells(5, "cw"))
TextBox15.Text = Int(Sheets(1).Cells(5, "cx"))
TextBox14.Text = Int(Sheets(1).Cells(2, "cf"))

'========================= «·⁄œœ «·ﬂ·Ì
TextBox16.Text = Int(Sheets(1).Cells(8, "cs"))

'================================«·⁄œœ «·ﬂ·Ì ··„” „—Ì‰
TextBox17.Text = Int(Sheets(1).Cells(8, "ct"))

'===================================«·⁄œœ «·ﬂ·Ì ··„Ã«“Ì‰
TextBox18.Text = Sheets(1).Cells(8, "cu")

'=====================================«·⁄œœ «·ﬂ·Ì ··„⁄Ì· «·„ ›—€
TextBox20.Text = Sheets(1).Cells(8, "cv")

'===================================«·⁄œœ «·ﬂ·Ì «··Ã«‰ «·ÿ»Ì…
TextBox19.Text = Sheets(1).Cells(8, "cw")

'==================================«·—« » «·«”„Ì ··„” „—Ì‰
TextBox21.Text = Int(Sheets(1).Cells(10, "cs"))

'====================================«·—« » «·«”„Ì ··„Ã«“Ì‰
TextBox22.Text = Int(Sheets(1).Cells(10, "ct"))

'===================================«·—« » «·«”„Ì ··„⁄Ì·
TextBox23.Text = Int(Sheets(1).Cells(10, "cu"))
'===================================«·—« » «·«”„Ì «··Ã«‰ «·ÿ»Ì…
TextBox24.Text = Int(Sheets(1).Cells(10, "cv"))
'===================================
TextBox25.Text = Int(Sheets(1).Cells(10, "cw"))

TextBox26.Text = Int(Sheets(1).Cells(8, "df"))

TextBox27.Text = Int(Sheets(1).Cells(10, "db"))
TextBox28.Text = Int(Sheets(1).Cells(10, "dc"))

TextBox29.Text = Int(Sheets(1).Cells(8, "dd"))


End Sub

Private Sub UserForm_Initialize()

End Sub
