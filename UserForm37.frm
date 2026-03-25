VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm37 
   Caption         =   "UserForm37"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7020
   OleObjectBlob   =   "UserForm37.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'If UserForm28.TextBox2.Text <> "" Or UserForm28.TextBox27.Text <> "" Then
'Call marge_down_form28_1
'Call export_change_month1

'Else



Call Data_appearance2
'Call vba_border_sheet_28_1


Call sum_report_28_1
Call marge_down_form28_1
Call export_change_month1
'Call sum_report_28_1


MsgBox "ÊãÊ ÚãáíÉÊÕÏíÑ ÇáÈíÇäÇÊ Çáì ÊŞÇÑíÑ  ÈÕíÛÉ ÇßÓá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå"
Call export_pdf_form_28_1
'End If
End Sub

Private Sub UserForm_Activate()
ComboBox1.value = Sheets(4).Cells(1, "k").value
ComboBox2.List = Array("ãŞÑ ÇáåíÃÉ", "ãŞÑ ÔÑßÉ ÇáÕäÇÚÇÊ ÇáÍÑÈíå", "ãÕäÚ ÇáßÑÇãÉ æÇáÍÇÑË", "ãÕäÚ ÇáÑÈíÚ", "ãÕäÚ ÇáäåÑæÇä", "ãÕäÚ ÍãæÑÇÈí", "ãÕäÚ ÇáíÑãæß", "ãÕäÚ ÚŞÈÉ æÈÏÑ", "ãÕäÚ ÇáŞÇÏÓíå", "ãÕäÚ ÇáÑÔíÏ", "ãÕäÚ ÍØíä", "ãÕäÚÌÇÈÑ Èä ÍíÇä", "ãÕäÚ ÊÈæß", "", "")
ComboBox3.List = Sheets(4).Range("j2:j14").value
ComboBox4.List = Array(Year(Date))

End Sub
