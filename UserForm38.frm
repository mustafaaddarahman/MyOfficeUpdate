VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm38 
   Caption         =   "UserForm38"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7020
   OleObjectBlob   =   "UserForm38.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'If UserForm32.TextBox2.Text <> "" Or UserForm32.TextBox27.Text <> "" Or UserForm28.TextBox32.Text <> "" Then
'Call delete_amployee_name
'Call delete_amployee_date
'Call delete_amployee_month

'Call marge_down_form32_1
'Call export_delete_leave1
'Else

Label11.Caption = ""

Call Data_appearance3
Call start_time
Call marge_down_form32_1
Call export_delete_leave1
Call stop_time1
Label11.Caption = Format(Sheets(41).Range("c2").value, "ss")
MsgBox "ÊãÊ ÚãáíÉÊÕÏíÑ ÇáÈíÇäÇÊ Çáì ÊŞÇÑíÑ  ÈÕíÛÉ ÇßÓá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå"
Call export_pdf_form_32_1

End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Activate()
ComboBox1.value = Sheets(4).Cells(1, "l").value
ComboBox2.List = Array("ãŞÑ ÇáåíÃÉ", "ãŞÑ ÔÑßÉ ÇáÕäÇÚÇÊ ÇáÍÑÈíå", "ãÕäÚ ÇáßÑÇãÉ æÇáÍÇÑË", "ãÕäÚ ÇáÑÈíÚ", "ãÕäÚ ÇáäåÑæÇä", "ãÕäÚ ÍãæÑÇÈí", "ãÕäÚ ÇáíÑãæß", "ãÕäÚ ÚŞÈÉ æÈÏÑ", "ãÕäÚ ÇáŞÇÏÓíå", "ãÕäÚ ÇáÑÔíÏ", "ãÕäÚ ÍØíä", "ãÕäÚÌÇÈÑ Èä ÍíÇä", "ãÕäÚ ÊÈæß", "", "")
ComboBox3.List = Sheets(4).Range("j2:j14").value
ComboBox4.List = Array(Year(Date))

End Sub

