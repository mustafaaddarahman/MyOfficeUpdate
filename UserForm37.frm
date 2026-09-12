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


MsgBox "ÊãÊ ÚãáíÉÊÕÏíÑ ÇáÈíÇäÇÊ Çáì ÊÞÇÑíÑ  ÈÕíÛÉ ÇßÓá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå"
Call export_pdf_form_28_1
'End If
End Sub

Private Sub UserForm_Activate()
ComboBox1.value = Sheets(4).Cells(1, "k").value



Dim lastRow As Long
Dim ws1 As Worksheet

' ÊÚííä ÇáæÑÞÉ ÇáÎÇãÓÉ ááãÊÛíÑ ws1
Set ws1 = Sheets(5)

' ÊÍÏíÏ ÑÞã ÂÎÑ ÕÝ íÍÊæí Úáì ÈíÇäÇÊ Ýí ÇáÚãæÏ B
' ÇáßæÏ íÈÏÃ ãä ÂÎÑ ÕÝ Ýí ÇáÅßÓíá æíÕÚÏ ááÃÚáì ÍÊì íÌÏ Ãæá ÎáíÉ ÈåÇ ÞíãÉ
lastRow = ws1.Cells(ws1.rowS.count, "B").End(xlUp).row

' ÊÚÈÆÉ ComboBox4 ÈÇáäØÇÞ ãä ÇáÎáíÉ B1 Åáì ÂÎÑ ÕÝ æÌÏäÇå
If lastRow >= 1 Then
    ComboBox2.List = ws1.Range("B1:B" & lastRow).value
End If


ComboBox3.List = Sheets(4).Range("j2:j14").value
ComboBox4.List = Array(Year(Date))

End Sub
