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
MsgBox "ÊãÊ ÚãáíÉÊÕÏíÑ ÇáÈíÇäÇÊ Çáì ÊÞÇÑíÑ  ÈÕíÛÉ ÇßÓá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå"
Call export_pdf_form_32_1

End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Activate()
ComboBox1.value = Sheets(4).Cells(1, "l").value


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

