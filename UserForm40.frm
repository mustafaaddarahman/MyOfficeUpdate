VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm40 
   Caption         =   "UserForm40"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7020
   OleObjectBlob   =   "UserForm40.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Sheets(44).Activate
Label11.Caption = ""



Call Data_appearance4
Call sum_report_44_1
Call marge_down_form39_1
Call expory_year_month_44_1
Call export_pdf_form_39_1

MsgBox "ÊãÊ ÚãáíÉÊÕÏíÑ ÇáÈíÇäÇÊ Çáì ÊÞÇÑíÑ  ÈÕíÛÉ ÇßÓá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå"


End Sub

Private Sub Label4_Click()

End Sub

Private Sub UserForm_Activate()
ComboBox1.List = Sheets(4).Range("q1: q2").value



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


