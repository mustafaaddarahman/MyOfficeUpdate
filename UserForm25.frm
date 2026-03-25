VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm25 
   Caption         =   "UserForm25"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   11040
   OleObjectBlob   =   "UserForm25.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton1_Click()

Sheets("title_factory").Cells(2, 1) = ComboBox1.value
Sheets("title_factory").Cells(2, 2) = TextBox1.Text
Sheets("title_factory").Cells(2, 3) = TextBox2.Text
Sheets("title_factory").Cells(2, 4) = Format(TextBox3.Text, "yyyy/mm/dd")
MsgBox "ÊãÊ ÊåíÆÉ ÇáÈíÇäÇÊ æÇÕÈÍ ÇáäÙÇã ÌÇåÒ ááÚãá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå  "


End Sub

Private Sub CommandButton2_Click()
If Sheets("title_factory").Cells(2, 1) = "" And Sheets("title_factory").Cells(2, 2) = "" And Sheets("title_factory").Cells(2, 3) = "" And Sheets("title_factory").Cells(2, 4) = "" Then
MsgBox "áã ÊŞã ÈÊåíÆÉ Çí ÈíÇäÇÊ ÇÓÇÓíå áãÚÑİÉ ÇÓã ÇáÊÔßíá Çæ ÇáãÕäÚ æãä åã ÇáãÓÊÎÏãíä áíÊÓäì ÇáÏÎæá Çáì ãäÕÉ ÇáäÙÇã", vbCritical, "ÑÓÇáÉ ÊäÈíå"
Unload Me
Application.Quit
Else
Unload Me
UserForm20.Show

End If
End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("ãŞÑ ÇáåíÃÉ", "ãŞÑ ÔÑßÉ ÇáÕäÇÚÇÊ ÇáÍÑÈíå", "ãÕäÚ ÇáßÑÇãÉ æÇáÍÇÑË", "ãÕäÚ ÇáÑÈíÚ", "ãÕäÚ ÇáäåÑæÇä", "ãÕäÚ ÍãæÑÇÈí", "ãÕäÚ ÇáíÑãæß", "ãÕäÚ ÚŞÈÉ æÈÏÑ", "ãÕäÚ ÇáŞÇÏÓíå", "ãÕäÚ ÇáÑÔíÏ", "ãÕäÚ ÍØíä", "ãÕäÚ ÌÇÈÑ Èä ÍíÇä", "ãÕäÚ ÊÈæß", "", "")
End Sub
