VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm26 
   Caption         =   "UserForm26"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10848
   OleObjectBlob   =   "UserForm26.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
Sheets("title_factory").Cells(2, 1) = ""
Sheets("title_factory").Cells(2, 2) = ""
Sheets("title_factory").Cells(2, 3) = ""
Sheets("title_factory").Cells(2, 4) = ""

Sheets("title_factory").Cells(2, 1) = ComboBox1.value
Sheets("title_factory").Cells(2, 2) = TextBox1.Text
Sheets("title_factory").Cells(2, 3) = TextBox2.Text
Sheets("title_factory").Cells(2, 4) = TextBox3.Text
MsgBox "ÊãÊ ÊåíÆÉ ÇáÈíÇäÇÊ æÇÕÈÍ ÇáäÙÇã ÌÇåÒ ááÚãá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå  "
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array("ãŞÑ ÇáåíÃÉ", "ãŞÑ ÔÑßÉ ÇáÕäÇÚÇÊ ÇáÍÑÈíå", "ãÕäÚ ÇáßÑÇãÉ æÇáÍÇÑË", "ãÕäÚ ÇáÑÈíÚ", "ãÕäÚ ÇáäåÑæÇä", "ãÕäÚ ÍãæÑÇÈí", "ãÕäÚ ÇáíÑãæß", "ãÕäÚ ÚŞÈÉ æÈÏÑ", "ãÕäÚ ÇáŞÇÏÓíå", "ãÕäÚ ÇáÑÔíÏ", "ãÕäÚ ÍØíä", "ãÕäÚ ÌÇÈÑ Èä ÍíÇä", "ãÕäÚ ÊÈæß", "", "")

End Sub
