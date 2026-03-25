Attribute VB_Name = "Module19"
Sub open_file()
Dim file_name As Variant
Dim open_wk As Workbook
On Error GoTo 1
file_name = Application.GetOpenFilename(title:="", FileFilter:="excel files *.xlsm (*.xlsm),")
If file_name = False Then
1: MsgBox "„‰ ›÷·ﬂ ﬁ„ »«Œ Ì«— «·„·› «·–Ì ” ﬁÊ„ «” Ì—«œ «·»Ì«‰«  „‰Â", vbCritical, "—”«·…  ‰»ÌÂ"
Exit Sub
Else
Set open_wk = Workbooks.Open(file_name)
End If
End Sub
