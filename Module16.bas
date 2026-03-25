Attribute VB_Name = "Module16"
Sub getimporfilename()
Filt = "exel file(*.xlsm),(*.xlsm)"
filterendex = 2
title = "select afile to import"
b = Application.GetOpenFilename _
(FileFilter:=Filt, FilterIndex:=filterindes, title:=title)
Sheets(4).Cells(1, "e") = b


End Sub
