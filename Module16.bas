Attribute VB_Name = "Module16"
Sub getimporfilename()
Filt = "exel file(*.xlsm),(*.xlsm)"
filterendex = 2
Title = "select afile to import"
b = Application.GetOpenFilename _
(FileFilter:=Filt, FilterIndex:=filterindes, Title:=Title)
Sheets(4).Cells(1, "e") = b


End Sub
