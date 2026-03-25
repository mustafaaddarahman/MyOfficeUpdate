Attribute VB_Name = "Module2"
Sub emad1()
On Error GoTo emad
Application.ThisWorkbook.Save
ThisWorkbook.SaveCopyAs fileName:="d:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
ThisWorkbook.SaveCopyAs fileName:="e:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
emad:
End Sub


