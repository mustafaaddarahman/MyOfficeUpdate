Attribute VB_Name = "Module21"
Sub emad()
On Error GoTo emad
ThisWorkbook.SaveCopyAs fileName:="d:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
ThisWorkbook.SaveCopyAs fileName:="e:\emad\" & "- " & Format(Date, "ddd") & "  " & Format(Now, " HH  MM") & " - " & ThisWorkbook.Name
emad:
End Sub


