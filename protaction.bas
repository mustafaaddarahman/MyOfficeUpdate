Attribute VB_Name = "protaction"
Option Explicit

Sub open_workbok_1()
Dim X As String
X = (ThisWorkbook.path & "\" & ThisWorkbook.Name)
If Ê—ﬁ…2.Range("RAND1") = X Then
Application.Visible = False
If Sheets("title_factory").Cells(2, 1) = "" Or Sheets("title_factory").Cells(2, 2) = "" Or Sheets("title_factory").Cells(2, 3) = "" Or Sheets("title_factory").Cells(2, 4) = "" Then
UserForm25.Show
Else
UserForm20.Show
End If
Else
MsgBox "Õ«Ê· «·« ’«· »«·„»—„Ã ·ÌﬁÊ„ » ›⁄Ì· «·‰Ÿ«„ ⁄·Ï ÃÂ«“ﬂ  ·« Ì„ﬂ‰ ‰ﬁ· «Ê  €Ì— «”„ «·‰Ÿ«„ ", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
UserForm50.Show
End If

End Sub

Sub open_workbok_2()
Dim X As String
X = (ThisWorkbook.path & "\" & ThisWorkbook.Name)
Ê—ﬁ…2.Range("RAND1") = X
End Sub
