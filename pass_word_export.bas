Attribute VB_Name = "pass_word_export"
Sub pass_word_export()
Dim PASS As Variant
Dim lastR, i As Integer
lastR = Sheets(45).Cells(Rows.count, 2).End(xlUp).row
For i = 1 To lastR
PASS = Sheets(45).Cells(i, 2).value
MsgBox PASS
Next i

End Sub
