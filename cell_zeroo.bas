Attribute VB_Name = "cell_zeroo"
Sub celzero()
Dim rng As Range
Dim cel As Range
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, "b").End(xlUp).row
Set rng = Sheets(1).Range("a9:br" & lastR)
For Each cel In rng
If IsEmpty(cel) Then
cel = 0
End If
If IsDate(cel) Then
cel = ""
End If
Next cel
MsgBox " „  ’›Ì— «·Œ·«Ì« «·›«—€Â", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
End Sub
