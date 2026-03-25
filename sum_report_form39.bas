Attribute VB_Name = "sum_report_form39"
Sub sum_report_44_1()

On Error Resume Next


Dim lastr13 As Integer
lastr13 = Sheets(44).Cells(Rows.count, "a").End(xlUp).row
Sheets(44).Cells(lastr13 + 1, "a") = "«·„Ã„Ê⁄ «·ﬂ·Ì"
Sheets(44).Cells(lastr13 + 1, "b") = Application.WorksheetFunction.Sum(Sheets(44).Range("b9 : b" & lastr13))
Sheets(44).Cells(lastr13 + 1, "c") = Application.WorksheetFunction.Sum(Sheets(44).Range("c9 : c" & lastr13))
Sheets(44).Cells(lastr13 + 1, "d") = Application.WorksheetFunction.Sum(Sheets(44).Range("d9 : d" & lastr13))
Sheets(44).Cells(lastr13 + 1, "e") = Application.WorksheetFunction.Sum(Sheets(44).Range("e9 : e" & lastr13))
Sheets(44).Cells(lastr13 + 1, "f") = Application.WorksheetFunction.Sum(Sheets(44).Range("f9 : f" & lastr13))
Sheets(44).Cells(lastr13 + 1, "g") = Application.WorksheetFunction.Sum(Sheets(44).Range("g9 : g" & lastr13))
Sheets(44).Cells(lastr13 + 1, "h") = Application.WorksheetFunction.Sum(Sheets(44).Range("h9 : h" & lastr13))
Sheets(44).Cells(lastr13 + 1, "i") = Application.WorksheetFunction.Sum(Sheets(44).Range("i9 : i" & lastr13))
Sheets(44).Cells(lastr13 + 1, "j") = Application.WorksheetFunction.Sum(Sheets(44).Range("j9 : j" & lastr13))
Sheets(44).Cells(lastr13 + 1, "k") = Application.WorksheetFunction.Sum(Sheets(44).Range("k9 : k" & lastr13))
Sheets(44).Cells(lastr13 + 1, "l") = Application.WorksheetFunction.Sum(Sheets(44).Range("l9 : l" & lastr13))
Sheets(44).Cells(lastr13 + 1, "m") = Application.WorksheetFunction.Sum(Sheets(44).Range("m9 : m" & lastr13))



Sheets(44).Cells(lastr13 + 1, "a").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "b").Font.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "c").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "d").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "e").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "f").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "g").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "h").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "i").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "j").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "k").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "l").Interior.Color = vbRed
Sheets(44).Cells(lastr13 + 1, "m").Interior.Color = vbRed



Sheets(44).Cells(lastr13 + 1, "a").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "b").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "c").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "d").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "e").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "f").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "g").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "h").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "i").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "j").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "k").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "l").Font.Color = vbYellow
Sheets(44).Cells(lastr13 + 1, "m").Font.Color = vbYellow

End Sub
