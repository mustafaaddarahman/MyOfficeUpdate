Attribute VB_Name = "copy_from_sheet1_to_sheet54"
Sub copy_columns_paste()

Worksheets("data54").Range("a9:bc10000").ClearContents
Dim ir As Integer, erow As Integer, sh1 As Worksheet, sh2 As Worksheet, i As Long
Set sh1 = Sheets(1)
Set sh2 = Worksheets("data54")
lr = sh1.Cells(Rows.count, 2).End(xlUp).row
For i = 8 To lr

erow = sh2.Cells(Rows.count, 1).End(xlUp).Offset(1, 0).row
sh2.Cells(erow, 1) = sh1.Cells(i, "a")
sh2.Cells(erow, 2) = sh1.Cells(i, "b")
sh2.Cells(erow, 3) = sh1.Cells(i, "c")
sh2.Cells(erow, 4) = sh1.Cells(i, "d")
sh2.Cells(erow, 5) = sh1.Cells(i, "e")
sh2.Cells(erow, 6) = sh1.Cells(i, "f")
sh2.Cells(erow, 7) = sh1.Cells(i, "g")
sh2.Cells(erow, 8) = sh1.Cells(i, "h")
sh2.Cells(erow, 9) = sh1.Cells(i, "i")
sh2.Cells(erow, 10) = sh1.Cells(i, "j")
sh2.Cells(erow, 11) = sh1.Cells(i, "k")
sh2.Cells(erow, 12) = sh1.Cells(i, "l")
sh2.Cells(erow, 13) = sh1.Cells(i, "m")
sh2.Cells(erow, 14) = sh1.Cells(i, "n")
sh2.Cells(erow, 15) = sh1.Cells(i, "o")
sh2.Cells(erow, 16) = sh1.Cells(i, "p")
sh2.Cells(erow, 17) = sh1.Cells(i, "q")
sh2.Cells(erow, 18) = sh1.Cells(i, "r")
sh2.Cells(erow, 19) = sh1.Cells(i, "s")
sh2.Cells(erow, 20) = sh1.Cells(i, "t")
sh2.Cells(erow, 21) = sh1.Cells(i, "u")
sh2.Cells(erow, 22) = sh1.Cells(i, "v")
sh2.Cells(erow, 23) = sh1.Cells(i, "x")
sh2.Cells(erow, 24) = sh1.Cells(i, "y")
sh2.Cells(erow, 25) = sh1.Cells(i, "z")
sh2.Cells(erow, 26) = sh1.Cells(i, "aa")
sh2.Cells(erow, 27) = sh1.Cells(i, "ab")
sh2.Cells(erow, 28) = sh1.Cells(i, "ac")
sh2.Cells(erow, 29) = sh1.Cells(i, "ad")
sh2.Cells(erow, 30) = sh1.Cells(i, "ae")
sh2.Cells(erow, 31) = sh1.Cells(i, "af")
sh2.Cells(erow, 32) = sh1.Cells(i, "ag")
sh2.Cells(erow, 33) = sh1.Cells(i, "ah")
sh2.Cells(erow, 34) = sh1.Cells(i, "ai")
sh2.Cells(erow, 35) = sh1.Cells(i, "aj")
sh2.Cells(erow, 36) = sh1.Cells(i, "ak")
sh2.Cells(erow, 37) = sh1.Cells(i, "al")
sh2.Cells(erow, 38) = sh1.Cells(i, "am")
sh2.Cells(erow, 39) = sh1.Cells(i, "an")
sh2.Cells(erow, 40) = sh1.Cells(i, "ao")
sh2.Cells(erow, 41) = sh1.Cells(i, "ap")
sh2.Cells(erow, 42) = sh1.Cells(i, "aq")
sh2.Cells(erow, 43) = sh1.Cells(i, "ar")
sh2.Cells(erow, 44) = sh1.Cells(i, "as")
sh2.Cells(erow, 45) = sh1.Cells(i, "at")
sh2.Cells(erow, 46) = sh1.Cells(i, "au")
sh2.Cells(erow, 47) = sh1.Cells(i, "av")
sh2.Cells(erow, 48) = sh1.Cells(i, "aw")
sh2.Cells(erow, 49) = sh1.Cells(i, "ay")
sh2.Cells(erow, 50) = sh1.Cells(i, "az")
sh2.Cells(erow, 51) = sh1.Cells(i, "ba")
sh2.Cells(erow, 52) = sh1.Cells(i, "bd")

sh2.Cells(erow, 53) = sh1.Cells(i, "bi")
sh2.Cells(erow, 54) = sh1.Cells(i, "be")


Next
Sheets(54).Activate
MsgBox " „ ‰”Œ «·»Ì«‰«  Ê ÂÌ√ Â«", vbDefaultButton1, "·—”«·…  ÊÃÌÂ"



End Sub
