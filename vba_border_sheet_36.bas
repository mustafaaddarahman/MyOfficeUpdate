Attribute VB_Name = "vba_border_sheet_36"
Sub vba_border_sheet_36_1()
Dim iRange As Range
Dim iCells As Range
Dim lastR As Integer
lastR = Sheets(35).Cells(Rows.count, "a").End(xlUp).row + 50
Set iRange = Sheets(36).Range("a9:cd" & lastR)
For Each iCells In iRange
     If Not IsEmpty(iCells) Or IsEmpty(iCells) Then
    iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    End If
Next iCells

Sheets(36).Range("a9:cd1000").Interior.Color = &HFFFFFF
Sheets(36).Range("a9:cd1000").Font.Color = vbBlack





End Sub
