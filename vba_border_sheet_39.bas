Attribute VB_Name = "vba_border_sheet_39"
Sub vba_border_sheet_39_1()
Dim iRange As Range
Dim iCells As Range
Dim lastR As Integer
lastR = Sheets(3).Cells(Rows.count, "a").End(xlUp).row + 50
Set iRange = Sheets(39).Range("a9:bw" & lastR)
For Each iCells In iRange
     If Not IsEmpty(iCells) Or IsEmpty(iCells) Then
    iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    End If
Next iCells

Sheets(39).Range("a9:bw1000").Interior.Color = &HFFFFFF
Sheets(39).Range("a9:bw1000").Font.Color = vbBlack








End Sub
