Attribute VB_Name = "vba_border_sheet_28"
Sub vba_border_sheet_28_1()
Dim iRange As Range
Dim iCells As Range
Dim lastR As Integer
lastR = Sheets(13).Cells(Rows.count, "b").End(xlUp).row + 50
Set iRange = Sheets(28).Range("a9:bz" & lastR)
For Each iCells In iRange
     If Not IsEmpty(iCells) Or IsEmpty(iCells) Then
    iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    End If
Next iCells

Sheets(28).Range("a9:bz10000").Interior.Color = &HFFFFFF
Sheets(28).Range("a9:bz10000").Font.Color = vbBlack


End Sub
