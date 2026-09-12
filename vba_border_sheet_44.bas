Attribute VB_Name = "vba_border_sheet_44"
Sub vba_border_sheet_44_1()
Dim iRange As Range
Dim iCells As Range
Dim lastR As Integer
lastR = Sheets(43).Cells(Rows.count, "a").End(xlUp).row + 50
Set iRange = Sheets(44).Range("a9:q" & lastR)
For Each iCells In iRange
    If Not IsEmpty(iCells) Or IsEmpty(iCells) Then
    iCells.BorderAround _
            LineStyle:=xlContinuous, _
            Weight:=xlThin
    End If
Next iCells
Sheets(44).Range("a9:m100").Interior.Color = &HFFFFFF
Sheets(44).Range("a9:m100").Font.Color = vbBlack


End Sub

