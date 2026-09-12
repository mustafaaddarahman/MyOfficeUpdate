Attribute VB_Name = "Module7"
Sub search_part()

'=========================
On Error Resume Next
With UserForm3.ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True

UserForm3.ListView1.Visible = True
UserForm3.ListView2.Visible = False
'On Error Resume Next
UserForm3.ListView1.ListItems.Clear


UserForm3.ListView1.Font = Bold
UserForm3.ListView1.Font.Size = 16
UserForm3.ListView1.Font.Name = "PT hrading"

End With

'==========================================================================








'=================================================


With UserForm3.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
UserForm3.Font.Name = "Calibri"

With UserForm3
Dim item1 As ListItem
Dim last1, frw1 As Integer

last1 = Sheets(3).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last1
If Sheets(3).Cells(frw1, "f") Like "*" & UserForm3.TextBox6.Text & "*" Then
Set item1 = UserForm3.ListView1.ListItems.Add(, , Sheets(3).Cells(frw1, "A"))
UserForm3.ListView1.ListItems.Add.SubItems(1) = Sheets(3).Cells(frw1, "b")
UserForm3.ListView1.ListItems.Add.SubItems(2) = Sheets(3).Cells(frw1, "c")
UserForm3.ListView1.ListItems.Add.SubItems(3) = Sheets(3).Cells(frw1, "d")
UserForm3.ListView1.ListItems.Add.SubItems(4) = Sheets(3).Cells(frw1, "e")
UserForm3.ListView1.ListItems.Add.SubItems(5) = Sheets(3).Cells(frw1, "f")
UserForm3.ListView1.ListItems.Add.SubItems(6) = Sheets(3).Cells(frw1, "g")
UserForm3.ListView1.ListItems.Add.SubItems(7) = Sheets(3).Cells(frw1, "h")
UserForm3.ListView1.ListItems.Add.SubItems(8) = Sheets(3).Cells(frw1, "i")
UserForm3.ListView1.ListItems.Add.SubItems(9) = Sheets(3).Cells(frw1, "j")
UserForm3.ListView1.ListItems.Add.SubItems(10) = Sheets(3).Cells(frw1, "k")
UserForm3.ListView1.ListItems.Add.SubItems(11) = Sheets(3).Cells(frw1, "l")
UserForm3.ListView1.ListItems.Add.SubItems(12) = Sheets(3).Cells(frw1, "m")
UserForm3.ListView1.ListItems.Add.SubItems(13) = Sheets(3).Cells(frw1, "n")
UserForm3.ListView1.ListItems.Add.SubItems(14) = Sheets(3).Cells(frw1, "o")
UserForm3.ListView1.ListItems.Add.SubItems(15) = Sheets(3).Cells(frw1, "p")
UserForm3.ListView1.ListItems.Add.SubItems(16) = Sheets(3).Cells(frw1, "q")
UserForm3.ListView1.ListItems.Add.SubItems(17) = Sheets(3).Cells(frw1, "r")
UserForm3.ListView1.ListItems.Add.SubItems(18) = Sheets(3).Cells(frw1, "s")
UserForm3.ListView1.ListItems.Add.SubItems(19) = Sheets(3).Cells(frw1, "t")
UserForm3.ListView1.ListItems.Add.SubItems(20) = Sheets(3).Cells(frw1, "u")
UserForm3.ListView1.ListItems.Add.SubItems(21) = Sheets(3).Cells(frw1, "v")
UserForm3.ListView1.ListItems.Add.SubItems(22) = Sheets(3).Cells(frw1, "w")
UserForm3.ListView1.ListItems.Add.SubItems(23) = Sheets(3).Cells(frw1, "x")
UserForm3.ListView1.ListItems.Add.SubItems(24) = Sheets(3).Cells(frw1, "y")
UserForm3.ListView1.ListItems.Add.SubItems(25) = Sheets(3).Cells(frw1, "z")
UserForm3.ListView1.ListItems.Add.SubItems(26) = Sheets(3).Cells(frw1, "aa")
UserForm3.ListView1.ListItems.Add.SubItems(27) = Sheets(3).Cells(frw1, "ab")
UserForm3.ListView1.ListItems.Add.SubItems(28) = Sheets(3).Cells(frw1, "ac")
UserForm3.ListView1.ListItems.Add.SubItems(29) = Sheets(3).Cells(frw1, "ad")
UserForm3.ListView1.ListItems.Add.SubItems(30) = Sheets(3).Cells(frw1, "ae")
UserForm3.ListView1.ListItems.Add.SubItems(31) = Sheets(3).Cells(frw1, "af")
UserForm3.ListView1.ListItems.Add.SubItems(32) = Sheets(3).Cells(frw1, "ag")
UserForm3.ListView1.ListItems.Add.SubItems(33) = Sheets(3).Cells(frw1, "ah")
UserForm3.ListView1.ListItems.Add.SubItems(34) = Sheets(3).Cells(frw1, "ai")
UserForm3.ListView1.ListItems.Add.SubItems(35) = Sheets(3).Cells(frw1, "aj")
UserForm3.ListView1.ListItems.Add.SubItems(36) = Sheets(3).Cells(frw1, "ak")
UserForm3.ListView1.ListItems.Add.SubItems(37) = Sheets(3).Cells(frw1, "al")
UserForm3.ListView1.ListItems.Add.SubItems(38) = Sheets(3).Cells(frw1, "am")
UserForm3.ListView1.ListItems.Add.SubItems(39) = Sheets(3).Cells(frw1, "an")
UserForm3.ListView1.ListItems.Add.SubItems(40) = Sheets(3).Cells(frw1, "ao")
UserForm3.ListView1.ListItems.Add.SubItems(41) = Sheets(3).Cells(frw1, "ap")
UserForm3.ListView1.ListItems.Add.SubItems(42) = Sheets(3).Cells(frw1, "aq")
UserForm3.ListView1.ListItems.Add.SubItems(43) = Sheets(3).Cells(frw1, "ar")
UserForm3.ListView1.ListItems.Add.SubItems(44) = Sheets(3).Cells(frw1, "as")
UserForm3.ListView1.ListItems.Add.SubItems(45) = Sheets(3).Cells(frw1, "at")
UserForm3.ListView1.ListItems.Add.SubItems(46) = Sheets(3).Cells(frw1, "au")
UserForm3.ListView1.ListItems.Add.SubItems(47) = Sheets(3).Cells(frw1, "av")
UserForm3.ListView1.ListItems.Add.SubItems(48) = Sheets(3).Cells(frw1, "aw")
UserForm3.ListView1.ListItems.Add.SubItems(49) = Sheets(3).Cells(frw1, "ax")
UserForm3.ListView1.ListItems.Add.SubItems(50) = Sheets(3).Cells(frw1, "ay")
UserForm3.ListView1.ListItems.Add.SubItems(51) = Sheets(3).Cells(frw1, "az")
UserForm3.ListView1.ListItems.Add.SubItems(52) = Sheets(3).Cells(frw1, "ba")
UserForm3.ListView1.ListItems.Add.SubItems(53) = Sheets(3).Cells(frw1, "bb")
UserForm3.ListView1.ListItems.Add.SubItems(54) = Sheets(3).Cells(frw1, "bc")
UserForm3.ListView1.ListItems.Add.SubItems(55) = Sheets(3).Cells(frw1, "bd")
UserForm3.ListView1.ListItems.Add.SubItems(56) = Sheets(3).Cells(frw1, "be")
UserForm3.ListView1.ListItems.Add.SubItems(57) = Sheets(3).Cells(frw1, "bf")
End If
 
Next frw1

Dim Columns As Long
Dim i As Long
With UserForm3.ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next





'=============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(58).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(59).Alignment = lvwColumnCenter
    
   














'=================================================




End With

End With


ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
'ListView1.ForeColor = "blue"
End With
End Sub
