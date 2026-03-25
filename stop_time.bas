Attribute VB_Name = "stop_time"
Sub stop_time1()
On Error Resume Next

Sheets(41).Range("b2").Find("").Select
ActiveCell.value = Time
Selection.NumberFormat = "mm:ss"

Sheets(41).Range("c2").Find("").Select
ActiveCell.value = Format(Time, "long time")
ActiveCell.Formula2R1C1 = "=RC[-1]-RC[-2]"
Selection.NumberFormat = "mm:ss"

End Sub
