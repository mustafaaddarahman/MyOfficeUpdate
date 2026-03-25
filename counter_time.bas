Attribute VB_Name = "counter_time"
Sub start_time()
On Error Resume Next
Sheets(41).Range("a2").value = ""
Sheets(41).Range("b2").value = ""
Sheets(41).Range("c2").value = ""
Sheets(41).Range("a2").Find("").Select
ActiveCell.value = Time
Selection.NumberFormat = "mm:ss"
End Sub

