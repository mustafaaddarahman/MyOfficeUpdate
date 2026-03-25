Attribute VB_Name = "copy_from_to53"
Sub copy_from_to53_1()
Sheets(53).Activate

Application.DisplayAlerts = False







'===============================================================

Sheets(53).Range("a1:ci10000").Delete
Application.ScreenUpdating = False
Sheets(53).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select
Selection.Copy Sheets(53).Range("A1:bz8000")

Application.ScreenUpdating = True
[a8].Select
Sheets(53).Select
Sheets(53).Activate
Application.ScreenUpdating = True
 Sheets(53).Range("a1:ci7").ClearContents
Sheets(53).Range("bw8:ci10000").EntireColumn.Hidden = True
End Sub
