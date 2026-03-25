Attribute VB_Name = "copy_sheet35_tosheet50"
Sub copy_sheet35_tosheet50_1()
On Error GoTo emad
'===============================================================
Application.DisplayAlerts = False
Sheets(50).Range("a9:cf8000").ClearContents
Sheets(35).Range("a9:cb8000").ClearContents
Sheets(35).Range("cd9:cf8000").ClearContents
Sheets(36).Range("b9:cf8000").ClearContents



Application.ScreenUpdating = False
Sheets(35).Select
Range("a8").CurrentRegion.Select
Selection.Copy Sheets(50).Range("A7:cf8000")
Application.ScreenUpdating = True
[a8].Select
Sheets(50).Select
Sheets(50).Activate
Application.ScreenUpdating = True
emad:
End Sub
