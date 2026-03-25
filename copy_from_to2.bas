Attribute VB_Name = "copy_from_to2"
Sub copy_from_to_t002()
On Error GoTo emad




Application.DisplayAlerts = False







'===============================================================

Sheets(13).Range("a9:bz8000").ClearContents
Application.ScreenUpdating = False
Sheets(13).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select
Selection.Copy Sheets(13).Range("A1:bz8000")

Application.ScreenUpdating = True
[a8].Select
Sheets(13).Select
Sheets(13).Activate
Application.ScreenUpdating = True


emad:
End Sub
