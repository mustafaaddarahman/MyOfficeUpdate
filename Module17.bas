Attribute VB_Name = "Module17"
Sub Hideit()
On Error Resume Next
Application.ScreenUpdating = False
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
Application.DisplayFormulaBar = False
Application.DisplayStatusBar = Not Application.DisplayStatusBar
ActiveWindow.DisplayWorkbookTabs = False
Application.ScreenUpdating = True
End Sub
