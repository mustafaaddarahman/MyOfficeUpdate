Attribute VB_Name = "Module18"
Sub Showit()
Application.ScreenUpdating = False
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Application.DisplayFormulaBar = True
Application.DisplayStatusBar = True
ActiveWindow.DisplayWorkbookTabs = True
Application.ScreenUpdating = True
End Sub
