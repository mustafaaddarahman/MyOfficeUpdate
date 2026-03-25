Attribute VB_Name = "Module14"
Sub ImportData()
On Error Resume Next
Application.ScreenUpdating = False

Dim sourceworkbook As Workbook
Dim targetworkbook As Workbook

'Set sourceworkbook = Workbooks.Open("& Sheets(4).Cells(1, "e") &")
Set sourceworkbook = Workbooks.Open(Sheets(4).Cells(1, "e"))

 Set targetworkbook = ThisWorkbook
 
  sourceworkbook.Sheets("sheet31").UsedRange.Copy
 'targetworkbook.Sheets("sheet25").Range("a1:bm1000").ClearContents
  targetworkbook.Sheets("sheet25").Range("a1").PasteSpecial (xlPasteValuesAndNumberFormats)
  
  Application.CutCopyMode = fals
  
  sourceworkbook.Close SaveChanges:=False
  
Application.ScreenUpdating = True
End Sub
