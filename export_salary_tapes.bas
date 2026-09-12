Attribute VB_Name = "export_salary_tapes"

Sub ExportSalaryTapesToSpecificFolder()
    Dim wsTapes As Worksheet
    Dim lastRow As Long
    Dim folderPath As String, fileName As String, fullPath As String
    Dim monthName As String
    Dim currentWorkbook As Workbook
    
    ' 1. «·≈⁄œ«œ«  «·√”«”Ì…
    Set currentWorkbook = ThisWorkbook
    Set wsTapes = currentWorkbook.Sheets("Salary_tapes")
    folderPath = "D:\EXPORT1\"
    
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    On Error Resume Next
    lastRow = wsTapes.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    On Error GoTo 0
    
    If lastRow = 0 Then
        MsgBox "«·‘Ì  ›«—€° ·« ÌÊÃœ »Ì«‰«  · ’œÌ—Â«!", vbExclamation
        Exit Sub
    End If
    
    monthName = Trim(wsTapes.Range("G1").Text)
    fileName = "«‘—ÿ… «·—Ê« » ·‘Â— " & monthName & ".xlsx"
    fullPath = folderPath & fileName

    ' 2.  ›⁄Ì· Ê÷⁄ «·Œ›«¡ «· «„ Ê„‰⁄ ≈‰‘«¡ ‰Ê«›– ÃœÌœ…
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- «·Õ· «·Ã–—Ì Â‰« ---
    ' ‰”Œ «·‘Ì  »«·ﬂ«„· ≈·Ï „’‰› ÃœÌœ (Â–Â «·ÿ—Ìﬁ… √”—⁄ Ê·«   —ﬂ ‰Ê«›– ⁄«·ﬁ… ⁄«œ…)
    wsTapes.Copy
    
    ' «·¬‰ «·„’‰› «·ÃœÌœ ÂÊ ActiveWorkbook
    With ActiveWorkbook
        '  ÕÊÌ· «·„⁄«œ·«  ≈·Ï ﬁÌ„ («Œ Ì«—Ì ·÷„«‰ À»«  «·»Ì«‰« )
        .Sheets(1).UsedRange.value = .Sheets(1).UsedRange.value
        
        ' Õ›Ÿ «·„·›
        .SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
        
        ' ≈€·«ﬁ «·„’‰› «·ÃœÌœ  „«„«
        .Close SaveChanges:=False
    End With

    ' 3. «· √ﬂœ „‰ Õ«·… «· ÿ»Ìﬁ «·—∆Ì”Ì («·»—‰«„Ã «·Œ«’ »ﬂ)
    ' ≈–« ﬂ«‰ «·ÌÊ“— ›Ê—„ Ì⁄„·° ‰÷„‰ »ﬁ«¡ «·≈ﬂ”· „Œ›Ì«  „«„«
    Application.Visible = False
    
    Application.CutCopyMode = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox " „  ’œÌ— «·»Ì«‰«  »‰Ã«Õ ≈·Ï: " & fullPath, vbInformation
End Sub
