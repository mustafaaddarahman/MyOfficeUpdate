Attribute VB_Name = "klassat_alrwtb"
Sub ExportSalarySummary_J27()
    Dim wsSource As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim folderPath As String, fileName As String, factoryTitle As String
    Dim timeStamp As String, footerRow As Long
    
    ' 1.  ÕœÌœ «·„’«œ—
    On Error Resume Next
    Set wsSource = Sheets("52")
    If wsSource Is Nothing Then Set wsSource = Sheets(52)
    On Error GoTo 0
    
    factoryTitle = Sheets("title_factory").Range("A2").value
    folderPath = "D:\export1\"
    
    ' ≈‰‘«¡ «·„Ã·œ ≈–« ·„ Ìﬂ‰ „ÊÃÊœ«
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    ' 2. «·ÿ«»⁄ «·“„‰Ì Ê«”„ «·„·›
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    fileName = "Œ·«’… «·—Ê« » ·‘Â— " & Format(Date, "mmmm") & " - " & factoryTitle & " - " & timeStamp & ".xlsx"
    
    ' 3. ≈‰‘«¡ «·„·› «·ÃœÌœ Ê«·‰”Œ («·„œÏ J27)
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    
    wsSource.Range("A1:J27").Copy
    With wsNew.Range("A1")
        .PasteSpecial Paste:=xlPasteAll
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With
    
    ' 4. ≈÷«›… «· –ÌÌ· ( —ﬂ ’›Ì‰ 28 Ê 29 Ê«·»œ¡ „‰ 30)
    footerRow = 30
    
    With wsNew.Range(wsNew.Cells(footerRow, 1), wsNew.Cells(footerRow + 1, 10))
        .Merge
        .value = "«·„‰Ÿ„             «·„Õ«”»             «·„œﬁﬁ             „œÌ— «· œﬁÌﬁ             „œÌ— «·„«·Ì…             „œÌ— «·„’‰⁄"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.Weight = xlThin
    End With
    
    ' 5. «·Õ›Ÿ Ê«·≈€·«ﬁ
    Application.DisplayAlerts = False
    wbNew.SaveAs fileName:=folderPath & fileName, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    MsgBox " „  ’œÌ— «·Œ·«’… («·„œÏ J27) »‰Ã«Õ.", vbInformation
End Sub
