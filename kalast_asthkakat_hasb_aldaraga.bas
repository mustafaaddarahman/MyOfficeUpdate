Attribute VB_Name = "kalast_asthkakat_hasb_aldaraga"
Sub ExportDataFinalV2()
    Dim wbSource As Workbook, wbNew As Workbook
    Dim ws26 As Worksheet, ws31 As Worksheet, wsNew As Worksheet
    Dim currentMonthName As String
    Dim folderPath As String, fileName As String, entityName As String, fullHeader As String
    Dim NextRow As Long, i As Integer, startCol As Integer
    Dim signatures As Variant

    ' 1. ≈⁄œ«œ«  «·„”«— «·ÃœÌœ ›Ì «·»«— ‘‰ D
    folderPath = "D:\export1\"
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir "D:\export1"
        On Error GoTo 0
    End If

    Set wbSource = ThisWorkbook
    Set ws26 = wbSource.Sheets(26)
    Set ws31 = wbSource.Sheets(31)
    
    entityName = ws31.Range("A2").value
    currentMonthName = Format(Date, "mmmm")
    
    '  ﬂÊÌ‰ ‰’ «· —ÊÌ”… ·«” Œœ«„Â ›Ì «·„·› Ê«”„ «·Õ›Ÿ
    fullHeader = "Œ·«’… «·«” Õﬁ«ﬁ«  Õ”» «·œ—Ã… ·‘Â— " & currentMonthName & " " & entityName

    ' 2. ≈‰‘«¡ «·„’‰› «·ÃœÌœ
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.DisplayRightToLeft = True

    ' 3. «· —ÊÌ”… «·⁄·ÊÌ… (A1:T2)
    With wsNew.Range("A1:T2")
        .Merge
        .value = fullHeader
        .Font.Size = 28
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' 4. ‰”Œ «·»Ì«‰«  ﬂﬁÌ„ Ê ‰”Ìﬁ«  (·÷„«‰ «·”—⁄… Ê⁄œ„ ÊÃÊœ √Œÿ«¡ œÊ«·)
    ws26.Range("A2:T19").Copy
    With wsNew.Range("A3")
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With

    ' 5. ≈÷«›… «· ÊﬁÌ⁄« 
    NextRow = 22
    signatures = Array("«·„‰Ÿ„", "«·„Õ«”»", "«·„œﬁﬁ", "„œÌ— «· œﬁÌﬁ", "«·„œÌ— «·„«·Ì", "„œÌ— «·„’‰⁄")
    startCol = 1

    For i = LBound(signatures) To UBound(signatures)
        With wsNew.Range(wsNew.Cells(NextRow, startCol), wsNew.Cells(NextRow + 1, startCol + 2))
            .Merge
            .value = signatures(i)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Borders.LineStyle = xlContinuous
        End With
        startCol = startCol + 3
    Next i

    ' 6. «·Õ›Ÿ »«·«”„ «·„” „œ „‰ «· —ÊÌ”…
    wsNew.Columns("A:T").AutoFit
    Application.CutCopyMode = False
    Application.DisplayAlerts = False

    ' «”„ «·„·› „‘ ﬁ „‰ «· —ÊÌ”… „⁄ ≈÷«›… «·Êﬁ  ·„‰⁄  ﬂ—«— «·«”„
    fileName = fullHeader & "_" & Format(Now, "hhmm") & ".xlsx"
    
    ' «·Õ›Ÿ ›Ì «·„”«— D:\export1
    wbNew.SaveAs fileName:=folderPath & fileName
    wbNew.Close SaveChanges:=False
    
    Application.DisplayAlerts = True
    MsgBox " „ «· ’œÌ— »‰Ã«Õ ≈·Ï: " & vbCrLf & folderPath & fileName, vbInformation
End Sub
