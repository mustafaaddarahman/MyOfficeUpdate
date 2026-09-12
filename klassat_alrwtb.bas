Attribute VB_Name = "klassat_alrwtb"
Sub ExportSalarySummary_J27()

  Dim wsSource As Worksheet, wsTarget As Worksheet, wsNew As Worksheet
    Dim wbNew As Workbook
    Dim lastRowSource As Long, i As Long, lastRowTarget As Long
    Dim reportText As String, agencyName As String
    Dim amountValue As Double
    Dim folderPath As String, fileName As String, factoryTitle As String
    Dim timeStamp As String, footerRow As Long
    Dim labelType As String
    
    ' «” Œœ«„ Dictionary ·Ã„⁄ «·„»«·€ Õ”» «·ÃÂ…
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    ' 1. ≈⁄œ«œ «·‘Ì « 
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets(1)
    Set wsTarget = ThisWorkbook.Sheets(52)
    On Error GoTo 0

    ' 2.  ‰ŸÌ› «·‘Ì  52 Ê›ﬂ «·œ„Ã
    wsTarget.Range("A35:J150").Clear
    wsTarget.Range("A35:J150").UnMerge

    lastRowSource = wsSource.Cells(wsSource.rowS.count, "E").End(xlUp).row
    
    ' 3. ŒÊ«—“„Ì… «· Ã„Ì⁄ (Aggregation)
    Dim colsAmt, colsAgn
    colsAmt = Array("AJ", "AN", "AR")
    colsAgn = Array("AK", "AO", "AS")
    
    For i = 9 To lastRowSource
        Dim j As Integer
        For j = 0 To 2
            amountValue = wsSource.Range(colsAmt(j) & i).value
            agencyName = Trim(wsSource.Range(colsAgn(j) & i).value)
            
            If amountValue <> 0 And agencyName <> "" Then
                ' ≈–« ﬂ«‰  «·ÃÂ… „ÊÃÊœ… „”»ﬁ«° √÷› «·„»·€ ··„Ã„Ê⁄ «·Õ«·Ì
                If dict.exists(agencyName) Then
                    dict(agencyName) = dict(agencyName) + amountValue
                Else
                    ' ≈–« ﬂ«‰  ÃÂ… ÃœÌœ…° «› Õ ·Â« ”Ã·« ÃœÌœ« »„»·€Â« «·Õ«·Ì
                    dict.Add agencyName, amountValue
                End If
            End If
        Next j
    Next i

    ' 4. »‰«¡ «·‰’ «·”—œÌ „‰ «·»Ì«‰«  «·„Ã„⁄…
    reportText = "Œ·«’… «·«” ﬁÿ«⁄«  Õ”» «·ÃÂ« : " & vbCrLf
    Dim key As Variant
    Dim count As Integer: count = 0
    
    For Each key In dict.Keys
        count = count + 1
        ' «· „ÌÌ“ »Ì‰ «·‰›ﬁ… Ê«·«” ﬁÿ«⁄ Õ”» «”„ «·ÃÂ…
        If InStr(key, "„œÌ—Ì…  ‰›Ì–") > 0 Then
            labelType = "‰›ﬁ… "
        Else
            labelType = "«” ﬁÿ«⁄ ÃÂ… "
        End If
        
        reportText = reportText & count & "- " & labelType & " (" & key & ") »„»·€ ≈Ã„«·Ì ﬁœ—Â: [" & Format(dict(key), "#,##0") & "] œ.⁄" & "  ∫  "
    Next key

    ' 5. ≈œ—«Ã «·‰’ ›Ì «·„‰ÿﬁ… «·„œ„Ã… (A35:J70)
    If dict.count > 0 Then
        With wsTarget.Range("A35:J70")
            .Merge
            .value = reportText
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlTop
            .WrapText = True
            .ReadingOrder = xlRTL
            .Font.Name = "Simplified Arabic"
            .Font.Size = 16
            .Font.Bold = True
        End With
    End If

    ' 6.  ’œÌ— «·„·›
    factoryTitle = Sheets("title_factory").Range("A2").value
    folderPath = "D:\export1\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    timeStamp = Format(Now, "yyyy-mm-dd_hh-mm-ss")
    fileName = "Œ·«’…_«·ÃÂ« _«·„Ã„⁄…_" & factoryTitle & "_" & timeStamp & ".xlsx"
    
    lastRowTarget = 70
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    
    wsTarget.Range("A1:J" & lastRowTarget).Copy
    With wsNew.Range("A1")
        .PasteSpecial Paste:=xlPasteAll
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With

    ' 7. ≈÷«›… «· –ÌÌ· ›Ì «·‰Â«Ì…
    footerRow = lastRowTarget + 3
    With wsNew.Range(wsNew.Cells(footerRow, 1), wsNew.Cells(footerRow + 1, 10))
        .Merge
        .value = "«·„‰Ÿ„             «·„Õ«”»             «·„œﬁﬁ             „œÌ— «· œﬁÌﬁ             „œÌ— «·„«·Ì…             „œÌ— «·„’‰⁄"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Borders.Weight = xlThin
    End With

    Application.DisplayAlerts = False
    wbNew.SaveAs fileName:=folderPath & fileName, FileFormat:=xlOpenXMLWorkbook
    wbNew.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    MsgBox " „  ’œÌ— «·Œ·«’… «·„Ã„⁄… ··ÃÂ«  »‰Ã«Õ.", vbInformation
End Sub

