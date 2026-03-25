Attribute VB_Name = "klasaat_astktaat_daraga"

Sub ExportDeductionsDataFinal()
    Dim wbSource As Workbook, wbNew As Workbook
    Dim ws26 As Worksheet, ws31 As Worksheet, wsNew As Worksheet
    Dim folderPath As String, fileName As String, currentMonthName As String
    
    ' 1. ≈⁄œ«œ «·„”«— ›Ì «·»«— ‘‰ D
    folderPath = "D:\export1\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir "D:\export1"

    Set wbSource = ThisWorkbook
    Set ws26 = wbSource.Sheets(26)
    Set ws31 = wbSource.Sheets(31)
    
    currentMonthName = Format(Date, "mmmm")

    ' 2. ≈‰‘«¡ «·„·› «·ÃœÌœ
    Set wbNew = Workbooks.Add
    Set wsNew = wbNew.Sheets(1)
    wsNew.DisplayRightToLeft = True

    ' 3. «· —ÊÌ”…
    With wsNew.Range("A1:Q2")
        .Merge
        .value = "Œ·«’… «·«” ﬁÿ«⁄«  ·‘Â— " & currentMonthName & " - " & ws31.Range("A2").value
        .Font.Size = 24: .Font.Bold = True: .HorizontalAlignment = xlCenter
    End With

    ' 4. «” Œœ«„ Special Paste ·‰ﬁ· «·ﬁÌ„ ›ﬁÿ (»œÊ‰ „⁄«œ·« )
    ws26.Range("A22:Q36").Copy
    
    ' ·’ﬁ «·ﬁÌ„
    wsNew.Range("A3").PasteSpecial Paste:=xlPasteValues
    ' ·’ﬁ «· ‰”Ìﬁ«  («·√·Ê«‰ Ê«·ÕœÊœ)
    wsNew.Range("A3").PasteSpecial Paste:=xlPasteFormats
    ' ·’ﬁ ⁄—÷ «·√⁄„œ…
    wsNew.Range("A3").PasteSpecial Paste:=xlPasteColumnWidths

    ' 5. ≈÷«›… «· ÊﬁÌ⁄«  (ﬂ„« ›Ì ﬂÊœﬂ «·”«»ﬁ)
    Dim signatures As Variant, i As Integer, colStart As Integer
    signatures = Array("«·„‰Ÿ„", "«·„Õ«”»", "«·„œﬁﬁ", "„œÌ— «· œﬁÌﬁ", "«·„œÌ— «·„«·Ì", "„œÌ— «·„’‰⁄")
    colStart = 1
    For i = LBound(signatures) To UBound(signatures)
        With wsNew.Range(wsNew.Cells(20, colStart), wsNew.Cells(21, colStart + 2))
            .Merge: .value = signatures(i): .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        colStart = colStart + 3
    Next i

    ' 6. «·Õ›Ÿ Ê«·≈€·«ﬁ
    Application.CutCopyMode = False
    fileName = "Œ·«’…_" & currentMonthName & "_" & Format(Now, "hhmm") & ".xlsx"
    
    wbNew.SaveAs fileName:=folderPath & fileName
    wbNew.Close SaveChanges:=False
    
    MsgBox " „ «· ’œÌ— ﬂﬁÌ„ »‰Ã«Õ ≈·Ï: " & folderPath & fileName, vbInformation
End Sub





