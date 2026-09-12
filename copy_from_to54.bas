Attribute VB_Name = "copy_from_to54"
Sub copy_from_to54_1()

Dim arrData As Variant
    Dim targetArr() As Variant
    Dim colOrder As Variant
    Dim i As Long, r As Long, lastRow As Long
    Dim sourceSheet As Worksheet, targetSheet As Worksheet
    
    '  ”—Ì⁄ «·ﬂÊœ ·√ﬁ’Ï œ—Ã…
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set sourceSheet = Sheets(1)
    Set targetSheet = Sheets(54)
    
    ' «· — Ì» «·„ÿ·Ê»
    colOrder = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 25, 26, 27, 28, 29, 83, 30, 31, 32, 33, 23, 24, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 50, 47, 48, 49, 51, 52, 53, 54, 55, 56, 58, 57)

    lastRow = sourceSheet.Cells(sourceSheet.rowS.count, "A").End(xlUp).row
    If lastRow < 8 Then GoTo Cleanup
    
    ' ”Õ» «·»Ì«‰«  ··„’›Ê›…
    arrData = sourceSheet.Range(sourceSheet.Cells(8, 1), sourceSheet.Cells(lastRow, 83)).value
    ReDim targetArr(1 To UBound(arrData, 1), 1 To UBound(colOrder) + 1)

    ' --- «· ⁄œÌ·: ‰ﬁ· «·ﬁÌ„ „»«‘—… »œÊ‰ » — «·ﬂ”Ê— ---
    For i = LBound(colOrder) To UBound(colOrder)
        For r = 1 To UBound(arrData, 1)
            ' ‰ﬁ· «·ﬁÌ„… ﬂ„« ÂÌ  „«„« ”Ê«¡ ﬂ«‰  ‰’« √Ê —ﬁ„« »ﬂ”Ê—
            targetArr(r, i + 1) = arrData(r, colOrder(i))
        Next r
    Next i

    '  ‰ŸÌ› «·‘Ì  «·Âœ› Ê’» «·»Ì«‰«  «·ÃœÌœ…
    targetSheet.Cells.Clear
    With targetSheet.Range("A8").Resize(UBound(targetArr, 1), UBound(targetArr, 2))
        .value = targetArr
        '  „  €ÌÌ— «· ‰”Ìﬁ ·‹ General ·Ìﬁ»· «·ﬂ”Ê—° √Ê Ì„ﬂ‰ﬂ  Œ’Ì’Â Õ”» Õ«Ã ﬂ „À· "0.00"
        .NumberFormat = "General"
    End With

    ' ‰ﬁ· «· ‰”Ìﬁ „‰ «·’› «·À«„‰
    sourceSheet.rowS(8).Copy
    targetSheet.rowS(8).PasteSpecial Paste:=xlPasteFormats
    
    targetSheet.Columns.AutoFit

Cleanup:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    targetSheet.Activate
    MsgBox " „  —ÕÌ· «·»Ì«‰«  »‰Ã«Õ „⁄ «·«Õ ›«Ÿ »«·ﬁÌ„ «·⁄‘—Ì…!", vbInformation, "⁄„«œ «·‰⁄Ì„Ì"


End Sub

