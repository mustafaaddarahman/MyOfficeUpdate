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
    
    ' «· — Ì» «·„ÿ·Ê» (Õ”» „« √—”· Â √‰ )
    colOrder = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 25, 26, 27, 28, 29, 83, 30, 31, 32, 33, 23, 24, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 50, 47, 48, 49, 51, 52, 53, 54, 55, 56, 58, 57)

    lastRow = sourceSheet.Cells(sourceSheet.Rows.count, "A").End(xlUp).row
    If lastRow < 8 Then GoTo CleanUp
    
    ' ”Õ» «·»Ì«‰«  ··„’›Ê›…
    arrData = sourceSheet.Range(sourceSheet.Cells(8, 1), sourceSheet.Cells(lastRow, 83)).value
    ReDim targetArr(1 To UBound(arrData, 1), 1 To UBound(colOrder) + 1)

    ' --- »œ«Ì… «· ⁄œÌ· «·ÃÊÂ—Ì ---
    For i = LBound(colOrder) To UBound(colOrder)
        For r = 1 To UBound(arrData, 1)
            Dim val As Variant
            val = arrData(r, colOrder(i))
            
            ' ≈–« ﬂ«‰  «·ﬁÌ„… —ﬁ„«° ‰Õ–› «·ﬂ”— ‰Â«∆Ì« »«” Œœ«„ Fix
            ' Â–« Ì÷„‰ √‰ «·—ﬁ„ ”Ì‰“· ﬂﬁÌ„… ’ÕÌÕ… ’—ÌÕ… (Value) Ê·Ì” „Ã—œ  ‰”Ìﬁ
            If IsNumeric(val) And Not IsEmpty(val) Then
                targetArr(r, i + 1) = Fix(val)
            Else
                targetArr(r, i + 1) = val
            End If
        Next r
    Next i
    ' --- ‰Â«Ì… «· ⁄œÌ· «·ÃÊÂ—Ì ---

    '  ‰ŸÌ› «·‘Ì  «·Âœ› Ê’» «·»Ì«‰«  «·ÃœÌœ…
    targetSheet.Cells.Clear
    With targetSheet.Range("A8").Resize(UBound(targetArr, 1), UBound(targetArr, 2))
        .value = targetArr
        .NumberFormat = "0" ' ÷„«‰  ‰”Ìﬁ «·√—ﬁ«„ ﬂ√⁄œ«œ ’ÕÌÕ…
    End With

    ' ‰ﬁ· «· ‰”Ìﬁ „‰ «·’› «·À«„‰
    sourceSheet.Rows(8).Copy
    targetSheet.Rows(8).PasteSpecial Paste:=xlPasteFormats
    
    targetSheet.Columns.AutoFit

CleanUp:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    targetSheet.Activate
    MsgBox " „  —ÕÌ· ’—›Ì… «·—« » ·Â–« «·‘Â— »‰Ã«Õ ﬂ√—ﬁ«„ ’ÕÌÕ… ’—ÌÕ…!", vbInformation, "⁄„«œ «·‰⁄Ì„Ì"




End Sub

