Attribute VB_Name = "summ_all_after_import"
Sub CalculateSalaryPositions()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' ÊÍÏíÏ ÇáÔíÊ ÇáãÓÊåÏİ
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 1. ÊÍÏíÏ ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ İí ÇáÚãæÏ BF áÖãÇä ÏíäÇãíßíÉ ÇáãÚÇÏáÇÊ
    lastRow = ws.Cells(ws.Rows.count, "BF").End(xlUp).row
    
    ' ÇáÊÃßÏ ãä Ãä ÇáäØÇŞ íÈÏÃ ãä ÇáÕİ 9 ßÍÏ ÃÏäì
    If lastRow < 9 Then lastRow = 9
    
    ' 2. ÊæáíÏ ÇáãÚÇÏáÇÊ ÏÇÎá ÇáÎáÇíÇ ÇáãÍÏÏÉ ÈÏŞÉ (ãÚÇÏáÇÊ ÍíÉ)
    With ws
        ' ÇáÑÇÊÈ ÇáÇÓãí ááãÓÊãÑíä ÈÇáÎÏãÉ
        .Range("CS10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ãÓÊãÑ ÈÇáÎÏãÉ"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí ááãÌÇÒíä (5 ÓäæÇÊ)
        .Range("CT10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""íÊãÊÚ ÈÇÌÇÒå ÎãÓ ÓäæÇÊ"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí áãÚíá ÇáãÊİÑÛ
        .Range("CU10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ÅÌÇÒÉ ãÚíá ÇáãÊİÑÛ"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí ááÌÇä ÇáØÈíÉ (Êã ÊÕÍíÍ ÇáÊÓãíÉ åäÇ)
        .Range("CV10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ÅÌÇÒÉ áÌÇä ØÈíå"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÕÇİí (ÈÇÓÊÎÏÇã SUBTOTAL ááŞíã ÇáãİáÊÑÉ)
        .Range("CW10").Formula = "=SUBTOTAL(9, AW9:AW" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí ááÃãæãÉ (Ãæá 6 ÃÔåÑ)
        .Range("CX10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""Çãæãå Çæá 6 ÇÔåÑ"", J9:J" & lastRow & ")"
        
        ' ÇáÅÌÇÒÇÊ ÈÏæä ÑÇÊÈ
        .Range("CY10").Formula = "=SUMIF(BF9:BF" & lastRow & ","" ÇÌÇÒÉ ÈÏæä ÑÇÊÈ"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí ááãäÓÈíä (ÊäÓíÈ)
        .Range("CZ10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ÊäÓíÈ"", J9:J" & lastRow & ")"
        
        ' ÑæÇÊÈ ÇáÚŞæÏ
        .Range("DA10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ÚŞÏ"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí ááÑÇÊÈ ÇáÌÒÆí
        .Range("DB10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ÑÇÊÈ ÌÒÆí"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí áÑİÚ ÇáíÏ Ãæ ÅíŞÇİ ÇáÕÑİ
        .Range("DC10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""ÑİÚ íÏ Çæ ÅíŞÇİ ÕÑİ"", J9:J" & lastRow & ")"
        
        ' ÇáÑÇÊÈ ÇáÇÓãí ááÃãæãÉ (ËÇäí 6 ÃÔåÑ)
        .Range("DD10").Formula = "=SUMIF(BF9:BF" & lastRow & ",""Çãæãå ËÇäí 6 ÇÔåÑ"", J9:J" & lastRow & ")"
        
        ' ÊäÓíŞ ÇáÃÑŞÇã áÊÙåÑ ÈİÇÕáÉ ÇáÂáÇİ
        .Range("CS10:DD10").NumberFormat = "#,##0"
    End With

End Sub


Sub UpdateLiveFormulas_Sheet1()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' ÊÍÏíÏ ÇáÔíÊ ÇáãÓÊåÏİ
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 1. ÊÍÏíÏ ÂÎÑ Õİ ÈäÇÁğ Úáì ÇáÚãæÏ BF (Ãæ Ãí ÚãæÏ ÈíÇäÇÊ)
    lastRow = ws.Cells(ws.Rows.count, "BF").End(xlUp).row
    If lastRow < 9 Then lastRow = 9
    
    ' 2. ÒÑÚ ÇáãÚÇÏáÇÊ ÇáÏíäÇãíßíÉ İí ÇáÎáÇíÇ ÇáãÍÏÏÉ
    With ws
        ' --- ÇáÓØÑ ÇáËÇäí (CC2:CH2) ---
        .Range("CC2").Formula = "=SUBTOTAL(9, X9:X" & lastRow & ")"
        .Range("CD2").Formula = "=SUBTOTAL(9, J9:J" & lastRow & ")"
        .Range("CE2").Formula = "=SUBTOTAL(9, AD9:AD" & lastRow & ")"
        .Range("CF2").Formula = "=SUBTOTAL(9, AE9:AE" & lastRow & ")"
        .Range("CG2").Formula = "=SUBTOTAL(9, AV9:AV" & lastRow & ")"
        .Range("CH2").Formula = "=SUBTOTAL(9, AW9:AW" & lastRow & ")"
        
        ' --- ÇáÓØÑ ÇáËÇáË (CJ3:CZ3) ---
        .Range("CJ3").Formula = "=SUBTOTAL(9, X9:X" & lastRow & ")"
        .Range("CK3").Formula = "=SUBTOTAL(9, AF9:AF" & lastRow & ")"
        .Range("CL3").Formula = "=SUM(AG9:AG" & lastRow & ")"
        .Range("CM3").Formula = "=SUBTOTAL(9, AH9:AH" & lastRow & ")"
        .Range("CN3").Formula = "=SUBTOTAL(9, AI9:AI" & lastRow & ")"
        .Range("CO3").Formula = "=SUBTOTAL(9, AJ9:AJ" & lastRow & ")"
        .Range("CP3").Formula = "=SUBTOTAL(9, AL9:AL" & lastRow & ")"
        .Range("CQ3").Formula = "=SUBTOTAL(9, AN9:AN" & lastRow & ")"
        .Range("CS3").Formula = "=SUBTOTAL(9, J9:J" & lastRow & ")"
        .Range("CT3").Formula = "=SUBTOTAL(9, N9:N" & lastRow & ")"
        .Range("CU3").Formula = "=SUBTOTAL(9, P9:P" & lastRow & ")"
        .Range("CV3").Formula = "=SUBTOTAL(9, R9:R" & lastRow & ")"
        .Range("CW3").Formula = "=SUBTOTAL(9, S9:S" & lastRow & ")"
        .Range("CX3").Formula = "=SUBTOTAL(9, T9:T" & lastRow & ")"
        .Range("CY3").Formula = "=SUBTOTAL(9, U9:U" & lastRow & ")"
        .Range("CZ3").Formula = "=SUBTOTAL(9, Y9:Y" & lastRow & ")"
        
        ' --- ÇáÓØÑ ÇáÑÇÈÚ (CC4:CH4) - ãÑÊÈØ ÈÇáÔÑØ CB4 ---
        .Range("CC4").Formula = "=SUMIF($BF$9:BF" & lastRow & ", $CB$4, X9:X" & lastRow & ")"
        .Range("CD4").Formula = "=SUMIF($BF$9:BF" & lastRow & ", $CB$4, J9:J" & lastRow & ")"
        .Range("CE4").Formula = "=SUMIF($BF$9:BF" & lastRow & ", $CB$4, AD9:AD" & lastRow & ")"
        .Range("CF4").Formula = "=SUMIF($BF$9:BF" & lastRow & ", $CB$4, AE9:AE" & lastRow & ")"
        .Range("CG4").Formula = "=SUMIF($BF$9:BF" & lastRow & ", $CB$4, AV9:AV" & lastRow & ")"
        .Range("CH4").Formula = "=SUMIF($BF$9:BF" & lastRow & ", $CB$4, AW9:AW" & lastRow & ")"
        
        ' --- ÇáÓØÑ ÇáÎÇãÓ (CJ5:CX5) ---
        .Range("CJ5").Formula = "=SUBTOTAL(9, AP9:AP" & lastRow & ")"
        .Range("CK5").Formula = "=SUBTOTAL(9, AR9:AR" & lastRow & ")"
        .Range("CL5").Formula = "=SUBTOTAL(9, AX9:AX" & lastRow & ")"
        .Range("CM5").Formula = "=SUBTOTAL(9, AT9:AT" & lastRow & ")"
        .Range("CN5").Formula = "=SUM(AU9:AU" & lastRow & ")"
        .Range("CO5").Formula = "=SUBTOTAL(9, BJ9:BJ" & lastRow & ")"
        .Range("CP5").Formula = "=SUBTOTAL(9, BO9:BO" & lastRow & ")"
        .Range("CQ5").Formula = "=SUBTOTAL(9, BQ9:BQ" & lastRow & ")"
        .Range("CS5").Formula = "=SUBTOTAL(9, AA9:AA" & lastRow & ")"
        .Range("CT5").Formula = "=SUBTOTAL(9, AC9:AC" & lastRow & ")"
        .Range("CU5").Formula = "=SUBTOTAL(9, BH9:BH" & lastRow & ")"
        .Range("CV5").Formula = "=SUBTOTAL(9, BP9:BP" & lastRow & ")"
        .Range("CW5").Formula = "=SUBTOTAL(9, BR9:BR" & lastRow & ")"
        .Range("CX5").Formula = "=SUBTOTAL(9, CE9:CE" & lastRow & ")"

        ' 3. áãÓÉ ÇÍÊÑÇİíÉ: ÊäÓíŞ ÇáÃÑŞÇã ÈİÇÕáÉ ÇáÂáÇİ
        .Range("CC2:CZ5").NumberFormat = "#,##0"
    End With
End Sub


Sub FillAT()
    Dim lr As Long
    lr = Cells(Rows.count, "J").End(xlUp).row
    If lr >= 9 Then Range("AT9:AT" & lr).Formula = "=J9*0.0025"
End Sub
Sub ClearAG()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.ActiveSheet
    
    ' ÊÍÏíÏ ÂÎÑ Õİ ÈäÇÁğ Úáì ÇáÚãæÏ BF áÖãÇä Ôãæá ÌãíÚ ÇáÈíÇäÇÊ
    lastRow = ws.Cells(ws.Rows.count, "BF").End(xlUp).row
    
    ' ÊÚØíá ÊÍÏíË ÇáÔÇÔÉ áÊÓÑíÚ ÇáßæÏ
    Application.ScreenUpdating = False
    
    ' ÍáŞÉ ÊßÑÇÑíÉ ÊÈÏÃ ãä ÇáÕİ 9 ÍÊì ÂÎÑ Õİ
    For i = 9 To lastRow
        ' ÇáÊÍŞŞ ÅĞÇ ßÇäÊ ŞíãÉ ÇáÚãæÏ BF áÇ ÊÍÊæí Úáì äÕ ÇáÅÌÇÒÉ ÇáãØáæÈÉ
        ' ãáÇÍÙÉ: ÊÃßÏ Ãä ÇáäÕ ÇáãßÊæÈ "ÅÌÇÒÉ ÎãÓ ÓäæÇÊ" ãØÇÈŞ ÊãÇãÇğ áãÇ İí ÇáÎáÇíÇ
        If ws.Cells(i, "BF").value <> "ÇÌÇÒÉ ÎãÓ ÓäæÇÊ" Then
            ws.Cells(i, "AG").value = 0
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "ÊãÊ ÚãáíÉ ÇáÊÕİíÑ ÈÇÓÊËäÇÁ ÇáãÊãÊÚíä ÈÅÌÇÒÉ ÎãÓ ÓäæÇÊ.", vbInformation
End Sub


Sub CalculateAG()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.ActiveSheet ' Ãæ ÍÏÏ ÇÓã ÇáæÑŞÉ Sheets("Sheet1")
    
    ' ÇáÈÍË Úä ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ İí ÇáÚãæÏ J
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row
    
    ' ÇáÊÃßÏ ãä æÌæÏ ÈíÇäÇÊ ááÈÏÁ ãä ÇáÕİ ÇáÊÇÓÚ
    If lastRow >= 9 Then
        With ws.Range("AG9:AG" & lastRow)
            ' æÖÚ ÇáÕíÛÉ ÇáÍÓÇÈíÉ (ÖÑÈ ÇáÚãæÏ J İí 0.15)
            .Formula = "=J9*0.15"
            ' ÇÎÊíÇÑ ÇÎÊíÇÑí: ÊÍæíá ÇáÕíÛ Åáì Şíã ËÇÈÊÉ áÊÓÑíÚ Çáãáİ
            '.Value = .Value
        End With
    End If
End Sub

Sub UpdateSummaryFormula()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' æÖÚ ÇáãÚÇÏáÉ İí ÇáÎáíÉ CI4
    ' Êã ÇÓÊÎÏÇã ÇáÕíÛÉ ÇáãÊæÇİŞÉ ãÚ VBA ãÚ ÊÛííÑ ÇáİæÇÕá Åáì İÇÕáÉ ÚÇÏíÉ
    ws.Range("CI4").Formula = "=SUMIF($BF$9:$BF$5000, $CB$4, AG9:AG5000)"
End Sub
Sub CalculateFiveYearLeaveOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' ÊÍÏíÏ æÑŞÉ ÇáÚãá
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ÊÍÏíÏ ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ İí ÇáÚãæÏ BF
    lastRow = ws.Cells(ws.Rows.count, "BF").End(xlUp).row
    
    ' ÅĞÇ ßÇäÊ ÇáÈíÇäÇÊ ÊÈÏÃ ÈÚÏ ÇáÕİ 9¡ íÊã ÇáÎÑæÌ ãä ÇáÏÇáÉ
    If lastRow < 9 Then Exit Sub
    
    ' ÅíŞÇİ ÊÍÏíË ÇáÔÇÔÉ áÊÓÑíÚ ÇáßæÏ
    Application.ScreenUpdating = False
    
    ' ÇáÍáŞÉ ÇáÊßÑÇÑíÉ ÊÈÏÃ ãä ÇáÕİ 9 æáÛÇíÉ äåÇíÉ ÇáÈíÇäÇÊ
    For i = 9 To lastRow
        ' ÇáãÚíÇÑ: ÅĞÇ ßÇäÊ ŞíãÉ ÇáÎáíÉ İí ÇáÚãæÏ BF åí "íÊãÊÚ ÈÇÌÇÒå ÎãÓ ÓäæÇÊ"
        If ws.Cells(i, "BF").value = "íÊãÊÚ ÈÇÌÇÒå ÎãÓ ÓäæÇÊ" Then
            ' ÊäİíĞ ÇáÍÓÇÈ: AG = J * 0.15
            ws.Cells(i, "AG").value = ws.Cells(i, "J").value * 0.15
        End If
    Next i
    
    ' ÅÚÇÏÉ ÊİÚíá ÊÍÏíË ÇáÔÇÔÉ
    Application.ScreenUpdating = True
    
    MsgBox "Êã ÇáÇäÊåÇÁ ãä ÇÍÊÓÇÈ ÇáŞíã ááİÆÉ ÇáãÍÏÏÉ.", vbInformation
End Sub

Sub SaveWorkbook()
    ' ÍİÙ ÇáÊÛííÑÇÊ İí ãáİ ÇáÅßÓá ÇáÍÇáí
    ThisWorkbook.Save
    
    ' ÅÙåÇÑ ÑÓÇáÉ ÊÃßíÏ (ÇÎÊíÇÑí)
    MsgBox "Êã ÍİÙ ÇáÈíÇäÇÊ ÈäÌÇÍ!", vbInformation, "ÍİÙ"
End Sub

Sub CalculateGradesAndStatus()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim gradeValue As Variant
    Dim isOnLeave As Boolean
    
    ' ÊÍÏíÏ æÑŞÉ ÇáÚãá
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ÊÍÏíÏ ÂÎÑ Õİ ÈäÇÁğ Úáì ÇáÚãæÏ H (ÇáÏÑÌÇÊ)
    lastRow = ws.Cells(ws.Rows.count, "H").End(xlUp).row
    
    ' ÇáÊÃßÏ ãä æÌæÏ ÈíÇäÇÊ ÇÈÊÏÇÁğ ãä ÇáÕİ 9
    If lastRow < 9 Then Exit Sub
    
    ' ÅíŞÇİ ÊÍÏíË ÇáÔÇÔÉ áÊÓÑíÚ ÇáÊäİíĞ
    Application.ScreenUpdating = False
    
    For i = 9 To lastRow
        gradeValue = ws.Cells(i, "H").value
        ' ÇáÊÍŞŞ ãä ÍÇáÉ ÇáÅÌÇÒÉ (True ÅĞÇ ßÇä íÊãÊÚ ÈÇáÅÌÇÒÉ)
        isOnLeave = (Trim(ws.Cells(i, "BF").value) = "íÊãÊÚ ÈÇÌÇÒå ÎãÓ ÓäæÇÊ")
        
        ' ÇáÔÑØ ÇáÏŞíŞ: ÏÑÌÉ >= 8 æó áíÓ áÏíå ÅÌÇÒÉ
        If IsNumeric(gradeValue) And gradeValue >= 8 And Not isOnLeave Then
            ' ÇÍÊÓÇÈ ÇáŞíãÉ
            ws.Cells(i, "CE").value = ws.Cells(i, "J").value * 0.5
            ' æÖÚ ŞíãÉ True İí CD
            ws.Cells(i, "CD").value = True
        
        ' ãÇ Ïæä Ğáß (ÏÑÌÉ ÃŞá ãä 8 Ãóæú æÌæÏ ÅÌÇÒÉ)
        Else
            ' ÊÕİíÑ ÇáŞíãÉ
            ws.Cells(i, "CE").value = 0
            ' æÖÚ ŞíãÉ False İí CD
            ws.Cells(i, "CD").value = False
        End If
    Next i
    
    ' ÅÚÇÏÉ ÊİÚíá ÊÍÏíË ÇáÔÇÔÉ
    Application.ScreenUpdating = True
    
    MsgBox "ÊãÊ ÇáãÚÇáÌÉ ÈÏŞÉ:" & vbCrLf & _
           "- ÏÑÌÉ >= 8 æÈÏæä ÅÌÇÒÉ: ÇÍÊÓÇÈ æ True" & vbCrLf & _
           "- ÛíÑ Ğáß: ÊÕİíÑ æ False", vbInformation
End Sub

