Attribute VB_Name = "ba8e_alksoor"
Sub Process_Salary_Vault_CH()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim dataArr As Variant
    Dim awCol As Integer, bCol As Integer, eCol As Integer, chCol As Integer
    Dim currentSalary As Double, currentCH As Double, totalAsFraction As Double
    Dim resultsAW() As Variant, resultsCH() As Variant

    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ÅíÌÇÏ ÂÎÑ Õİ ÈäÇÁğ Úáì ÇáÑŞã ÇáæÙíİí İí ÇáÚãæÏ B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 9 Then Exit Sub

    ' ÊÍÏíÏ ÃÑŞÇã ÇáÃÚãÏÉ ÈÑãÌíÇğ áÖãÇä ÇáÏŞÉ
    bCol = ws.Range("B1").Column   ' ÇáÑŞã ÇáæÙíİí
    eCol = ws.Range("E1").Column   ' ÇáÇÓã
    awCol = ws.Range("AW1").Column ' ÇáÑÇÊÈ ÇáÕÇİí
    chCol = ws.Range("CH1").Column ' ÎÒíäÉ ÇáßÓæÑ (ÇáÚãæÏ ÇáÌÏíÏ)
    
    ' ŞÑÇÁÉ ÇáÈíÇäÇÊ ãä ÇáÚãæÏ A ÍÊì CH Åáì ãÕİæİÉ æÇÍÏÉ
    dataArr = ws.Range("A1:CH" & lastRow).value
    
    ' ÊÌåíÒ ãÕİæİÇÊ ÇáäÊÇÆÌ ááÃÚãÏÉ ÇáÊí ÓíÊã ÊÚÏíáåÇ
    ReDim resultsAW(9 To lastRow, 1 To 1)
    ReDim resultsCH(9 To lastRow, 1 To 1)

    For i = 9 To lastRow
        ' ÇáÊÍŞŞ ãä ÇáãÚÇííÑ (ÇáÑŞã ÇáæÙíİí æÇáÇÓã)
        If Trim(dataArr(i, bCol)) <> "" And Trim(dataArr(i, eCol)) <> "" Then
            
            ' 1. ÌáÈ ÇáŞíã ÇáÍÇáíÉ ãä ÇáãÕİæİÉ (ÇáĞÇßÑÉ)
            currentSalary = val(dataArr(i, awCol))
            currentCH = val(dataArr(i, chCol))
            
            ' 2. ÍÓÇÈ ÇáãÌãæÚ ÇáÊÑÇßãí: (ãÇ åæ ãæÌæÏ İí CH + ßÓÑ ÇáÑÇÊÈ ÇáÍÇáí)
            totalAsFraction = Round(currentCH + (currentSalary - Int(currentSalary)), 4)
            
            ' 3. ãäØŞ ÇáÌÈÑ æÇáÊÑÍíá:
            If totalAsFraction >= 1 Then
                ' ÅÖÇİÉ 1 ÕÍíÍ ááÑÇÊÈ
                resultsAW(i, 1) = Int(currentSalary) + 1
                ' ÎÕã 1 ãä ÇáÎÒíäÉ CH æÅÈŞÇÁ ÇáİßÉ ÇáãÊÈŞíÉ
                resultsCH(i, 1) = Round(totalAsFraction - 1, 4)
            Else
                ' ÇáÑÇÊÈ íÕÈÍ ÕÍíÍÇğ (ÈÏæä İßÉ)
                resultsAW(i, 1) = Int(currentSalary)
                ' ÈŞÇÁ ßÇãá ÇáŞíãÉ İí CH ßãÊÑÇßã
                resultsCH(i, 1) = totalAsFraction
            End If
            
        Else
            ' ÅĞÇ ßÇä ÇáÕİ İÇÑÛÇğ¡ äÍÊİÙ ÈÇáŞíã ÇáŞÏíãÉ ßãÇ åí
            resultsAW(i, 1) = dataArr(i, awCol)
            resultsCH(i, 1) = dataArr(i, chCol)
        End If
    Next i

    ' ÊäİíĞ ÇáßÊÇÈÉ ÇáäåÇÆíÉ ÏİÚÉ æÇÍÏÉ áÖãÇä ÇáÓÑÚÉ æÚÏã ÇáÊÚáíŞ
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' ÅÑÌÇÚ ÇáäÊÇÆÌ Åáì ÇáÔíÊ İí ÇáÃÚãÏÉ ÇáãÍÏÏÉ
    ws.Range("AW9").Resize(lastRow - 8, 1).value = resultsAW
    ws.Range("CH9").Resize(lastRow - 8, 1).value = resultsCH
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

   End Sub
