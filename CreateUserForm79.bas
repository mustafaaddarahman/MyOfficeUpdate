Attribute VB_Name = "CreateUserForm79"
Option Explicit

' ==============================================================================
' «”„ «·„ÊœÌÊ·: SalaryAuditInterfaceBuilder_V4
' «·Ê’›: ÌﬁÊ„ » ’„Ì„ «·Ê«ÃÂ… «·—ﬁ«»Ì… »«·ﬂ«„· œ«Œ· ‘Ì  "between_two_month" «·›«—€
'        ÊÌ⁄ „œ ⁄·Ï ﬁ—«¡… √”„«¡ —ƒÊ” «·√⁄„œ… «·ÕﬁÌﬁÌ… „‰ «·”ÿ— 8 ›Ì Sheet1
'        „⁄  ”„Ì… √⁄„œ… «·„ﬁ«—‰…  ·ﬁ«∆Ì« »√”„«¡ «·√‘Â— «·ÕﬁÌﬁÌ… «·„” ‰œ… · «—ÌŒ «·ÊÌ‰œÊ“.
' ==============================================================================

Private Type MonthSummary
    CurrentAllowances(1 To 14) As Double
    PastAllowances(1 To 14) As Double
    CurrentDeductions(1 To 16) As Double
    PastDeductions(1 To 16) As Double
    CurrentNetSalary As Double
    PastNetSalary As Double
End Type

Sub DesignAndFillAuditSheet()
    Dim wsCurrent As Worksheet, wsPast As Worksheet, wsReport As Worksheet
    Dim lrCurrent As Long, lrPast As Long
    Dim arrCurrent As Variant, arrPast As Variant
    Dim AuditData As MonthSummary
    Dim AllowCols As Variant, DeductCols As Variant
    Dim AllowIdx As Variant, DeductIdx As Variant
    Dim i As Long, j As Integer, r As Long
    Dim Diff As Double, NetColIdx As Long
    Dim ReportSheetName As String
    Dim headerRow As Long
    
    ' „ €Ì—«  „”„Ì«  «·√‘Â— «·œÌ‰«„ÌﬂÌ…
    Dim CurrentMonthName As String
    Dim PastMonthName As String
    
    ReportSheetName = "between_two_month"
    NetColIdx = 49 ' ⁄„Êœ ’«›Ì «·—« » AW
    headerRow = 8  ' «·”ÿ— «·–Ì ÌÕ ÊÌ ⁄·Ï √”„«¡ «·⁄‰«ÊÌ‰ Ê—ƒÊ” «·√⁄„œ…
    
    ' «” Œ—«Ã √”„«¡ «·√‘Â— „‰ ‰Ÿ«„ «·ÊÌ‰œÊ“  ·ﬁ«∆Ì« («·‘Â— «·Õ«·Ì Ê«·‘Â— «·›«∆ )
    CurrentMonthName = Format(Date, "mmmm yyyy")
    PastMonthName = Format(DateAdd("m", -1, Date), "mmmm yyyy")
    
    ' 1. «· Õﬁﬁ „‰ «·‘Ì «  «·√”«”Ì… Ê«·‘Ì  «·„” Âœ› «·›«—€
    On Error Resume Next
    Set wsCurrent = ThisWorkbook.Sheets("Sheet1")
    Set wsPast = ThisWorkbook.Sheets("past_month76")
    Set wsReport = ThisWorkbook.Sheets(ReportSheetName)
    On Error GoTo 0
    
    If wsCurrent Is Nothing Or wsPast Is Nothing Then
        MsgBox "Œÿ√:  √ﬂœ „‰ ÊÃÊœ «·‘Ì «  «·„—Ã⁄Ì… (Sheet1) Ê (past_month55).", vbCritical, "›‘· «·‰Ÿ«„"
        Exit Sub
    End If
    
    If wsReport Is Nothing Then
        MsgBox "Œÿ√: «·‘Ì  '" & ReportSheetName & "' €Ì— „ÊÃÊœ! Ì—ÃÏ ≈‰‘«ƒÂ ﬂ‘Ì  ›«—€ √Ê·«.", vbCritical, "‘Ì  „›ﬁÊœ"
        Exit Sub
    End If
    
    ' 2. „⁄«·Ã… «·»Ì«‰«  »”—⁄… ›«∆ﬁ… œ«Œ· «·–«ﬂ—… (RAM)
    lrCurrent = wsCurrent.Cells(wsCurrent.rowS.count, "AW").End(xlUp).row
    lrPast = wsPast.Cells(wsPast.rowS.count, "AW").End(xlUp).row
    If lrCurrent < 9 Then lrCurrent = 9
    If lrPast < 9 Then lrPast = 9
    
    arrCurrent = wsCurrent.Range("A1:CE" & lrCurrent).value
    arrPast = wsPast.Range("A1:CE" & lrPast).value
    
    ' √—ﬁ«„ √œ·… «·√⁄„œ… ›Ì „’›Ê›… «·‹ VBA
    AllowIdx = Array(14, 16, 18, 19, 20, 21, 25, 27, 29, 60, 68, 70, 10, 83)
    DeductIdx = Array(24, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46, 47, 62, 67, 69, 50)
    
    '  Ã„Ì⁄ ﬁÌ„ «·«” Õﬁ«ﬁ«  Ê«·’«›Ì ›Ì «·–«ﬂ—…
    For i = 9 To lrCurrent
        For j = 0 To 13
            AuditData.CurrentAllowances(j + 1) = AuditData.CurrentAllowances(j + 1) + val(arrCurrent(i, AllowIdx(j)))
        Next j
        AuditData.CurrentNetSalary = AuditData.CurrentNetSalary + val(arrCurrent(i, NetColIdx))
    Next i
    
    For i = 9 To lrPast
        For j = 0 To 13
            AuditData.PastAllowances(j + 1) = AuditData.PastAllowances(j + 1) + val(arrPast(i, AllowIdx(j)))
        Next j
        AuditData.PastNetSalary = AuditData.PastNetSalary + val(arrPast(i, NetColIdx))
    Next i
    
    '  Ã„Ì⁄ ﬁÌ„ «·«” ﬁÿ«⁄«  Ê«·„”ÕÊ»«  ›Ì «·–«ﬂ—…
    For i = 9 To lrCurrent
        For j = 0 To 15
            AuditData.CurrentDeductions(j + 1) = AuditData.CurrentDeductions(j + 1) + val(arrCurrent(i, DeductIdx(j)))
        Next j
    Next i
    
    For i = 9 To lrPast
        For j = 0 To 15
            AuditData.PastDeductions(j + 1) = AuditData.PastDeductions(j + 1) + val(arrPast(i, DeductIdx(j)))
        Next j
    Next i
    
    ' 3. Â‰œ”… Ê ’„Ì„ «·‘Ì  «·›«—€ Ê»‰«¡ «·Ê«ÃÂ… »—„ÃÌ«
    With wsReport
        .Cells.Clear '  ‰ŸÌ› «·‘Ì   „«„« ·· √ﬂœ „‰ √‰Â ›«—€ ﬁ»· «· ’„Ì„
        .DisplayRightToLeft = True
        .Cells.Font.Name = "Segoe UI"
        .Cells.Font.Size = 10
        
        ' √ -  ’„Ì„ ‘—Ìÿ «·⁄‰Ê«‰ «·„—ﬂ“Ì «·›Œ„
        With .Range("A1:E1")
            .Merge
            .value = "„’›Ê›… «·—ﬁ«»… Ê«·„ÿ«»ﬁ… «·„—ﬂ“Ì… ·—Ê« » «·„ÊŸ›Ì‰ („ﬁ«—‰… «·⁄‰«ÊÌ‰ »Ì‰ ‘Â—Ì‰)"
            .Font.Size = 14
            .Font.Bold = True
            .Interior.Color = RGB(26, 54, 93) ' √“—ﬁ „·ﬂÌ œ«ﬂ‰ ··„ƒ””« 
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 38
        End With
        
        ' » -  —ÊÌ”… ⁄‰«ÊÌ‰ «·√⁄„œ… Ê ‰”ÌﬁÂ« «·ÂÌﬂ·Ì »«· Ê«—ÌŒ «·œÌ‰«„ÌﬂÌ… „‰ «·ÊÌ‰œÊ“
        .Range("A3:E3").value = Array("«·»‰œ «·„” Âœ› (—√” «·⁄„Êœ «·„«·Ì)", "«·‘Â— «·Õ«·Ì (" & CurrentMonthName & ")", "«·‘Â— «·„«÷Ì (" & PastMonthName & ")", "«·›—ﬁ «·„«·Ì «·‰ﬁÌ", "„ƒ‘— «·Õ—ﬂ… «·—ﬁ«»Ì")
        With .Range("A3:E3")
            .Font.Bold = True
            .Interior.Color = RGB(226, 232, 240)
            .Font.Color = RGB(26, 54, 93)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .RowHeight = 26
        End With
        
        r = 4
        ' Ã - »‰«¡ ﬁ”„ «·«” Õﬁ«ﬁ«  Ê«· ⁄ÊÌ÷«  »√”„«∆Â« «·ÕﬁÌﬁÌ…
        .Cells(r, 1).value = "? ﬁ”‹‹‹‹„ «·«” Õﬁ«ﬁ‹‹‹‹«  «·„«·Ì‹‹‹‹… Ê«· ⁄ÊÌ÷‹‹‹‹« "
        .Range(.Cells(r, 1), .Cells(r, 5)).Merge
        .Range(.Cells(r, 1), .Cells(r, 5)).Font.Bold = True
        .Range(.Cells(r, 1), .Cells(r, 5)).Interior.Color = RGB(247, 250, 252)
        .Range(.Cells(r, 1), .Cells(r, 5)).Font.Color = RGB(44, 82, 130)
        .Range(.Cells(r, 1), .Cells(r, 5)).RowHeight = 22
        r = r + 1
        
        For j = 1 To 14
            .Cells(r, 1).value = wsCurrent.Cells(headerRow, AllowIdx(j - 1)).value
            If Trim(.Cells(r, 1).value) = "" Then .Cells(r, 1).value = "«” Õﬁ«ﬁ €Ì— „”„Ï (⁄„Êœ " & Split(wsCurrent.Cells(headerRow, AllowIdx(j - 1)).Address, "$")(1) & ")"
            
            .Cells(r, 2).value = AuditData.CurrentAllowances(j)
            .Cells(r, 3).value = AuditData.PastAllowances(j)
            Diff = AuditData.CurrentAllowances(j) - AuditData.PastAllowances(j)
            .Cells(r, 4).value = Diff
            WriteIndicator wsReport, r, Diff
            .rowS(r).RowHeight = 20
            r = r + 1
        Next j
        
        ' œ - »‰«¡ ﬁ”„ «·«” ﬁÿ«⁄«  Ê«·„”ÕÊ»«  »√”„«∆Â« «·ÕﬁÌﬁÌ…
        r = r + 1
        .Cells(r, 1).value = "? ﬁ”‹‹‹‹„ «·«” ﬁÿ«⁄‹‹‹‹«  Ê«·„”ÕÊ»‹‹‹‹«  «· ﬁ«⁄œÌ‹… Ê«·÷—Ì»Ì‹…"
        .Range(.Cells(r, 1), .Cells(r, 5)).Merge
        .Range(.Cells(r, 1), .Cells(r, 5)).Font.Bold = True
        .Range(.Cells(r, 1), .Cells(r, 5)).Interior.Color = RGB(247, 250, 252)
        .Range(.Cells(r, 1), .Cells(r, 5)).Font.Color = RGB(44, 82, 130)
        .Range(.Cells(r, 1), .Cells(r, 5)).RowHeight = 22
        r = r + 1
        
        For j = 1 To 16
            .Cells(r, 1).value = wsCurrent.Cells(headerRow, DeductIdx(j - 1)).value
            If Trim(.Cells(r, 1).value) = "" Then .Cells(r, 1).value = "«” ﬁÿ«⁄ €Ì— „”„Ï (⁄„Êœ " & Split(wsCurrent.Cells(headerRow, DeductIdx(j - 1)).Address, "$")(1) & ")"
            
            .Cells(r, 2).value = AuditData.CurrentDeductions(j)
            .Cells(r, 3).value = AuditData.PastDeductions(j)
            Diff = AuditData.CurrentDeductions(j) - AuditData.PastDeductions(j)
            .Cells(r, 4).value = Diff
            WriteIndicator wsReport, r, Diff
            .rowS(r).RowHeight = 20
            r = r + 1
        Next j
        
        ' Â‹ - »‰«¡ Œ·«’… «· œﬁÌﬁ «·Œ «„Ì… (’«›Ì «·—Ê« » «·≈Ã„«·Ì «·ÕﬁÌﬁÌ)
        r = r + 1
        .Range(.Cells(r, 1), .Cells(r, 5)).BorderAround xlContinuous, xlMedium
        r = r + 1
        
        .Cells(r, 1).value = wsCurrent.Cells(headerRow, NetColIdx).value
        If Trim(.Cells(r, 1).value) = "" Then .Cells(r, 1).value = "’«›Ì «·—Ê« » «·≈Ã„«·Ì «·‰Â«∆Ì"
        
        .Cells(r, 2).value = AuditData.CurrentNetSalary
        .Cells(r, 3).value = AuditData.PastNetSalary
        Diff = AuditData.CurrentNetSalary - AuditData.PastNetSalary
        .Cells(r, 4).value = Diff
        .Range(.Cells(r, 1), .Cells(r, 4)).Font.Bold = True
        
        .Range(.Cells(r, 1), .Cells(r, 5)).Interior.Color = RGB(237, 242, 247)
        .rowS(r).RowHeight = 24
        WriteIndicator wsReport, r, Diff
        
        ' Ê - ÷»ÿ «·ÕœÊœ Ê«· ‰”Ìﬁ«  «·Õ”«»Ì… ··ÃœÊ· «·„ﬂ „·
        With .Range("B4:D" & r)
            .NumberFormat = "#,##0.00"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        .Range("A3:E" & r).Borders.LineStyle = xlContinuous
        .Range("A3:E" & r).Borders.Color = RGB(203, 213, 224)
        
        ' ÷»ÿ √»⁄«œ «·√⁄„œ…  ·ﬁ«∆Ì« ·„‰⁄  œ«Œ· «·‰’Ê’ Ê«·√”„«¡ «·ÿÊÌ·…
        .Columns("A:E").AutoFit
        .Columns("A").ColumnWidth = 45
    End With
    
    MsgBox " „  ÕœÌÀ «·Ê«ÃÂ… »‰Ã«Õ° Ê Õ„· «·√⁄„œ… «·¬‰ √”„«¡ «·√‘Â— «·ÕﬁÌﬁÌ… ·‰Ÿ«„ «· ‘€Ì·!", vbInformation, "«ﬂ „· «· œﬁÌﬁ «·ÂÌﬂ·Ì Ê«·“„‰Ì"
End Sub

Private Sub WriteIndicator(ws As Worksheet, rowIdx As Long, DiffVal As Double)
    With ws.Cells(rowIdx, 5)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        If DiffVal > 0.009 Then
            .value = "“Ì«œ… ??"
            .Font.Color = RGB(34, 139, 34)
            .Font.Bold = True
            ws.Cells(rowIdx, 4).Font.Color = RGB(34, 139, 34)
        ElseIf DiffVal < -0.009 Then
            .value = "‰ﬁ’«‰ ??"
            .Font.Color = RGB(196, 30, 58)
            .Font.Bold = True
            ws.Cells(rowIdx, 4).Font.Color = RGB(196, 30, 58)
        Else
            .value = "„ ÿ«»ﬁ ??"
            .Font.Color = RGB(113, 128, 150)
            ws.Cells(rowIdx, 4).Font.Color = RGB(0, 0, 0)
        End If
    End With
End Sub
