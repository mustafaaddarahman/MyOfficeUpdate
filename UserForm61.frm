VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm61 
   Caption         =   "UserForm61"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   16608
   OleObjectBlob   =   "UserForm61.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim deductionCols As Variant
    Dim lastRowSrc As Long, lastRowDst As Long
    Dim i As Long, c As Long
    Dim EmpID As Variant, EmpName As String
    Dim DeductionCategory As String
    Dim DeductionValue As Double
    Dim duration As Long
    Dim colFound As Range
    Dim todayDate As Date, EndDate As Date
    Dim MonthsPassed As Long, monthsRemaining As Long
    Dim deductionType As String
    Dim RiskLevel As String
    Dim notesMsg As String
    Dim newID As Long
    Dim r As Long
    Dim hasExpired As Boolean, hasNearExpire As Boolean

    todayDate = Date
    Set wsSrc = Sheets("Sheet1")

    ' ≈‰‘«¡ /  ÂÌ∆… Financialsector
    On Error Resume Next
    Set wsDst = Sheets("Financialsector")
    On Error GoTo 0

    If wsDst Is Nothing Then
        Set wsDst = Sheets.Add
        wsDst.Name = "Financialsector"
    Else
        wsDst.Cells.ClearContents
    End If

    ' —ƒÊ” «·√⁄„œ…
    wsDst.Range("A1:O1").value = Array( _
        "—ﬁ„ «·«” ﬁÿ«⁄", "«·—ﬁ„ «·ÊŸÌ›Ì", "«·«”„", "›∆… «·«” ﬁÿ«⁄", _
        "‰Ê⁄ «·«” ﬁÿ«⁄", "ﬁÌ„… «·«” ﬁÿ«⁄", " «—ÌŒ «·»œ¡", "„œ… «·«” ﬁÿ«⁄ (‘Â—)", _
        " «—ÌŒ «·«‰ Â«¡", "«·√‘Â— «·„Õ ”»…", "ﬁÌ„… ‘Â—Ì…", "«·Õ«·…", _
        "«·„»·€ «· —«ﬂ„Ì", "«·„·«ÕŸ« ", "„” ÊÏ «·„Œ«ÿ—", "«·«‘Â— «·„ »ﬁÌ…")
    wsDst.Rows(1).Font.Bold = True

    newID = 1

    ' ›∆«  «·«” ﬁÿ«⁄ = √⁄„œ… Sheet1
    deductionCols = Array( _
        "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%", _
        "«·‰›ﬁ…", _
        "«·«” ﬁÿ«⁄ 1", _
        "«·«” ﬁÿ«⁄ «·„ ›—ﬁ", _
        "„»·€ «·€Ì«»", _
        "«·”·›…", _
        "«·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ ", _
        "«” ﬁÿ«⁄ ⁄«„", _
        "«” ﬁÿ«⁄ „»·€ Õ”» «·‰”»…")

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.count, "B").End(xlUp).row

    ' «·„—Ê— ⁄·Ï «·„ÊŸ›Ì‰
    For i = 9 To lastRowSrc

        EmpID = wsSrc.Cells(i, "B").value     ' «·—ﬁ„ «·ÊŸÌ›Ì
        EmpName = Trim(wsSrc.Cells(i, "E").value) ' «·«”„

        If EmpID = "" Or EmpName = "" Then GoTo NextEmployee

        For c = LBound(deductionCols) To UBound(deductionCols)

            Set colFound = wsSrc.Rows(8).Find( _
                What:=deductionCols(c), LookIn:=xlValues, LookAt:=xlWhole)

            If colFound Is Nothing Then GoTo NextDeduction

            DeductionCategory = colFound.value
            DeductionValue = val(wsSrc.Cells(i, colFound.Column).value)

            ' «·‘—ÿ «·√”«”Ì
            If DeductionValue <= 0 Then GoTo NextDeduction

            '  ÕœÌœ «·„œ…  ·ﬁ«∆Ì«
            If InStr(DeductionCategory, "%") > 0 Or InStr(DeductionCategory, "‰”»…") > 0 Then
                duration = 999
                deductionType = "‰”»…"
            ElseIf InStr(DeductionCategory, "”·›…") > 0 Or InStr(DeductionCategory, "„ƒﬁ ") > 0 Then
                duration = 12
                deductionType = "„»·€"
            Else
                duration = 1
                deductionType = "„»·€"
            End If

            EndDate = DateAdd("m", duration, todayDate)
            MonthsPassed = DateDiff("m", todayDate, EndDate) * -1
            If MonthsPassed < 0 Then MonthsPassed = 0

            monthsRemaining = duration - MonthsPassed
            If monthsRemaining < 0 Then monthsRemaining = 0

            ' «·Õ«·…
            If monthsRemaining = 0 Then
                notesMsg = "«‰ Â  „œ… «·«” ﬁÿ«⁄"
                wsDst.Cells(newID + 1, "L").value = "„‰ ÂÌ"
            ElseIf monthsRemaining <= 1 Then
                notesMsg = " ‰»ÌÂ: ﬁ«—» ⁄·Ï «·«‰ Â«¡"
                wsDst.Cells(newID + 1, "L").value = "ﬁ«—» ⁄·Ï «·«‰ Â«¡"
            Else
                notesMsg = "«” ﬁÿ«⁄ ‰‘ÿ"
                wsDst.Cells(newID + 1, "L").value = "‰‘ÿ"
            End If

            ' «·„Œ«ÿ—
            If DeductionValue >= 1000 Then
                RiskLevel = "⁄«·Ì"
            ElseIf DeductionValue >= 500 Then
                RiskLevel = "„ Ê”ÿ"
            Else
                RiskLevel = "„‰Œ›÷"
            End If

            lastRowDst = wsDst.Cells(wsDst.Rows.count, "A").End(xlUp).row + 1

            wsDst.Cells(lastRowDst, "A").value = newID
            wsDst.Cells(lastRowDst, "B").value = EmpID
            wsDst.Cells(lastRowDst, "C").value = EmpName
            wsDst.Cells(lastRowDst, "D").value = DeductionCategory
            wsDst.Cells(lastRowDst, "E").value = deductionType
            wsDst.Cells(lastRowDst, "F").value = DeductionValue
            wsDst.Cells(lastRowDst, "G").value = todayDate
            wsDst.Cells(lastRowDst, "H").value = duration
            wsDst.Cells(lastRowDst, "I").value = EndDate
            wsDst.Cells(lastRowDst, "J").value = MonthsPassed
            wsDst.Cells(lastRowDst, "K").value = DeductionValue / duration
            wsDst.Cells(lastRowDst, "M").value = RiskLevel
            wsDst.Cells(lastRowDst, "N").value = notesMsg
            wsDst.Cells(lastRowDst, "O").value = "„” ÊÏ «·„Œ«ÿ—"

            newID = newID + 1

NextDeduction:
        Next c

NextEmployee:
    Next i

    '  ·ÊÌ‰ «·Õ«·…
    hasExpired = False
    hasNearExpire = False
    For r = 2 To wsDst.Cells(wsDst.Rows.count, "L").End(xlUp).row
        Select Case wsDst.Cells(r, "L").value
            Case "‰‘ÿ"
                wsDst.Cells(r, "L").Interior.Color = RGB(198, 239, 206)
            Case "ﬁ«—» ⁄·Ï «·«‰ Â«¡"
                wsDst.Cells(r, "L").Interior.Color = RGB(255, 235, 156)
                hasNearExpire = True
            Case "„‰ ÂÌ"
                wsDst.Cells(r, "L").Interior.Color = RGB(255, 199, 206)
                hasExpired = True
        End Select
    Next r

    wsDst.Columns("A:O").AutoFit

    ' «· ‰»ÌÂ«  „»«‘—… »⁄œ «·„·¡
    If hasExpired Then
        MsgBox " ‰»ÌÂ: Â‰«ﬂ «” ﬁÿ«⁄«  «‰ Â  „œ… «” Õﬁ«ﬁÂ«!", vbCritical, "«” ﬁÿ«⁄«  „‰ ÂÌ…"
    End If

    If hasNearExpire Then
        MsgBox " ‰»ÌÂ: Â‰«ﬂ «” ﬁÿ«⁄«  ﬁ«—»  ⁄·Ï «·«‰ Â«¡!", vbExclamation, "«” ﬁÿ«⁄«  ﬁ—Ì»… „‰ «·«‰ Â«¡"
    End If

End Sub

Private Sub CommandButton2_Click()
 Dim ws As Worksheet
    Dim EmpID As String
    Dim lastRow As Long, r As Long
    
    Set ws = ThisWorkbook.Sheets("Financialsector")
    EmpID = Trim(Me.TextBox1.value)
    
    If EmpID = "" Then
        MsgBox "«·—Ã«¡ ≈œŒ«· «·—ﬁ„ «·ÊŸÌ›Ì.", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    For r = 2 To lastRow
        If ws.Cells(r, "B").Text = EmpID Then
            ws.Cells(r, "C").value = Me.TextBox2.value
            ws.Cells(r, "D").value = Me.TextBox3.value
            ws.Cells(r, "E").value = Me.TextBox4.value
            ws.Cells(r, "F").value = Me.TextBox5.value
            If IsDate(Me.TextBox6.value) Then ws.Cells(r, "G").value = CDate(Me.TextBox6.value)
            ws.Cells(r, "H").value = Me.TextBox7.value
            If IsDate(Me.TextBox8.value) Then ws.Cells(r, "I").value = CDate(Me.TextBox8.value)
            ws.Cells(r, "J").value = Me.TextBox9.value
           ws.Cells(r, "K").value = Me.TextBox10.value
            ws.Cells(r, "L").value = Me.TextBox11.value
            ws.Cells(r, "M").value = Me.TextBox14.value
            ws.Cells(r, "N").value = Me.TextBox12.value
             ws.Cells(r, "o").value = Me.TextBox16.value
             ws.Cells(r, "p").value = Me.TextBox15.value
            Exit For
        End If
    Next r
    
    MsgBox " „  ÕœÌÀ «·»Ì«‰«  ÊÕ”«» «·«” ﬁÿ«⁄«  »‰Ã«Õ.", vbInformation
End Sub

Private Sub CommandButton3_Click()
Unload Me
UserForm1.Show
End Sub

Private Sub CommandButton4_Click()
 Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    
    Dim StartDate As Date
    Dim EndDate As Date
    Dim TotalMonths As Long
    Dim PassedMonths As Long
    Dim RemainingMonths As Long
    Dim DeductionValue As Double
    Dim MonthlyValue As Double
    Dim Category As String
    Dim RiskLevel As String
    Dim pctRemaining As Double
    
    Set ws = ThisWorkbook.Worksheets("Financialsector")
    
    Application.ScreenUpdating = False
    
    lastRow = ws.Cells(ws.Rows.count, "D").End(xlUp).row
    
    For r = lastRow To 2 Step -1
        
        Category = Trim(ws.Cells(r, "D").value)       ' ›∆… «·«” ﬁÿ«⁄
        DeductionValue = val(ws.Cells(r, "F").value)  ' ﬁÌ„… «·«” ﬁÿ«⁄ «·ﬂ·Ì…
        TotalMonths = val(ws.Cells(r, "H").value)     ' „œ… «·«” ﬁÿ«⁄ »«·√‘Â—
        
        ' ===== Õ–› „»·€ «·€Ì«» =====
        If Category = "„»·€ «·€Ì«»" Then
            ws.Rows(r).Delete
            GoTo NextRow
        End If
        
        ' «· Õﬁﬁ „‰  «—ÌŒ «·»œ¡
        If Not IsDate(ws.Cells(r, "G").value) Then GoTo NextRow
        StartDate = ws.Cells(r, "G").value
        
        ' ===============================
        ' Õ”«»  «—ÌŒ «·«‰ Â«¡
        ' ===============================
        EndDate = DateAdd("m", TotalMonths, StartDate)
        EndDate = DateAdd("d", -1, EndDate)
        ws.Cells(r, "I").value = EndDate
        ws.Cells(r, "I").NumberFormat = "dd/mm/yyyy"
        
        ' ===============================
        ' Õ”«» «·√‘Â— «·„‰ﬁ÷Ì… Ê«·„ »ﬁÌ…
        ' ===============================
        If Date < StartDate Then
            PassedMonths = 0
        Else
            PassedMonths = DateDiff("m", StartDate, Date)
            If Date < DateAdd("m", PassedMonths, StartDate) Then
                PassedMonths = PassedMonths - 1
                 ws.Cells(r, "p").value = PassedMonths
            End If
        End If
        
        If PassedMonths < 0 Then PassedMonths = 0
        If PassedMonths > TotalMonths Then PassedMonths = TotalMonths
        
        RemainingMonths = TotalMonths - PassedMonths
        If RemainingMonths < 0 Then RemainingMonths = 0
        
        ws.Cells(r, "J").value = PassedMonths      ' «·√‘Â— «·„Õ ”»…
         ws.Cells(r, "p").value = ws.Cells(r, "h") - ws.Cells(r, "j")
        ' ===============================
        ' «Õ ”«» «·ﬁÌ„… «·‘Â—Ì…
        ' ===============================
        Select Case Category
            Case "«·‰›ﬁ…", "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%"
                MonthlyValue = DeductionValue
            Case Else
                If TotalMonths > 0 Then
                    MonthlyValue = DeductionValue / TotalMonths
                Else
                    MonthlyValue = 0
                End If
        End Select
        ws.Cells(r, "K").value = MonthlyValue
        
        ' ===============================
        ' «·⁄„Êœ M = «·ﬁÌ„… «· —«ﬂ„Ì…
        ' ===============================
        Select Case Category
            Case "«·‰›ﬁ…", "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%"
                ws.Cells(r, "M").value = DeductionValue
            Case Else
                ws.Cells(r, "M").value = MonthlyValue * PassedMonths
        End Select
        
        ' ===============================
        ' «·⁄„Êœ N ? „·«ÕŸ« 
        ' ===============================
        Select Case Category
            Case "«·‰›ﬁ…"
                ws.Cells(r, "N").value = "«·‰›ﬁ… „” „—… Õ Ï ’œÊ— √„— ≈œ«—Ì »≈Ìﬁ«›Â«"
                ws.Cells(r, "L").value = ""  ' ·« Õ«·…
            Case "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%"
                ws.Cells(r, "N").value = "Â–« «·«” ﬁÿ«⁄ «· ﬁ«⁄œÌ ·« Ì Ã“√ ÊÌ” „— Õ Ï ’œÊ— √„— ≈œ«—Ì »⁄œ «·≈Ã«“…"
                ws.Cells(r, "L").value = ""  ' ·« Õ«·…
            Case Else
                ' »«ﬁÌ «·›∆«  ? Õ«·… «·«” ﬁÿ«⁄
                If DeductionValue = ws.Cells(r, "N").value Or RemainingMonths > 0 Then
                    ws.Cells(r, "N").value = "«·«” ﬁÿ«⁄ ·« Ì“«· „” „—"
                End If
                
                ' Õ«·… ﬁ—» «·«‰ Â«¡ ›Ì «·⁄„Êœ L
                If RemainingMonths <= 1 And RemainingMonths > 0 Then
                    ws.Cells(r, "L").value = "«·”·›… ‘«—›  ⁄·Ï «·«‰ Â«¡"
                ElseIf RemainingMonths = 0 Then
                    ws.Cells(r, "L").value = "«‰ ÂÏ «·«” ﬁÿ«⁄"
                Else
                    ws.Cells(r, "L").value = ""
                End If
        End Select
        
        ' ===============================
        ' «·⁄„Êœ O ? „” ÊÏ «·„Œ«ÿ—
        ' ===============================
        Select Case Category
            Case "«·‰›ﬁ…"
                RiskLevel = "„” „— Õ Ï ’œÊ— √„— ≈œ«—Ì"
            Case "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%"
                If RemainingMonths = 0 Then
                    RiskLevel = "«·«” ﬁÿ«⁄ «· ﬁ«⁄œÌ «ﬂ „·° √»·€ «·≈œ«—…"
                Else
                    RiskLevel = "À«»  Õ Ï ‰Â«Ì… «·«· “«„"
                End If
            Case Else
                ' »«ﬁÌ «·›∆«  «·⁄«œÌ… ? Õ”«»  œ—Ã «·„Œ«ÿ—
                If TotalMonths > 0 Then
                    pctRemaining = RemainingMonths / TotalMonths
                    Select Case pctRemaining
                        Case Is >= 0.75
                            RiskLevel = "„— ›⁄ Ãœ«"
                        Case Is >= 0.5
                            RiskLevel = "„— ›⁄"
                        Case Is >= 0.25
                            RiskLevel = "„ Ê”ÿ"
                        Case Is > 0
                            RiskLevel = "„‰Œ›÷"
                        Case Else
                            RiskLevel = "«‰ ÂÏ «·«” ﬁÿ«⁄"
                    End Select
                Else
                    RiskLevel = "€Ì— „Õœœ"
                End If
        End Select
        
        ws.Cells(r, "O").value = RiskLevel
        
NextRow:
    Next r
    
    Application.ScreenUpdating = True
    
    MsgBox " „  ÕœÌÀ Ã„Ì⁄ «·»Ì«‰«  Ê«·«Õ ”«»«  »‰Ã«Õ.", vbInformation

End Sub

Private Sub Label10_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub TextBox1_Change() ' ⁄‰œ ≈œŒ«· «·—ﬁ„ «·ÊŸÌ›Ì
    Call LoadEmployeeData
End Sub







Private Sub TextBox13_AfterUpdate()

    
    
End Sub

Private Sub TextBox13_Change()
Call LoadEmployeeByPartialSearch(TextBox13.Text)
End Sub



Private Sub TextBox14_Change()

End Sub



Private Sub TextBox15_Change()

End Sub

Private Sub TextBox5_Change()
    Call UpdateRemainingAmountAndMonths
End Sub





Private Sub TextBox6_Change()
  
    Call UpdateMonthsPassed
    Call UpdateRemainingAmountAndMonths
    Call UpdateRemainingMonths
    Call Calc_Retirement_Stop
End Sub



Private Sub TextBox7_Change()
 
    Call UpdateMonthsPassed
    Call UpdateRemainingAmountAndMonths
    Call UpdateRemainingMonths
    Call Calc_Retirement_Stop
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
ComboBox1.List = Array( _
        "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%", _
        "«·‰›ﬁ…", _
        "«·«” ﬁÿ«⁄ 1", _
        "«·«” ﬁÿ«⁄ «·„ ›—ﬁ", _
        "„»·€ «·€Ì«»", _
        "«·”·›…", _
        "«·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ ", _
        "«” ﬁÿ«⁄ ⁄«„", _
        "«” ﬁÿ«⁄ „»·€ Õ”» «·‰”»…")

    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim lvwItem As ListItem

    Set ws = ThisWorkbook.Sheets("Financialsector")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row

    '  ÂÌ∆… ListView
    With ListView1
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .ListItems.Clear
        .ColumnHeaders.Clear
        
        ' ≈÷«›… ﬂ· «·√⁄„œ… »œÊ‰  ÕœÌœ «·„Õ«–«…
        .ColumnHeaders.Add , , "—ﬁ„ «·«” ﬁÿ«⁄", 70
        .ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 120
        .ColumnHeaders.Add , , "«·«”„", 180
        .ColumnHeaders.Add , , "›∆… «·«” ﬁÿ«⁄", 120
        .ColumnHeaders.Add , , "‰Ê⁄ «·«” ﬁÿ«⁄", 100
        .ColumnHeaders.Add , , "ﬁÌ„… «·«” ﬁÿ«⁄", 90
        .ColumnHeaders.Add , , " «—ÌŒ «·»œ¡", 90
        .ColumnHeaders.Add , , "„œ… (‘Â—)", 70
        .ColumnHeaders.Add , , " «—ÌŒ «·«‰ Â«¡", 90
        .ColumnHeaders.Add , , "«·√‘Â— «·„Õ ”»…", 90
        .ColumnHeaders.Add , , "«·ﬁÌ„… «·‘Â—Ì…", 90
        .ColumnHeaders.Add , , "«·Õ«·…", 90
        .ColumnHeaders.Add , , "«·„»·€ «· —«ﬂ„Ì", 80
        .ColumnHeaders.Add , , "«·„·«ÕŸ« ", 150
        .ColumnHeaders.Add , , "„” ÊÏ «·„Œ«ÿ—", 120
    End With
    
    '  ⁄»∆… «·»Ì«‰«  ›Ì ListView
    For r = 2 To lastRow
        Set lvwItem = ListView1.ListItems.Add(, , ws.Cells(r, "A").Text)
        lvwItem.SubItems(1) = ws.Cells(r, "B").Text
        lvwItem.SubItems(2) = ws.Cells(r, "C").Text
        lvwItem.SubItems(3) = ws.Cells(r, "D").Text
        lvwItem.SubItems(4) = ws.Cells(r, "E").Text
        lvwItem.SubItems(5) = ws.Cells(r, "F").Text
        lvwItem.SubItems(6) = ws.Cells(r, "G").Text
        lvwItem.SubItems(7) = ws.Cells(r, "H").Text
        lvwItem.SubItems(8) = ws.Cells(r, "I").Text
        lvwItem.SubItems(9) = ws.Cells(r, "J").Text
        lvwItem.SubItems(10) = ws.Cells(r, "K").Text
        lvwItem.SubItems(11) = ws.Cells(r, "L").Text
        lvwItem.SubItems(12) = ws.Cells(r, "M").Text
       lvwItem.SubItems(13) = ws.Cells(r, "N").Text
        lvwItem.SubItems(14) = ws.Cells(r, "O").Text
         
    Next r


End Sub


Private Sub UpdateRemainingAmountAndMonths()
    Dim StartDate As Date
    Dim totalAmount As Double
    Dim TotalMonths As Long
    Dim monthlyDeduction As Double
    Dim MonthsPassed As Long
    Dim previousMonths As Long
    Dim consumedAmount As Double
    Dim remainingAmount As Double
    Dim RemainingMonths As Long
    Dim EndDate As Date
    
    ' «· Õﬁﬁ „‰ ’Õ… «·„œŒ·« 
    If Not IsDate(Me.TextBox6.value) Then Exit Sub
    If Not IsNumeric(Me.TextBox5.value) Then Exit Sub
    If Not IsNumeric(Me.TextBox7.value) Then Exit Sub
    
    ' ﬁ—«¡… «·ﬁÌ„ „‰ TextBoxes
    StartDate = CDate(Me.TextBox6.value)
    totalAmount = CDbl(Me.TextBox5.value)
    TotalMonths = CLng(Me.TextBox7.value)
    
    If TotalMonths <= 0 Then TotalMonths = 1
    
    ' Õ”«»  «—ÌŒ «·«‰ Â«¡  ·ﬁ«∆Ì«
    EndDate = DateAdd("m", TotalMonths - 1, StartDate)
    Me.TextBox8.value = Format(EndDate, "dd/mm/yyyy")
    
    ' Õ”«» «·ﬁÌ„… «·‘Â—Ì… ·ﬂ· «” ﬁÿ«⁄
    monthlyDeduction = totalAmount / TotalMonths
    Me.TextBox10.value = monthlyDeduction
    
    ' «·√‘Â— «·„‰ﬁ÷Ì… „‰–  «—ÌŒ «·»œ¡ Õ Ï «·ÌÊ„
    If Date >= StartDate Then
        MonthsPassed = DateDiff("m", StartDate, Date) + 1
        If MonthsPassed > TotalMonths Then MonthsPassed = TotalMonths
    Else
        MonthsPassed = 0
    End If
    
    ' «·√‘Â— «·„Õ ”»… ”«»ﬁ« „‰ TextBox9
    If IsNumeric(Me.TextBox9.value) Then
        previousMonths = CLng(Me.TextBox9.value)
    Else
        previousMonths = 0
    End If
    
    ' Õ”«» «·„»·€ «·„” Â·ﬂ Õ Ï «·ÌÊ„ „⁄ «·√‘Â— «·”«»ﬁ…
    consumedAmount = monthlyDeduction * MonthsPassed
    
    ' «·„»·€ «·„ »ﬁÌ
    remainingAmount = totalAmount - consumedAmount
    If remainingAmount < 0 Then remainingAmount = 0
    Me.TextBox14.value = remainingAmount
    
    ' Õ”«» «·√‘Â— «·„ »ﬁÌ… »⁄œ ÿ—Õ «·√‘Â— «·„Õ ”»… ”«»ﬁ«
    RemainingMonths = TotalMonths - (MonthsPassed + previousMonths)
    If RemainingMonths < 0 Then RemainingMonths = 0
    Me.TextBox15.value = RemainingMonths & " " & "«·«‘Â— «·„ »ﬁÌ…"
End Sub



Private Sub LoadEmployeeData()
    ' ﬂÊœ Ã·» «·»Ì«‰«  „‰ «·‘Ì  »‰«¡ ⁄·Ï TextBox1 («·—ﬁ„ «·ÊŸÌ›Ì)
    Dim ws As Worksheet
    Dim EmpID As String
    Dim lastRow As Long, r As Long
    Dim found As Boolean
    
    Set ws = ThisWorkbook.Sheets("Financialsector")
    EmpID = Trim(Me.TextBox1.Text)
    found = False
    
    If EmpID = "" Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    For r = 2 To lastRow
        If ws.Cells(r, "B").Text = EmpID Then
            found = True
            
            '  Õ„Ì· »Ì«‰«  «·„ÊŸ› ≈·Ï TextBoxes
            Me.TextBox2.value = ws.Cells(r, "C").Text
            Me.TextBox3.value = ws.Cells(r, "D").Text
            Me.TextBox4.value = ws.Cells(r, "E").Text
            Me.TextBox5.value = ws.Cells(r, "F").Text
            If IsDate(ws.Cells(r, "G").value) Then Me.TextBox6.value = Format(ws.Cells(r, "G").value, "dd/mm/yyyy")
            Me.TextBox7.value = ws.Cells(r, "H").Text
            If IsDate(ws.Cells(r, "I").value) Then Me.TextBox8.value = Format(ws.Cells(r, "I").value, "dd/mm/yyyy")
            Me.TextBox9.value = ws.Cells(r, "J").Text
           Me.TextBox10.value = ws.Cells(r, "K").Text
            Me.TextBox11.value = ws.Cells(r, "L").Text
            Me.TextBox14.value = ws.Cells(r, "M").Text
            Me.TextBox12.value = ws.Cells(r, "N").Text
            Me.TextBox16.value = ws.Cells(r, "o").Text
            '  ÕœÌÀ «·Õ”«»«   ·ﬁ«∆Ì«
           Call UpdateRemainingAmountAndMonths
            Exit For
        End If
    Next r
    
    If Not found Then
        MsgBox "·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·—ﬁ„ «·ÊŸÌ›Ì.", vbExclamation
    End If
End Sub

Private Sub LoadEmployeeByPartialSearch(ByVal txt As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim found As Boolean
    Dim EmpID As String
    
    Set ws = ThisWorkbook.Sheets("Financialsector")
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    found = False
    
    For r = 2 To lastRow
        ' «·»ÕÀ «·Ã“∆Ì: ›Ì «·«”„ √Ê «·—ﬁ„ «·ÊŸÌ›Ì
        If ws.Cells(r, "C").Text Like "*" & TextBox13.Text & "*" And ws.Cells(r, "d") = ComboBox1.value Then
           
            found = True
            EmpID = ws.Cells(r, "B").Text
            
            '  Õ„Ì· «·»Ì«‰«  ›Ì TextBoxes «·√Œ—Ï ›ﬁÿ° œÊ‰  €ÌÌ— TextBox13
            Me.TextBox1.value = EmpID
            Me.TextBox2.value = ws.Cells(r, "C").Text
            Me.TextBox3.value = ws.Cells(r, "D").Text
            Me.TextBox4.value = ws.Cells(r, "E").Text
            Me.TextBox5.value = ws.Cells(r, "F").Text
            If IsDate(ws.Cells(r, "G").value) Then Me.TextBox6.value = Format(ws.Cells(r, "G").value, "dd/mm/yyyy")
            Me.TextBox7.value = ws.Cells(r, "H").Text
            If IsDate(ws.Cells(r, "I").value) Then Me.TextBox8.value = Format(ws.Cells(r, "I").value, "dd/mm/yyyy")
            Me.TextBox9.value = ws.Cells(r, "J").Text
            Me.TextBox10.value = ws.Cells(r, "K").Text
            Me.TextBox11.value = ws.Cells(r, "L").Text
            Me.TextBox14.value = ws.Cells(r, "M").Text
            Me.TextBox12.value = ws.Cells(r, "n").Text
            Me.TextBox16.value = ws.Cells(r, "o").Text
             Me.TextBox15.value = ws.Cells(r, "p").Text
            '  ÕœÌÀ «·Õ”«»«  «·„ ﬁœ„…  ·ﬁ«∆Ì«
           ' Call UpdateRemainingAmountAndMonths
            
            Exit For
        End If
    Next r
    
    If Not found Then
        ' ≈–« ·„ Ì „ «·⁄ÀÊ— ⁄·Ï √Ì ‰ ÌÃ…° Ì„ﬂ‰ „”Õ TextBoxes «·√Œ—Ï
        Me.TextBox1.value = ""
        Me.TextBox2.value = ""
        Me.TextBox3.value = ""
        Me.TextBox4.value = ""
        Me.TextBox5.value = ""
        Me.TextBox6.value = ""
        Me.TextBox7.value = ""
        Me.TextBox8.value = ""
        Me.TextBox9.value = ""
        Me.TextBox10.value = ""
        Me.TextBox11.value = ""
        Me.TextBox12.value = ""
        Me.TextBox14.value = ""
        Me.TextBox15.value = ""
    End If
End Sub

Private Sub UpdateMonthsPassed()
    Dim StartDate As Date
    Dim TotalMonths As Long
    Dim MonthsPassed As Long
    Dim daysDiff As Long
    Dim fractionMonth As Double
    
    ' «· Õﬁﬁ „‰ ’Õ… «·„œŒ·« 
    If Not IsDate(Me.TextBox6.value) Then Exit Sub
    If Not IsNumeric(Me.TextBox7.value) Then Exit Sub
    
    StartDate = CDate(Me.TextBox6.value)
    TotalMonths = CLng(Me.TextBox7.value)
    
    If TotalMonths <= 0 Then TotalMonths = 1
    
    ' Õ”«» «·√‘Â— «·„‰ﬁ÷Ì… „‰–  «—ÌŒ «·»œ¡ Õ Ï «·ÌÊ„
    If Date < StartDate Then
        MonthsPassed = 0
    Else
        daysDiff = Date - StartDate
        fractionMonth = daysDiff / 30.44       ' „ Ê”ÿ √Ì«„ «·‘Â—
        MonthsPassed = Int(fractionMonth) + 1  ' +1 ··‘Â— «·Õ«·Ì
        If MonthsPassed > TotalMonths Then MonthsPassed = TotalMonths
    End If
    
    ' Ê÷⁄ «·‰ ÌÃ… ›Ì TextBox9
    Me.TextBox9.value = MonthsPassed
End Sub

Private Sub UpdateRemainingMonths()
    Dim TotalMonths As Long
    Dim MonthsPassed As Long
    Dim RemainingMonths As Long
    
    ' «· Õﬁﬁ „‰ ’Õ… «·„œŒ·« 
    If Not IsNumeric(Me.TextBox7.value) Then Exit Sub
    If Not IsNumeric(Me.TextBox9.value) Then Exit Sub
    
    TotalMonths = CLng(Me.TextBox7.value)
    MonthsPassed = CLng(Me.TextBox9.value)
    
    RemainingMonths = TotalMonths - MonthsPassed
    If RemainingMonths < 0 Then RemainingMonths = 0
    
    ' Ê÷⁄ «·‰ ÌÃ… ›Ì TextBox15
    Me.TextBox15.value = RemainingMonths
End Sub

Private Sub Calc_Retirement_Stop()

    ' ›ﬁÿ ·›∆… «· ÊﬁÌ›«  «· ﬁ«⁄œÌ…
    If Me.TextBox3.value <> "«· ÊﬁÌ›«  «· ﬁ«⁄œÌ… 15%" Then Exit Sub
    
    ' «· Õﬁﬁ „‰  «—ÌŒ «·»œ¡
    If Not IsDate(Me.TextBox6.value) Then Exit Sub
    
    Dim StartDate As Date
    Dim PassedMonths As Long
    
    StartDate = CDate(Me.TextBox6.value)
    
    ' Õ”«» «·√‘Â— «·„‰ﬁ÷Ì… „‰  «—ÌŒ «·»œ¡ ≈·Ï «·ÌÊ„
    PassedMonths = DateDiff("m", StartDate, Date)
    
    ' ﬁ»· „—Ê— 5 ”‰Ê«  ? «·„»·€ «·„ »ﬁÌ Ì”«ÊÌ ﬁÌ„… «·«” ﬁÿ«⁄
    If PassedMonths < 60 Then
        Me.TextBox14.value = Me.TextBox5.value
        Me.TextBox12.value = Me.TextBox5.value
        Exit Sub
    End If
    
    ' »⁄œ „—Ê— 5 ”‰Ê«  ?  ’›Ì— «·«” ﬁÿ«⁄
    If PassedMonths >= 60 Then
       
        Me.TextBox10.value = 0
        Me.TextBox12.value = 0
        Me.TextBox14.value = 0
    End If

End Sub

