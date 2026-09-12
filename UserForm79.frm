VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm79 
   Caption         =   "UserForm79"
   ClientHeight    =   13572
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   14784
   OleObjectBlob   =   "UserForm79.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm79"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Option Explicit

' ==============================================================================
' ‘«‘… «·—ﬁ«»… Ê«·„ÿ«»ﬁ… «·„—ﬂ“Ì… ·—Ê« » «·„ÊŸ›Ì‰ (»Ì‰ ‘Â—Ì‰) - «·≈’œ«— «·„” ﬁ·
' «·„ﬂ«‰: ÌÊ÷⁄ »«·ﬂ«„· œ«Œ· ‰«›–… √ﬂÊ«œ (UserForm79)
' «·Ê’›:  „  ⁄œÌ· ‰“Ê· ‘—Ìÿ «·⁄‰«ÊÌ‰ «·—∆Ì”Ì ·Ìœ√ „‰ «·ﬁ„… 60  „«„«°
'        Ê »œ√ «·»‰Êœ «·„«·Ì… »«· »⁄Ì… „‰ «·ﬁ„… 105 „⁄ ‰“Ê· 60 »ﬂ”· »Ì‰ «·»‰Êœ.
' ==============================================================================

Private Type MonthSummary
    CurrentAllowances(1 To 14) As Double
    PastAllowances(1 To 14) As Double
    CurrentDeductions(1 To 16) As Double
    PastDeductions(1 To 16) As Double
    CurrentNetSalary As Double
    PastNetSalary As Double
End Type

Private Sub UserForm_Activate()
Call DesignAndFillAuditSheet
End Sub

Private Sub UserForm_Initialize()
    Dim wsCurrent As Worksheet, wsPast As Worksheet, wsReport As Worksheet
    Dim lrCurrent As Long, lrPast As Long
    Dim arrCurrent As Variant, arrPast As Variant
    Dim AuditData As MonthSummary
    Dim AllowIdx As Variant, DeductIdx As Variant
    Dim i As Long, j As Integer, r As Long
    Dim Diff As Double, NetColIdx As Long
    Dim headerRow As Long
    
    ' „ €Ì—«  «· Ê·Ìœ «·œÌ‰«„ÌﬂÌ ·ﬂ«∆‰«  «·Ê«ÃÂ…
    Dim lblItem As MSForms.Label
    Dim txtCurrent As MSForms.TextBox
    Dim txtPast As MSForms.TextBox
    Dim txtDiff As MSForms.TextBox
    Dim lblSection As MSForms.Label
    Dim tTop As Long
    Dim itemVal As String
    
    ' „ €Ì—«  „”„Ì«  «·√‘Â— «·œÌ‰«„ÌﬂÌ… «·„” ‰œ… ·”«⁄… «·‰Ÿ«„
    Dim CurrentMonthName As String
    Dim PastMonthName As String
    
    ' ÷»ÿ «·ÀÊ«»  Ê«·„ƒ‘—«  «·ÂÌﬂ·Ì… ··—Ê« »
    NetColIdx = 49 ' ⁄„Êœ ’«›Ì «·—« » AW
    headerRow = 8  ' ”ÿ— «·⁄‰«ÊÌ‰ Ê—ƒÊ” «·√⁄„œ… «·ÕﬁÌﬁÌ…
    
    ' «” Œ—«Ã √”„«¡ «·√‘Â—  ·ﬁ«∆Ì« „‰ ‰Ÿ«„ «·ÊÌ‰œÊ“
    CurrentMonthName = Format(Date, "mmmm yyyy")
    PastMonthName = Format(DateAdd("m", -1, Date), "mmmm yyyy")
    
    ' 1. «· Õﬁﬁ „‰ ÊÃÊœ «·‘Ì «  «·√”«”Ì… Ê«·„” Âœ›… ›Ì «·„·›
    On Error Resume Next
    Set wsCurrent = ThisWorkbook.Sheets("Sheet1")
    Set wsPast = ThisWorkbook.Sheets("past_month76")
    Set wsReport = ThisWorkbook.Sheets("between_two_month")
    On Error GoTo 0
    
    If wsCurrent Is Nothing Or wsPast Is Nothing Or wsReport Is Nothing Then
        MsgBox "Œÿ√:  √ﬂœ „‰ ÊÃÊœ «·‘Ì «  «·„—Ã⁄Ì… (Sheet1) Ê (past_month55) Ê‘Ì  «· ﬁ—Ì— «·›«—€ (between_two_month).", vbCritical, "›‘·  ÂÌ∆… «·Ê«ÃÂ…"
        Unload Me
        Exit Sub
    End If
    
    ' 2. ”Õ» Ê„⁄«·Ã… «·»Ì«‰«  »”—⁄… ›«∆ﬁ… œ«Œ· «·–«ﬂ—… (RAM) ·—›⁄ «·√œ«¡
    lrCurrent = wsCurrent.Cells(wsCurrent.rowS.count, "AW").End(xlUp).row
    lrPast = wsPast.Cells(wsPast.rowS.count, "AW").End(xlUp).row
    If lrCurrent < 9 Then lrCurrent = 9
    If lrPast < 9 Then lrPast = 9
    
    arrCurrent = wsCurrent.Range("A1:CE" & lrCurrent).value
    arrPast = wsPast.Range("A1:CE" & lrPast).value
    
    ' √œ·… «·√⁄„œ… «·ÕﬁÌﬁÌ… ›Ì „’›Ê›… «·‹ VBA («·«” Õﬁ«ﬁ«  Ê«·«” ﬁÿ«⁄« )
    AllowIdx = Array(14, 16, 18, 19, 20, 21, 25, 27, 29, 60, 68, 70, 10, 83)
    DeductIdx = Array(24, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46, 47, 62, 67, 69, 50)
    
    ' Õ”«» „Ã«„Ì⁄ «·«” Õﬁ«ﬁ«  Ê’«›Ì «·—Ê« » ›Ì «·–«ﬂ—…
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
    
    ' Õ”«» „Ã«„Ì⁄ «·«” ﬁÿ«⁄«  Ê«·„”ÕÊ»«  ›Ì «·–«ﬂ—…
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
    
    ' 3.  ÂÌ∆… Ê≈⁄œ«œ √»⁄«œ «·›Ê—„ Ê ›⁄Ì· ‘—Ìÿ «· „—Ì— «·—√”Ì «·„ Ã«Ê»
    With Me
        .RightToLeft = True
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .ScrollBars = fmScrollBarsVertical
        .ScrollHeight = 2250 ' “Ì«œ… ÿ›Ì›… · €ÿÌ… ‰“Ê· «·ÂÌœ— «·ÃœÌœ «·»«·€ 60 »ﬂ”·
    End With
    
    ' 4. »‰«¡ ‘—Ìÿ «· —ÊÌ”… Ê⁄‰«ÊÌ‰ «·√⁄„œ… ›Ì √⁄·Ï «·›Ê—„ · »œ√ „‰ «·ﬁ„… 60  „«„«
    BuildFormHeader_Internal CurrentMonthName, PastMonthName
    
    ' ‰ﬁÿ… «‰ÿ·«ﬁ «·ﬁ”„ «·√Ê· »⁄œ «· —ÊÌ”… (60 ‰“Ê· «· —ÊÌ”… + 25 «— ›«⁄Â« + 20 „”«›… √„«‰ = 105)
    tTop = 105
    
    ' 5. «·»œ¡ ›Ì ≈‰‘«¡ «·√œÊ«  (Labels & TextBoxes) Êﬁ—«¡… «·„”„Ì«  „‰ «·”ÿ— 8
    
    ' --- √ -  Ê·Ìœ ﬁ”„ «·«” Õﬁ«ﬁ«  Ê«· ⁄ÊÌ÷«  ---
    Set lblSection = Me.Controls.Add("Forms.Label.1")
    With lblSection
        .Caption = "   ? ﬁ”‹‹‹‹„ «·«” Õﬁ«ﬁ‹‹‹‹«  «·„«·Ì‹‹‹‹… Ê«· ⁄ÊÌ÷‹‹‹‹« "
        .Font.Bold = True: .Font.Size = 11: .ForeColor = RGB(44, 82, 130)
        .BackColor = RGB(247, 250, 252): .BackStyle = fmBackStyleOpaque
        .Left = 20: .Top = tTop: .Width = 680: .Height = 24: .TextAlign = fmTextAlignRight
    End With
    tTop = tTop + 45 ' „”«›… ‰“Ê· »⁄œ ⁄‰Ê«‰ «·ﬁ”„ «·—∆Ì”Ì
    
    For j = 1 To 14
        itemVal = wsCurrent.Cells(headerRow, AllowIdx(j - 1)).value
        If Trim(itemVal) = "" Then itemVal = "«” Õﬁ«ﬁ €Ì— „”„Ï (⁄„Êœ " & Split(wsCurrent.Cells(headerRow, AllowIdx(j - 1)).Address, "$")(1) & ")"
        
        Diff = AuditData.CurrentAllowances(j) - AuditData.PastAllowances(j)
        CreateRowControls itemVal, AuditData.CurrentAllowances(j), AuditData.PastAllowances(j), Diff, tTop, False
        tTop = tTop + 60 ' «·‰“Ê· »„ﬁœ«— 60 »ﬂ”· ··»‰œ «· «·Ì
    Next j
    
    ' --- » -  Ê·Ìœ ﬁ”„ «·«” ﬁÿ«⁄«  Ê«·„”ÕÊ»«  ---
    tTop = tTop + 15
    Set lblSection = Me.Controls.Add("Forms.Label.1")
    With lblSection
        .Caption = "   ? ﬁ”‹‹‹‹„ «·«” ﬁÿ«⁄‹‹‹‹«  Ê«·„”ÕÊ»‹‹‹‹«  «· ﬁ«⁄œÌ‹… Ê«·÷—Ì»Ì‹…"
        .Font.Bold = True: .Font.Size = 11: .ForeColor = RGB(44, 82, 130)
        .BackColor = RGB(247, 250, 252): .BackStyle = fmBackStyleOpaque
        .Left = 20: .Top = tTop: .Width = 680: .Height = 24: .TextAlign = fmTextAlignRight
    End With
    tTop = tTop + 45 ' „”«›… ‰“Ê· »⁄œ ⁄‰Ê«‰ ﬁ”„ «·«” ﬁÿ«⁄« 
    
    For j = 1 To 16
        itemVal = wsCurrent.Cells(headerRow, DeductIdx(j - 1)).value
        If Trim(itemVal) = "" Then itemVal = "«” ﬁÿ«⁄ €Ì— „”„Ï (⁄„Êœ " & Split(wsCurrent.Cells(headerRow, DeductIdx(j - 1)).Address, "$")(1) & ")"
        
        Diff = AuditData.CurrentDeductions(j) - AuditData.PastDeductions(j)
        CreateRowControls itemVal, AuditData.CurrentDeductions(j), AuditData.PastDeductions(j), Diff, tTop, False
        tTop = tTop + 60 ' «·‰“Ê· »„ﬁœ«— 60 »ﬂ”· ··»‰œ «· «·Ì
    Next j
    
    ' --- Ã -  Ê·Ìœ Œ·«’… «· œﬁÌﬁ «·Œ «„Ì… (’«›Ì «·—Ê« » «·≈Ã„«·Ì «·‰Â«∆Ï) ---
    tTop = tTop + 15
    itemVal = wsCurrent.Cells(headerRow, NetColIdx).value
    If Trim(itemVal) = "" Then itemVal = "’«›Ì «·—Ê« » «·≈Ã„«·Ì «·‰Â«∆Ì (⁄„Êœ AW)"
    
    Diff = AuditData.CurrentNetSalary - AuditData.PastNetSalary
    CreateRowControls itemVal, AuditData.CurrentNetSalary, AuditData.PastNetSalary, Diff, tTop, True
End Sub

' ==============================================================================
' ≈Ã—«¡ ›—⁄Ì œ«Œ·Ì: · Ê·Ìœ «·√œÊ«  «·‰’Ì… Ê«·—ﬁ„Ì… ·ﬂ· ”ÿ— „«·Ì „⁄  ‰”Ìﬁ «·√·Ê«‰
' ==============================================================================
Private Sub CreateRowControls(CaptionName As String, CurVal As Double, PastVal As Double, DiffVal As Double, topPos As Long, IsTotalRow As Boolean)
    Dim lbl As MSForms.Label
    Dim txtCur As MSForms.TextBox
    Dim txtPast As MSForms.TextBox
    Dim txtDiff As MSForms.TextBox
    
    ' «”„ «·»‰œ «·„«·Ì
    Set lbl = Me.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = CaptionName
        .Left = 450: .Top = topPos: .Width = 250: .Height = 24: .TextAlign = fmTextAlignRight
        If IsTotalRow Then .Font.Bold = True
    End With
    
    ' Œ«‰… «·‘Â— «·Õ«·Ì
    Set txtCur = Me.Controls.Add("Forms.TextBox.1")
    With txtCur
        .Text = Format(CurVal, "#,##0.00")
        .Left = 310: .Top = topPos: .Width = 120: .Height = 24: .Locked = True: .TextAlign = fmTextAlignCenter
        .BackColor = IIf(IsTotalRow, RGB(237, 242, 247), RGB(255, 255, 255))
        If IsTotalRow Then .Font.Bold = True
    End With
    
    ' Œ«‰… «·‘Â— «·„«÷Ì
    Set txtPast = Me.Controls.Add("Forms.TextBox.1")
    With txtPast
        .Text = Format(PastVal, "#,##0.00")
        .Left = 170: .Top = topPos: .Width = 120: .Height = 24: .Locked = True: .TextAlign = fmTextAlignCenter
        .BackColor = IIf(IsTotalRow, RGB(237, 242, 247), RGB(255, 255, 255))
        If IsTotalRow Then .Font.Bold = True
    End With
    
    ' Œ«‰… «·›—ﬁ «·„«·Ì Ê«·„ƒ‘— «··Ê‰Ì
    Set txtDiff = Me.Controls.Add("Forms.TextBox.1")
    With txtDiff
        .Text = Format(DiffVal, "#,##0.00")
        .Left = 30: .Top = topPos: .Width = 120: .Height = 24: .Locked = True: .TextAlign = fmTextAlignCenter
        .Font.Bold = True
        
        ' ›—“ Ê ·ÊÌ‰ «·Œ«‰«  »‰«¡ ⁄·Ï «·Õ—ﬂ… «·„«·Ì… «·„ﬂ ‘›…
        If DiffVal > 0.009 Then
            .ForeColor = RGB(34, 139, 34)   ' “Ì«œ… (√Œ÷—)
            .BackColor = RGB(240, 253, 244)
        ElseIf DiffVal < -0.009 Then
            .ForeColor = RGB(196, 30, 58)   ' ‰ﬁ’«‰ (√Õ„—)
            .BackColor = RGB(254, 242, 242)
        Else
            .ForeColor = RGB(113, 128, 150)  ' „ ÿ«»ﬁ (—„«œÌ Œ›Ì›)
            .BackColor = RGB(248, 250, 252)
        End If
        
        If IsTotalRow Then .BackColor = RGB(226, 232, 240)
    End With
End Sub

' ==============================================================================
' ≈Ã—«¡ ›—⁄Ì œ«Œ·Ì: ·—”„  —ÊÌ”… «·⁄‰«ÊÌ‰ «·À«» … ›Ì √⁄·Ï «·›Ê—„ „‰ «·ﬁ„… 60  „«„«
' ==============================================================================
Private Sub BuildFormHeader_Internal(CurMonth As String, PastMonth As String)
    Dim h1 As MSForms.Label
    Dim LeftPositions As Variant, Widths As Variant, Captions As Variant, i As Integer
    
    LeftPositions = Array(450, 310, 170, 30)
    Widths = Array(250, 120, 120, 120)
    Captions = Array("«·»‰œ «·„” Âœ› Ê›∆… «·›Õ’ «·Õ”«»Ì", "«·‘Â— «·Õ«·Ì (" & CurMonth & ")", "«·‘Â— «·„«÷Ì (" & PastMonth & ")", "«·›—ﬁ «·„«·Ì «·‰ﬁÌ")
    
    For i = 0 To 3
        Set h1 = Me.Controls.Add("Forms.Label.1")
        With h1
            .Caption = Captions(i)
            .Left = LeftPositions(i)
            .Top = 60 '  „  ⁄œÌ· «·‰“Ê· Â‰« ·Ì’»Õ 60 »ﬂ”· »œ·« „‰ 20 · » ⁄œ «· —ÊÌ”… ⁄‰ «·”ﬁ›
            .Width = Widths(i)
            .Height = 25
            .Font.Bold = True
            .BackColor = RGB(226, 232, 240)
            .ForeColor = RGB(26, 54, 93)
            .BackStyle = fmBackStyleOpaque
            .TextAlign = fmTextAlignCenter
            With .Font: .Name = "Segoe UI": .Size = 9.5: End With
        End With
    Next i
End Sub




Private Sub CommandButton1_Click()
Unload Me
UserForm10.Show

End Sub

Private Sub CommandButton3_Click()

Dim wsSource As Worksheet
    Dim wbNew As Workbook
    Dim TargetFolderPath As String
    Dim fileName As String
    Dim FullPath As String
    Dim CurrentTimeStr As String
    
    ' 1.  ÕœÌœ «·„”«— «·„” Âœ› ⁄·Ï «·»«— ‘‰ D
    TargetFolderPath = "D:\export1\"
    
    ' 2. «· Õﬁﬁ „‰ ÊÃÊœ «·‘Ì  «·„«·Ì «·„—«œ  ’œÌ—Â
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("between_two_month")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "Œÿ√: ‘Ì  'between_two_month' €Ì— „ÊÃÊœ ›Ì Â–« «·„·› Õ«·Ì«! Ì—ÃÏ  Ê·ÌœÂ √Ê·«.", vbCritical, "›‘· «· ’œÌ—"
        Exit Sub
    End If
    
    ' 3. «· Õﬁﬁ «·–ﬂÌ „‰ ÊÃÊœ «·„Ã·œ «· ’œÌ—Ì° Ê≈‰‘«ƒÂ  ·ﬁ«∆Ì« ≈–« ﬂ«‰ „›ﬁÊœ«
    If Dir(TargetFolderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir "D:\export1"
        On Error GoTo 0
        ' ≈⁄«œ… «· Õﬁﬁ ·· √ﬂœ „‰ ‰Ã«Õ ≈‰‘«¡ «·„Ã·œ √Ê ’·«ÕÌ«  «·Ê’Ê· ··‹ D
        If Dir(TargetFolderPath, vbDirectory) = "" Then
            MsgBox "Œÿ√ Õ—Ã:  ⁄–— «·Ê’Ê· √Ê ≈‰‘«¡ «·„Ã·œ 'export1' ⁄·Ï «·»«— ‘‰ D." & vbCrLf & _
                   "Ì—ÃÏ «· √ﬂœ „‰ ’·«ÕÌ«  «·‰Ÿ«„ √Ê ÊÃÊœ «·»«— ‘‰ D ⁄·Ï ÃÂ«“ﬂ.", vbCritical, "Œÿ√ ›Ì „”«— «·Õ›Ÿ"
            Exit Sub
        End If
    End If
    
    ' 4. ’Ì«€… «”„ „·› œÌ‰«„ÌﬂÌ „—‰ (ÌÕ ÊÌ ⁄·Ï «”„ «· ﬁ—Ì— Ê«·Êﬁ  «·Õ«·Ì ·„‰⁄  œ«Œ· «·‰”Œ)
    CurrentTimeStr = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    fileName = " ﬁ—Ì—_„ÿ«»ﬁ…_«·—Ê« »_" & CurrentTimeStr & ".xlsx"
    FullPath = TargetFolderPath & fileName
    
    '  ”—Ì⁄ «·⁄„·Ì… Ê≈Ìﬁ«›  –»–» «·‘«‘… √À‰«¡ «· ’œÌ—
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 5. √Œ– ‰”Œ… „‰›’·… „‰ «·‘Ì  Ê‰ﬁ·Â« ≈·Ï „’‰› (Workbook) ÃœÌœ  „«„«
    wsSource.Copy
    Set wbNew = ActiveWorkbook
    
    ' 6. Õ›Ÿ „·› «·≈ﬂ”Ì· «·ÃœÌœ ›Ì «·„”«— «·„Õœœ Ê≈€·«ﬁÂ »√„«‰
    With wbNew
        .SaveAs fileName:=FullPath, FileFormat:=xlOpenXMLWorkbook
        .Close SaveChanges:=False
    End With
    
    ' ≈⁄«œ…  ‘€Ì·  ÕœÌÀ«  «·‘«‘… Ê«·—”«∆·
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    ' 7. ≈‘⁄«— «·„” Œœ„ »‰Ã«Õ «·⁄„·Ì… Ê≈ŸÂ«— «·„”«— «·ﬂ«„·
    MsgBox " „  ’œÌ— „·› «·„ÿ«»ﬁ… Ê«·„—«Ã⁄… »‰Ã«Õ!" & vbCrLf & _
           "«”„ «·„·›: " & fileName & vbCrLf & _
           "«·„”«—: " & TargetFolderPath, vbInformation, "«ﬂ „· «· ’œÌ— »√„«‰"

End Sub


