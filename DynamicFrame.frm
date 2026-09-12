VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DynamicFrame 
   Caption         =   "UserForm79"
   ClientHeight    =   10944
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16560
   OleObjectBlob   =   "DynamicFrame.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DynamicFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- 1.  ⁄—Ì› «·ÂÌ«ﬂ· Ê«·„’›Ê›«  ⁄·Ï „” ÊÏ «·›Ê—„ ·Õ›Ÿ «” ﬁ·«·Ì… «·ﬂÊœ ---
Private Type MonthSummary
    CurrentMonthName As String
    PastMonthName As String
    CurrentAllowances(1 To 14) As Double
    PastAllowances(1 To 14) As Double
    CurrentDeductions(1 To 16) As Double
    PastDeductions(1 To 16) As Double
    CurrentNetSalary As Double
    PastNetSalary As Double
End Type

Private AuditData As MonthSummary
Private AllowCols As Variant
Private DeductCols As Variant
Private AllowIdx As Variant
Private DeductIdx As Variant
Private DynamicFrame As Object '  „  ÕÊÌ·Â« ≈·Ï ﬂ«∆‰ ⁄«„ · ›«œÌ  ÷«—» «·„ﬂ »«  »—„ÃÌ«

' --- 2. ÕœÀ «· ÂÌ∆… «·„—ﬂ“Ì ---
Private Sub UserForm_Initialize()
    ' √:  ÂÌ∆… √»⁄«œ Ê„ŸÂ— «·Ê«ÃÂ… · œ⁄„ «·»Ì∆… «·—ﬁ«»Ì… »«··€… «·⁄—»Ì…
    Me.Caption = "„‰ŸÊ„… «·—ﬁ«»… «·„—ﬂ“Ì… Ê„ÿ«»ﬁ… «·—Ê« » - «·‰”Œ… «·„” ﬁ·… «·„” œ«„…"
    Me.RightToLeft = True
    Me.Width = 650
    Me.Height = 580
    Me.BackColor = RGB(245, 247, 250)
    
    ' »: „’›Ê›«  √”„«¡ «·√⁄„œ… (··⁄—÷ ›Ì «·Ê«ÃÂ…)
    AllowCols = Array("N", "P", "R", "S", "T", "U", "Y", "AA", "AC", "BH", "BP", "BR", "J", "CE")
    DeductCols = Array("X", "AF", "AG", "AH", "AI", "AJ", "AL", "AN", "AP", "AR", "AT", "AU", "BJ", "BO", "BQ", "AX")
    
    ' Ã: „’›Ê›«  √—ﬁ«„ «·√⁄„œ… «·À«» … «·„ﬁ«»·… ··Õ—Ê›
    AllowIdx = Array(14, 16, 18, 19, 20, 21, 25, 27, 29, 60, 68, 70, 10, 83)
    DeductIdx = Array(24, 32, 33, 34, 35, 36, 38, 40, 42, 44, 46, 47, 62, 67, 69, 50)
    
    ' œ: «” œ⁄«¡ „Õ—ﬂ „⁄«·Ã… «·»Ì«‰«  „‰ «·‘Ì «  ≈·Ï «·–«ﬂ—… (RAM Engine)
    LoadAndProcessSalaryData
    
    ' Â‹: ≈‰‘«¡ «·⁄‰«ÊÌ‰ «· ›”Ì—Ì… «·—∆Ì”Ì… ··√⁄„œ… »—„ÃÌ«
    CreateHeaderLabels
    
    ' Ê: ≈‰‘«¡ ≈ÿ«— «· „—Ì— «·—√”Ì «·–ﬂÌ (Scrollable Frame) ·«” Ì⁄«» «·√Ê»Ãﬂ ”
    Set DynamicFrame = Me.Controls.Add("Forms.Frame.1", "DynamicFrame", True)
    With DynamicFrame
        .Top = 45
        .Left = 15
        .Width = 605
        .Height = 420
        .BackColor = RGB(255, 255, 255)
        .Caption = ""
        
        ' «” Œœ«„ „⁄«·Ã √Œÿ«¡ „Ê÷⁄Ì · ›«œÌ ﬁÌÊœ »⁄÷ ≈’œ«—«  ≈ﬂ”· √À‰«¡ «· „—Ì—
        On Error Resume Next
        .ScrollBars = 2 ' 2  ⁄‰Ì ‘—Ìÿ  „—Ì— ⁄„ÊœÌ À«»‹ 
        .KeepScrollBarsVisible = 2
        On Error GoTo 0
    End With
    
    ' “: »œ¡ «· Ê·Ìœ «· ·ﬁ«∆Ì ·Ã„Ì⁄ «·√Ê»Ãﬂ ” Ê«·„ƒ‘—«  «·»Ì«‰Ì… Õ”«”… «·≈‘«—…
    BuildAuditControls
    
    ' Õ: »‰«¡ «·Œ·«’… «·≈Ã„«·Ì… «·”›·Ì… ·’«›Ì «·—« » AW
    BuildFooterSummary
End Sub

' --- 3. ŒÊ«—“„Ì… „⁄«·Ã… Êﬁ—«¡… «·»Ì«‰«  œ«Œ· «·–«ﬂ—… ·÷„«‰ «·”—⁄… «·›«∆ﬁ… ---
Private Sub LoadAndProcessSalaryData()
    Dim wsCurrent As Worksheet, wsPast As Worksheet
    Dim lrCurrent As Long, lrPast As Long
    Dim arrCurrent As Variant, arrPast As Variant
    Dim i As Long, j As Integer
    Dim NetColIdx As Long
    
    NetColIdx = 49 ' ⁄„Êœ AW
    
    On Error Resume Next
    Set wsCurrent = ThisWorkbook.Sheets("Sheet1")
    Set wsPast = ThisWorkbook.Sheets("past_month55")
    On Error GoTo 0
    
    If wsCurrent Is Nothing Or wsPast Is Nothing Then
        MsgBox "Œÿ√ ›Ì »‰Ì… «·„·›:  √ﬂœ „‰ ÊÃÊœ «·‘Ì «  (Sheet1) Ê (past_month55) ·⁄„· «·„‰ŸÊ„….", vbCritical, "‰Ÿ«„ «·—ﬁ«»… «·„” ﬁ·"
        End
    End If
    
    AuditData.CurrentMonthName = Format(Date, "mmmm yyyy", vbUseSystemDayOfWeek, vbUseSystem)
    AuditData.PastMonthName = Format(DateAdd("m", -1, Date), "mmmm yyyy", vbUseSystemDayOfWeek, vbUseSystem)
    
    lrCurrent = wsCurrent.Cells(wsCurrent.rowS.count, "AW").End(xlUp).row
    lrPast = wsPast.Cells(wsPast.rowS.count, "AW").End(xlUp).row
    
    If lrCurrent < 9 Then lrCurrent = 9
    If lrPast < 9 Then lrPast = 9
    
    ' ﬁ—«¡… «·‰ÿ«ﬁ«  ≈·Ï «·–«ﬂ—… · ”—Ì⁄ «·√œ«¡ [cite: user_context]
    arrCurrent = wsCurrent.Range("A1:CE" & lrCurrent).value
    arrPast = wsPast.Range("A1:CE" & lrPast).value
    
    For j = 1 To 14: AuditData.CurrentAllowances(j) = 0: AuditData.PastAllowances(j) = 0: Next j
    For j = 1 To 16: AuditData.CurrentDeductions(j) = 0: AuditData.PastDeductions(j) = 0: Next j
    AuditData.CurrentNetSalary = 0: AuditData.PastNetSalary = 0
    
    ' Õ”«» „Ã«„Ì⁄ «·«” Õﬁ«ﬁ« 
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
    
    ' Õ”«» „Ã«„Ì⁄ «·«” ﬁÿ«⁄« 
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
End Sub

' --- 4. „Õ—ﬂ »‰«¡ «· —ÊÌ”… «·⁄·Ì« ··„ÿ«»ﬁ… ---
Private Sub CreateHeaderLabels()
    Dim Titles As Variant, Widths As Variant, Lefts As Variant, i As Integer
    Dim lbl As MSForms.Label
    
    Titles = Array("»‰œ Ê„” ‰œ «·⁄„Êœ", "«·‘Â— «·Õ«·Ì", "«·‘Â— «·„«÷Ì", "«·›—ﬁ «·„«·Ì", "„ƒ‘— Õ«·… «·Õ—ﬂ… «·„Ê«“Ì…")
    Lefts = Array(15, 140, 250, 360, 470)
    Widths = Array(120, 100, 100, 100, 130)
    
    For i = 0 To 4
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblHeader" & i, True)
        With lbl
            .Caption = Titles(i)
            .Left = Lefts(i)
            .Top = 15
            .Width = Widths(i)
            .Height = 20
            .Font.Name = "Segoe UI"
            .Font.Bold = True
            .Font.Size = 10
            .ForeColor = RGB(60, 70, 90)
            .Alignment = 1
        End With
    Next i
End Sub

' --- 5. ŒÊ«—“„Ì…  Ê“Ì⁄ «·√Ê»Ãﬂ ” Ê«·›’· »Ì‰ √ﬁ”«„ «·Õ”«»«  ---
Private Sub BuildAuditControls()
    Dim topPos As Long, ItemIndex As Long
    Dim i As Integer, DiffValue As Double
    Dim CurrentVal As Double, PastVal As Double
    
    topPos = 10
    ItemIndex = 1
    
    ' √: »‰«¡ ﬁ”„ «·«” Õﬁ«ﬁ«  (14 ⁄„Êœ „«·Ì „ ”·”·)
    CreateSectionDivider "ﬁ”‹‹‹‹„ «·«” Õﬁ«ﬁ‹‹‹‹«  «·„«·Ì‹‹‹‹… «·„»«‘—…", topPos
    
    For i = 1 To 14
        CurrentVal = AuditData.CurrentAllowances(i)
        PastVal = AuditData.PastAllowances(i)
        DiffValue = CurrentVal - PastVal
        
        CreateRowUnit "«” Õﬁ«ﬁ («·⁄„Êœ " & AllowCols(i - 1) & ")", CurrentVal, PastVal, DiffValue, topPos, ItemIndex
        ItemIndex = ItemIndex + 1
    Next i
    
    ' »: »‰«¡ ﬁ”„ «·«” ﬁÿ«⁄«  (16 ⁄„Êœ „«·Ì „ ”·”·)
    CreateSectionDivider "ﬁ”‹‹‹‹„ «·«” ﬁÿ«⁄‹‹‹‹«  Ê«·„”ÕÊ»‹‹‹‹«  «·„«·Ì…", topPos
    
    For i = 1 To 16
        CurrentVal = AuditData.CurrentDeductions(i)
        PastVal = AuditData.PastDeductions(i)
        DiffValue = CurrentVal - PastVal
        
        CreateRowUnit "«” ﬁÿ«⁄ («·⁄„Êœ " & DeductCols(i - 1) & ")", CurrentVal, PastVal, DiffValue, topPos, ItemIndex
        ItemIndex = ItemIndex + 1
    Next i
    
    '  ⁄ÌÌ‰ «·„œÏ «·√ﬁ’Ï ·· „—Ì— «·—√”Ì ··Õ«ÊÌ… »‰«¡ ⁄·Ï «·ÕÃ„ «·„ Ê·œ ¬·Ì«
    On Error Resume Next
    DynamicFrame.ScrollHeight = topPos + 20
    On Error GoTo 0
End Sub

' --- 6. œ«·… ≈‰‘«¡ «·›Ê«’· «·»’—Ì… ··√ﬁ”«„ ---
Private Sub CreateSectionDivider(Caption As String, ByRef topPos As Long)
    Dim lblDiv As MSForms.Label
    Set lblDiv = DynamicFrame.Controls.Add("Forms.Label.1", "lblDiv" & topPos, True)
    With lblDiv
        .Caption = Caption
        .Left = 10
        .Top = topPos
        .Width = 560
        .Height = 22
        .BackColor = RGB(230, 235, 245)
        .ForeColor = RGB(30, 50, 90)
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
        .Alignment = 1
    End With
    topPos = topPos + 30
End Sub

' --- 7. œ«·… «· Ê·Ìœ «· ·ﬁ«∆Ì ··ÊÕœ«  «·—ﬁ«»Ì… «·’›¯Ì… ---
Private Sub CreateRowUnit(Title As String, CurVal As Double, PastVal As Double, DiffVal As Double, ByRef topPos As Long, Idx As Long)
    Dim lblTitle As MSForms.Label
    Dim txtCur As MSForms.TextBox, txtPast As MSForms.TextBox, txtDiff As MSForms.TextBox
    Dim lblIndicator As MSForms.Label
    
    Set lblTitle = DynamicFrame.Controls.Add("Forms.Label.1", "lblT" & Idx, True)
    With lblTitle
        .Caption = Title
        .Left = 10
        .Top = topPos + 3
        .Width = 115
        .Font.Name = "Segoe UI"
        .Font.Size = 9.5
    End With
    
    Set txtCur = DynamicFrame.Controls.Add("Forms.TextBox.1", "txtCur" & Idx, True)
    ConfigureTextBox txtCur, Format(CurVal, "#,##0.00"), 130, topPos
    
    Set txtPast = DynamicFrame.Controls.Add("Forms.TextBox.1", "txtPast" & Idx, True)
    ConfigureTextBox txtPast, Format(PastVal, "#,##0.00"), 240, topPos
    
    Set txtDiff = DynamicFrame.Controls.Add("Forms.TextBox.1", "txtDiff" & Idx, True)
    ConfigureTextBox txtDiff, Format(DiffVal, "#,##0.00"), 350, topPos
    txtDiff.Font.Bold = True
    
    Set lblIndicator = DynamicFrame.Controls.Add("Forms.Label.1", "lblInd" & Idx, True)
    With lblIndicator
        .Left = 460
        .Top = topPos
        .Width = 120
        .Height = 20
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 9
        .Alignment = 1
        
        If DiffVal > 0.009 Then
            .Caption = "“Ì«œ… ??"
            .ForeColor = RGB(0, 110, 0)
            .BackColor = RGB(220, 245, 220)
        ElseIf DiffVal < -0.009 Then
            .Caption = "‰ﬁ’«‰ ??"
            .ForeColor = RGB(170, 0, 0)
            .BackColor = RGB(255, 225, 225)
        Else
            .Caption = "„ ÿ«»ﬁ ??"
            .ForeColor = RGB(90, 90, 90)
            .BackColor = RGB(242, 242, 242)
        End If
    End With
    
    topPos = topPos + 26
End Sub

' --- 8. ≈⁄œ«œ«  «· ‰”Ìﬁ «·„ÊÕœ ·’‰«œÌﬁ «·‰’Ê’ «·—ﬁ„Ì… ---
Private Sub ConfigureTextBox(txt As MSForms.TextBox, ValText As String, leftPos As Long, topPos As Long)
    With txt
        .Text = ValText
        .Left = leftPos
        .Top = topPos
        .Width = 95
        .Height = 20
        .Locked = True
        .Alignment = 2
        .Font.Name = "Consolas"
        .Font.Size = 9.5
        .BackColor = RGB(253, 253, 255)
        .BorderStyle = 1
    End With
End Sub

' --- 9. „Õ—ﬂ ‘—Ìÿ «·„Ê«“‰… «·≈Ã„«·Ì… «·‰Â«∆Ì (Œ·«’… AW ·’«›Ì «·—Ê« ») ---
Private Sub BuildFooterSummary()
    Dim TopBar As Long, netDiff As Double
    Dim lblNetTitle As MSForms.Label
    Dim txtNetCur As MSForms.TextBox, txtNetPast As MSForms.TextBox, txtNetDiff As MSForms.TextBox
    
    TopBar = 485
    netDiff = AuditData.CurrentNetSalary - AuditData.PastNetSalary
    
    Set lblNetTitle = Me.Controls.Add("Forms.Label.1", "lblNetTitle", True)
    With lblNetTitle
        .Caption = "’«›Ì «·—Ê« » «·≈Ã„«·Ì (AW):"
        .Left = 15
        .Top = TopBar + 3
        .Width = 145
        .Font.Name = "Segoe UI"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    Set txtNetCur = Me.Controls.Add("Forms.TextBox.1", "txtNetCur", True)
    ConfigureTextBox txtNetCur, Format(AuditData.CurrentNetSalary, "#,##0.00"), 145, TopBar
    txtNetCur.Font.Bold = True
    
    Set txtNetPast = Me.Controls.Add("Forms.TextBox.1", "txtNetPast", True)
    ConfigureTextBox txtNetPast, Format(AuditData.PastNetSalary, "#,##0.00"), 255, TopBar
    txtNetPast.Font.Bold = True
    
    Set txtNetDiff = Me.Controls.Add("Forms.TextBox.1", "txtNetDiff", True)
    ConfigureTextBox txtNetDiff, Format(netDiff, "#,##0.00"), 365, TopBar
    txtNetDiff.Font.Bold = True
    
    If netDiff > 0.009 Then
        txtNetDiff.BackColor = RGB(195, 245, 195)
        txtNetDiff.ForeColor = RGB(0, 90, 0)
    ElseIf netDiff < -0.009 Then
        txtNetDiff.BackColor = RGB(255, 195, 195)
        txtNetDiff.ForeColor = RGB(120, 0, 0)
    End If
End Sub
