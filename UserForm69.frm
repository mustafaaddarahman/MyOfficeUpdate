VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm69 
   Caption         =   "UserForm69"
   ClientHeight    =   15624
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9564.001
   OleObjectBlob   =   "UserForm69.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm69"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dictCount As Object, dictSum As Object
Public WithEvents btnPrint As MSForms.CommandButton
Attribute btnPrint.VB_VarHelpID = -1
Public WithEvents btnExit As MSForms.CommandButton
Attribute btnExit.VB_VarHelpID = -1

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton3_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub UserForm_Initialize()
    Dim ws As Worksheet: Dim lastRow As Long: Dim i As Integer: Dim topPos As Double
    Dim lblHeader As Object, lbl As Object, txtCount As Object, txtSum As Object, chk As Object
    Dim totalAllSalaries As Double: Dim key As Variant

    ' 1. ≈Œ›«¡ Ê«ÃÂ… «·≈ﬂ”Ì·
    Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", False)"
    Application.DisplayFormulaBar = False: ActiveWindow.DisplayHeadings = False
    
    ' 2.  ‰”Ìﬁ «·›Ê—„ (“Ì«œ… «·⁄—÷ ﬁ·Ì·« ·· ‰”Ìﬁ)
    Me.BackColor = &HFFFFC0: Me.Width = 650: Me.Caption = "‰Ÿ«„ «·—Ê« » Ê«·Ê’› «·ÊŸÌ›Ì"
    
    Set ws = ThisWorkbook.Sheets(1)
    Set dictCount = CreateObject("Scripting.Dictionary")
    Set dictSum = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.count, "CF").End(xlUp).row
    
    totalAllSalaries = 0
    For i = 9 To lastRow
        Dim job As String: job = Trim(ws.Cells(i, "CF").value)
        Dim sal As Double: sal = val(ws.Cells(i, "AW").value)
        If job <> "" Then
            dictCount(job) = dictCount(job) + 1
            dictSum(job) = dictSum(job) + sal
            totalAllSalaries = totalAllSalaries + sal
        End If
    Next i

    ' 3. ≈‰‘«¡ «·√“—«— (Œÿ 14 »Ê·œ Times New Roman)
    Set btnPrint = Me.Controls.Add("Forms.CommandButton.1", "btnTemp", True)
    With btnPrint: .Caption = "ÿ»«⁄… «·„Œ «—« ": .Left = 10: .Top = 10: .Width = 130: .Height = 35
        .BackColor = &H80FF80: .Font.Name = "Times New Roman": .Font.Size = 14: .Font.Bold = True: End With

    Set btnExit = Me.Controls.Add("Forms.CommandButton.1", "btnExit9", True)
    With btnExit: .Caption = "«·⁄Êœ… ··ﬁ«∆„…": .Left = 150: .Top = 10: .Width = 120: .Height = 35
        .BackColor = &H8080FF: .Font.Name = "Times New Roman": .Font.Size = 14: .Font.Bold = True: End With

    ' 4. «·≈Ã„«·Ì «·⁄«„
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblGrandTotal", True)
    With lbl: .Caption = "«·≈Ã„«·Ì: " & Format(totalAllSalaries, "#,##0.00"): .Left = 280: .Top = 10: .Width = 340: .Height = 35
        .Font.Name = "Times New Roman": .Font.Size = 16: .Font.Bold = True: .TextAlign = 2: .SpecialEffect = 1: End With

    ' 5. ≈‰‘«¡ ⁄‰«ÊÌ‰ «·√⁄„œ… ( Ê”Ì⁄ ⁄„Êœ ÿ»«⁄… ·ÌŸÂ— «·‰’ ﬂ«„·«)
    Dim hdrs: hdrs = Array("ÿ»«⁄…", "«·Ê’› «·ÊŸÌ›Ì", "«·⁄œœ", "„Ã„Ê⁄ «·—« »")
    Dim lPos: lPos = Array(10, 70, 260, 340): Dim wds: wds = Array(50, 180, 70, 260)
    
    For i = 0 To 3
        Set lblHeader = Me.Controls.Add("Forms.Label.1", "Hdr_" & i, True)
        With lblHeader
            .Caption = hdrs(i): .Left = lPos(i): .Top = 55: .Width = wds(i): .Height = 30
            .Font.Name = "Times New Roman": .Font.Size = 14: .Font.Bold = True: .TextAlign = 2
            .SpecialEffect = 1: .BackColor = &H80FFFF: End With
    Next i

    ' 6. »‰«¡ «·ﬁ«∆„… «·œÌ‰«„ÌﬂÌ…
    topPos = 90: i = 1
    For Each key In dictCount.Keys
        Set chk = Me.Controls.Add("Forms.CheckBox.1", "chkPrint_" & i, True): chk.Left = 25: chk.Top = topPos + 5: chk.value = True: chk.Width = 20
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblJob_" & i, True)
        With lbl: .Caption = key: .Left = 70: .Top = topPos: .Width = 180: .Height = 30: .Font.Name = "Times New Roman": .Font.Size = 16: .Font.Bold = True: .TextAlign = 2: .SpecialEffect = 1: End With
        Set txtCount = Me.Controls.Add("Forms.TextBox.1", "txtCount_" & i, True)
        With txtCount: .value = dictCount(key): .Left = 260: .Top = topPos: .Width = 70: .Height = 30: .Font.Name = "Times New Roman": .Font.Size = 12: .Font.Bold = True: .TextAlign = 2: .SpecialEffect = 1: End With
        Set txtSum = Me.Controls.Add("Forms.TextBox.1", "txtSum_" & i, True)
        With txtSum: .value = Format(dictSum(key), "#,##0.00"): .Left = 340: .Top = topPos: .Width = 260: .Height = 30: .Font.Name = "Times New Roman": .Font.Size = 12: .Font.Bold = True: .TextAlign = 2: .SpecialEffect = 1: End With
        topPos = topPos + 35: i = i + 1
    Next key
    Me.ScrollBars = 2: Me.ScrollHeight = topPos + 50
End Sub

' 7. «·√Õœ«À
Private Sub btnExit_Click(): Unload Me: UserForm9.Show: End Sub

Private Sub btnPrint_Click()
    Me.Hide
   ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Call StartPrinting
    Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", False)"
   ThisWorkbook.Application.Visible = False
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Me.Show
End Sub

Sub StartPrinting()
    Dim wsPrint As Worksheet: Dim i As Integer: Dim r As Long: r = 2
    On Error Resume Next
    Application.DisplayAlerts = False: ThisWorkbook.Sheets("Temporary_Print").Delete
    Set wsPrint = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)): wsPrint.Name = "Temporary_Print"
    Application.DisplayAlerts = True: On Error GoTo 0
    
    wsPrint.Columns("C:C").NumberFormat = "#,##0.00"
    wsPrint.Range("A1:C1").value = Array("«·Ê’› «·ÊŸÌ›Ì", "«·⁄œœ", "≈Ã„«·Ì «·—Ê« »")
    
    For i = 1 To dictCount.count
        If Me.Controls("chkPrint_" & i).value = True Then
            wsPrint.Cells(r, 1) = Me.Controls("lblJob_" & i).Caption
            wsPrint.Cells(r, 2) = Me.Controls("txtCount_" & i).value
            wsPrint.Cells(r, 3) = CDbl(Replace(Me.Controls("txtSum_" & i).value, ",", "")): r = r + 1
        End If
    Next i
    
    With wsPrint.Range("A1:C" & r - 1)
        .Font.Name = "Times New Roman": .Font.Size = 12: .Borders.LineStyle = 1: .HorizontalAlignment = -4108: .Columns.AutoFit: End With
    wsPrint.Columns("C:C").ColumnWidth = wsPrint.Columns("C:C").ColumnWidth + 7 ' “Ì«œ… «·√„«‰ ··Â«‘ «ﬁ
    
    wsPrint.PrintPreview
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ThisWorkbook.Save: Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", True)"
    Application.DisplayFormulaBar = True: ActiveWindow.DisplayHeadings = True
    On Error Resume Next: Application.DisplayAlerts = False: ThisWorkbook.Sheets("Temporary_Print").Delete: Application.DisplayAlerts = True
End Sub
