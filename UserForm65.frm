VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm65 
   Caption         =   "UserForm65"
   ClientHeight    =   8460.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13716
   OleObjectBlob   =   "UserForm65.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm65"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Unload Me
UserForm1.Show
End Sub

Private Sub CommandButton2_Click()
    Dim wbPrint As Workbook
    Dim wsPrint As Worksheet
    Dim i As Long, j As Long
    
    ' 1. «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  ›Ì «··”  »Êﬂ”
    If ListBox1.ListCount <= 1 Then ' 1 ·√‰ «·’› «·√Ê· ÂÊ «·⁄‰Ê«‰
        MsgBox "·«  ÊÃœ »Ì«‰«  ··ÿ»«⁄…", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    
    ' 2. ≈‰‘«¡ „·› „ƒﬁ  ··ÿ»«⁄…
    Set wbPrint = Workbooks.Add
    Set wsPrint = wbPrint.Sheets(1)
    wsPrint.DisplayRightToLeft = True
    
    ' 3. ‰ﬁ· «·»Ì«‰«  „‰ ListBox1 ≈·Ï «·‘Ì  «·„ƒﬁ 
    For i = 0 To ListBox1.ListCount - 1
        For j = 0 To 12
            wsPrint.Cells(i + 1, j + 1).value = ListBox1.List(i, j)
            ' ≈÷«›… ÕœÊœ ··Œ·«Ì«
            wsPrint.Cells(i + 1, j + 1).Borders.LineStyle = xlContinuous
        Next j
        '  „ÌÌ“ ’› «·⁄‰Ê«‰ »«··Ê‰ Ê«·Œÿ «·⁄—Ì÷
        If i = 0 Then
            wsPrint.Rows(1).Font.Bold = True
            wsPrint.Rows(1).Interior.Color = RGB(220, 220, 220)
        End If
    Next i
    
    wsPrint.Columns.AutoFit
    
    ' 4. ÷»ÿ ≈⁄œ«œ«  «·’›Õ… (Landscape Ê«Õ Ê«¡ «·√⁄„œ…)
    With wsPrint.PageSetup
        .Orientation = xlLandscape ' ÷»ÿ «·Ê—ﬁ… »«·⁄—÷
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1 ' «Õ Ê«¡ «·‹ 13 ⁄„Êœ ›Ì ⁄—÷ «·’›Õ…
        .FitToPagesTall = False
        .CenterHorizontally = True '  Ê”Ìÿ «·ÃœÊ· √›ﬁÌ«
        ' ≈÷«›…  «—ÌŒ «·ÿ»«⁄… Ê—ﬁ„ «·’›Õ… ›Ì «· –ÌÌ· («Œ Ì«—Ì)
        .LeftFooter = "&D"
        .RightFooter = "’›Õ… &P „‰ &N"
    End With

    ' 5. ≈ŸÂ«— «·≈ﬂ”Ì· Ê«·„⁄«Ì‰…
    Me.Hide ' ≈Œ›«¡ «·›Ê—„ ··”„«Õ »«·„⁄«Ì‰…
    Application.Visible = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    
    ' ⁄—÷ «·„⁄«Ì‰…
    wsPrint.PrintPreview
    
    ' 6. »⁄œ ≈€·«ﬁ «·„⁄«Ì‰…: «·⁄Êœ… ·Ê÷⁄ «·ÌÊ“— ›Ê—„ ›ﬁÿ
    wbPrint.Close SaveChanges:=False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.Visible = False
    
    Me.Show
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton3_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(37)
    Dim i As Long, j As Long, lastRow As Long
    Dim IDToFind As String: IDToFind = Me.TextBox1.Text
    
    If IDToFind = "" Then MsgBox "Ì—ÃÏ «Œ Ì«— ”Ã·": Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).value) = IDToFind Then
            For j = 1 To 13
                Dim val As String: val = Me.Controls("TextBox" & j).Text
                
                ' ≈—”«· «·ﬁÌ„… ··‘Ì  ﬂ‹  «—ÌŒ ÕﬁÌﬁÌ ≈–« ﬂ«‰   ’·Õ ·–·ﬂ
                If IsDate(val) Then
                    ws.Cells(i, j).value = CDate(val)
                    ws.Cells(i, j).NumberFormat = "yyyy/mm/dd" ' ÷„«‰  ‰”Ìﬁ «·Œ·Ì…
                Else
                    ws.Cells(i, j).value = val
                End If
            Next j
            
            DoEvents ' ≈⁄ÿ«¡ Êﬁ  ··≈ﬂ”Ì· · ÕœÌÀ «·‘Ì  œ«Œ·Ì«
            MsgBox " „  ÕœÌÀ «·”Ã· »‰Ã«Õ", vbInformation
            
            ' «” œ⁄«¡ «· ÕœÌÀ ›Ê—« ·ÌŸÂ— «· «—ÌŒ „‰”ﬁ« „⁄ «·⁄‰«ÊÌ‰
            Call RefreshMyListBox(Me.TextBox14.Text)
            Exit Sub
        End If
    Next i
End Sub

Private Sub CommandButton4_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(37)
    Dim lastRow As Long, i As Long
    Dim IDToFind As String: IDToFind = Me.TextBox1.Text
    Dim confirm As VbMsgBoxResult
    
    ' 1. «· Õﬁﬁ „‰ «Œ Ì«— ”Ã·
    If IDToFind = "" Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ”Ã· „‰ «·ﬁ«∆„… √Ê·« ·Õ–›Â", vbExclamation, " ‰»ÌÂ"
        Exit Sub
    End If
    
    ' 2. —”«·…  √ﬂÌœ «·Õ–›
    confirm = MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› Â–« «·”Ã· ‰Â«∆Ì«ø", vbCritical + vbYesNo, " √ﬂÌœ «·Õ–›")
    
    If confirm = vbYes Then
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        
        ' 3. «·»ÕÀ ⁄‰ «·’› ÊÕ–›Â
        For i = 2 To lastRow
            If CStr(ws.Cells(i, 1).value) = IDToFind Then
                ws.Rows(i).Delete
                
                MsgBox " „ Õ–› «·”Ã· »‰Ã«Õ", vbInformation, " „"
                
                ' 4.  ›—Ì€ «· Ìﬂ”  »Êﬂ”«  »⁄œ «·Õ–›
                Dim ctl As Control
                For Each ctl In Me.Controls
                    If TypeName(ctl) = "TextBox" Then ctl.value = ""
                Next ctl
                
                ' 5.  ÕœÌÀ «··”  »Êﬂ” ›Ê—« ·Ì⁄ﬂ” «·Õ–›
                Call RefreshMyListBox(Me.TextBox14.Text)
                Exit Sub
            End If
        Next i
    End If
End Sub

Private Sub Label1_Click()

End Sub



Private Sub ListBox1_Click()
    ' «· Õﬁﬁ „‰ ÊÃÊœ ”Ã· „Œ «— (ListIndex = -1 Ì⁄‰Ì ·« ÌÊÃœ «Œ Ì«—)
    ' Ê«· Õﬁﬁ „‰ ⁄œ„ «·‰ﬁ— ⁄·Ï ’› «·⁄‰Ê«‰ (ListIndex = 0)
    If ListBox1.ListIndex <= 0 Then Exit Sub
    
    Dim i As Integer
    Dim cellVal As Variant
    
    On Error Resume Next ' · Ã‰» √Ì Œÿ√ „›«Ã∆ √À‰«¡ «· ⁄»∆…
    For i = 1 To 13
        cellVal = ListBox1.List(ListBox1.ListIndex, i - 1)
        Me.Controls("TextBox" & i).Text = cellVal
    Next i
    On Error GoTo 0
End Sub

Private Sub UserForm_Initialize()

 Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(37)
    Dim i As Integer

    ' 1.  ÕœÌÀ √”„«¡ «·‹ Labels „‰ «·’› «·√Ê· ›Ì «·‘Ì 
    On Error Resume Next ' · Ã‰» «·Œÿ√ ›Ì Õ«· ﬂ«‰ ⁄œœ Labels √ﬁ· „‰ 13
    For i = 1 To 13
        Me.Controls("Label" & i).Caption = ws.Cells(1, i).value
    Next i
    On Error GoTo 0

    ' 2. ≈⁄œ«œ«  «··”  »Êﬂ” «·√”«”Ì…
    ListBox1.ColumnCount = 13
      ListBox1.ColumnWidths = "80;140;80;100;120;120;120;120;120;120;120;120;120"
    
    ' 3. «” œ⁄«¡ «·„Õ—ﬂ «·„ÊÕœ · ⁄»∆… «·»Ì«‰«  Ê—ƒÊ” «·√⁄„œ…
    Call RefreshMyListBox
   
    

End Sub
    
   

Private Sub TextBox14_Change()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(37)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim filterVal As String: filterVal = Trim(Me.TextBox14.Text)
    Dim i As Long, j As Long, k As Long, r As Long
    Dim outArr() As Variant
    
    ' ≈–« ﬂ«‰ «·‘Ì  ›«—€«  „«„«
    If lastRow < 1 Then ListBox1.Clear: Exit Sub
    
    '  ÃÂÌ“ „’›Ê›… ·‰ «∆Ã «·›· —… (»Õœ √ﬁ’Ï ⁄œœ ’›Ê› «·‘Ì )
    ReDim outArr(1 To lastRow, 1 To 13)
    
    ' 1. «·ŒÿÊ… «·√Â„: ≈÷«›… —ƒÊ” «·√⁄„œ… œ«∆„« ›Ì «·’› «·√Ê· ··„’›Ê›…
    r = 1
    For k = 1 To 13
        outArr(r, k) = ws.Cells(1, k).Text
    Next k
    
    ' 2. «·»ÕÀ Ê«·›· —…
    If lastRow >= 2 Then
        For i = 2 To lastRow
            Dim isMatch As Boolean: isMatch = False
            
            ' ≈–« ﬂ«‰ „—»⁄ «·»ÕÀ ›«—€«  ŸÂ— ﬂ· «·»Ì«‰« 
            If filterVal = "" Then
                isMatch = True
            Else
                ' «·»ÕÀ ›Ì «·‹ 13 ⁄„Êœ«
                For j = 1 To 13
                    If InStr(1, ws.Cells(i, j).Text, filterVal, vbTextCompare) > 0 Then
                        isMatch = True
                        Exit For
                    End If
                Next j
            End If
            
            ' ≈–« ÊÃœ‰«  ÿ«»ﬁ ‰÷Ì› «·”Ã· ··„’›Ê›…
            If isMatch Then
                r = r + 1
                For k = 1 To 13
                    outArr(r, k) = ws.Cells(i, k).Text ' «” Œœ«„ .Text ·÷„«‰  ‰”Ìﬁ «· «—ÌŒ
                Next k
            End If
        Next i
    End If
    
    ' 3.  ÕœÌÀ «··”  »Êﬂ”
    ListBox1.Clear
    If r > 0 Then
        '  ﬁ·Ì’ «·„’›Ê›… · ‰«”» ⁄œœ «·‰ «∆Ã ›ﬁÿ (»„« ›ÌÂ« «·⁄‰Ê«‰)
        Dim finalArr() As Variant
        ReDim finalArr(1 To r, 1 To 13)
        For i = 1 To r
            For j = 1 To 13: finalArr(i, j) = outArr(i, j): Next j
        Next i
        ListBox1.List = finalArr
    End If
End Sub




Sub RefreshMyListBox(Optional filterVal As String = "")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(37)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    Dim outArr() As Variant
    Dim i As Long, j As Long, k As Long, r As Long
    
    If lastRow < 1 Then Exit Sub
    
    '  ÃÂÌ“ „’›Ê›… «·„Œ—Ã« 
    ReDim outArr(1 To lastRow, 1 To 13)
    
    ' 1. ≈÷«›… «·’› «·√Ê· («·⁄‰«ÊÌ‰)
    r = 1
    For k = 1 To 13: outArr(r, k) = ws.Cells(1, k).Text: Next k
    
    ' 2. ≈÷«›… «·»Ì«‰«  „⁄ «·›· —…
    For i = 2 To lastRow
        Dim matchFound As Boolean: matchFound = False
        
        If filterVal = "" Then
            matchFound = True
        Else
            ' ›Õ’ ﬂ· «·√⁄„œ… »ÕÀ« ⁄‰ «·‰’
            For j = 1 To 13
                If InStr(1, ws.Cells(i, j).Text, filterVal, vbTextCompare) > 0 Then
                    matchFound = True: Exit For
                End If
            Next j
        End If
        
        If matchFound Then
            r = r + 1
            For k = 1 To 13
                ' «” Œœ«„ .Text Ì÷„‰ Ã·» «· «—ÌŒ ﬂ„« ÂÊ Ÿ«Â— (2024/05/15) Ê·Ì” (45123)
                outArr(r, k) = ws.Cells(i, k).Text
            Next k
        End If
    Next i
    
    ' 3.  ÕœÌÀ «··”  »Êﬂ”
    ListBox1.Clear
    If r > 0 Then
        '  ﬁ·Ì’ «·„’›Ê›… · ‰«”» ⁄œœ «·‰ «∆Ã ›ﬁÿ · Ã‰» «·√”ÿ— «·›«—€…
        Dim finalArr() As Variant
        ReDim finalArr(1 To r, 1 To 13)
        For i = 1 To r
            For j = 1 To 13: finalArr(i, j) = outArr(i, j): Next j
        Next i
        ListBox1.List = finalArr
    End If
End Sub
