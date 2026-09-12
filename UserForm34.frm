VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm34 
   Caption         =   "UserForm34"
   ClientHeight    =   9048.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   15588
   OleObjectBlob   =   "UserForm34.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Dim FullData() As Variant ' „’›Ê›… · Œ“Ì‰ ﬂ· «·»Ì«‰«  «·√’·Ì…

Private Sub CommandButton1_Click()
On Error Resume Next
Sheets("clearsheet").Activate
 Dim ws As Worksheet
    Dim i As Long
    Dim itm As ListItem
    
    ' ≈Œ›«¡ «·›Ê—„ √À‰«¡ «·⁄„·Ì…
    Me.Hide
    
    ' ≈ŸÂ«— Excel „ƒﬁ «
    Application.Visible = True
    Application.ScreenUpdating = True
    
    '  ⁄ÌÌ‰ «·‘Ì  «·„ÊÃÊœ „”»ﬁ«
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets("sheettemp")
    On Error GoTo 0
    
    ' „”Õ √Ì »Ì«‰«  ﬁœÌ„… („⁄ «·Õ›«Ÿ ⁄·Ï —ƒÊ” «·√⁄„œ…)
    ws.Cells.Clear
    
    ' ≈÷«›… —ƒÊ” «·√⁄„œ…
    ws.Cells(1, 1).value = "«·—ﬁ„ «·ÊŸÌ›Ì"
    ws.Cells(1, 2).value = "«·«”„"
    ws.Cells(1, 3).value = "”»» «·≈Õ«·… ≈·Ï ‰Ÿ«„ –ﬂ—‰Ì"
    ws.Cells(1, 4).value = "«·ﬁÌ„…"
    ws.Cells(1, 5).value = " «—ÌŒ «·≈Õ«·…"
    
    ws.rowS(1).Font.Bold = True
    ws.rowS(1).Font.Size = 12
    
    '  ⁄»∆… «·»Ì«‰«  „‰ ListView „»«‘—…
    For i = 1 To Me.ListView1.ListItems.count
        Set itm = Me.ListView1.ListItems(i)
        ws.Cells(i + 1, 1).value = itm.Text                   ' «·—ﬁ„ «·ÊŸÌ›Ì
        ws.Cells(i + 1, 2).value = itm.ListSubItems(1).Text   ' «·«”„
        ws.Cells(i + 1, 3).value = itm.ListSubItems(2).Text   ' ”»» «·≈Õ«·…
        ws.Cells(i + 1, 4).value = itm.ListSubItems(3).Text   ' «·ﬁÌ„…
        ws.Cells(i + 1, 5).value = itm.ListSubItems(4).Text   '  «—ÌŒ «·≈Õ«·…
    Next i
    
    ' ÷»ÿ ⁄—÷ «·√⁄„œ…  ·ﬁ«∆Ì
    ws.Columns("A:E").AutoFit
   ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    ' ⁄—÷ Print Preview
    ws.PrintPreview
    ThisWorkbook.Application.Visible = False
    ' ≈⁄«œ… ≈ŸÂ«— «·›Ê—„
    Me.Show
    
    Exit Sub

ErrHandler:
    MsgBox "«·‘Ì  'sheettemp' €Ì— „ÊÃÊœ ›Ì «·„’‰›.", vbExclamation
    Me.Show
End Sub

Private Sub CommandButton2_Click()
  Unload Me
    
    ' ????? UserForm1 ??? ????? ??? ??????
    UserForm1.Show
End Sub

Private Sub CommandButton3_Click()
Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dt As Variant
    Dim Idx As Long
    
    ' 1. ≈⁄œ«œ«  √Ê·Ì…
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.rowS.count, "A").End(xlUp).row
    
    ' ÿ·»  √ﬂÌœ ﬁ»· «·„”Õ («Œ Ì«—Ì ·ﬂ‰Â ¬„‰)
    If MsgBox("Â· √‰  „ √ﬂœ „‰  ’›Ì— €Ì«»«  «·‘Â— «·ÃœÌœø", vbQuestion + vbYesNo, " √ﬂÌœ") = vbNo Then Exit Sub

    Application.ScreenUpdating = False

    ' 2. ⁄„·Ì… «· ’›Ì— ›Ì «·‘Ì  («·⁄„Êœ K Ê AU)
    If lastRow >= 9 Then
        For i = 9 To lastRow
            If val(ws.Cells(i, "K").value) > 0 Then ws.Cells(i, "K").value = 0
            If val(ws.Cells(i, "l").value) < 30 Then ws.Cells(i, "l").value = 30
            If val(ws.Cells(i, "AU").value) > 0 Then ws.Cells(i, "AU").value = 0
        Next i
    End If

    ' 3.  ÕœÌÀ „ƒ‘— «·€Ì«»«  Label8 (Õ«·… «·«” ﬁ—«—)
    With Label8
        .BackColor = RGB(0, 150, 0) ' √Œ÷— Â«œ∆ ··«” ﬁ—«—
        .ForeColor = vbWhite
        .Caption = "·« ÌÊÃœ €Ì«»« "
    End With

    ' 4. ‰œ«¡ «·œ«·… «·Õ”«»Ì… «·Œ«’… »ﬂ
    Call count_all

    ' 5. ≈⁄«œ… »‰«¡ «·‹ ListView («” Œœ«„ «·ﬂÊœ «·Œ«’ »ﬂ ·· ÕœÌÀ)
    '  ›—Ì€ «·ﬁ«∆„… «·Õ«·Ì…
    Me.ListView1.ListItems.Clear
    Idx = 0
    Erase FullData '  ›—Ì€ «·„’›Ê›… «·ﬁœÌ„…
    
    ' ≈⁄«œ… «·›Õ’ (»„« √‰ K √’»Õ ’›—«° ·‰ ÌŸÂ— «·€Ì«» ›Ì «·ﬁ«∆„…)
    For i = 9 To lastRow
        dt = ws.Cells(i, "BN").value
        If IsNumeric(dt) And dt <> "" Then dt = CDate(dt)
        
        ' ›Õ’ «·€Ì«» (”ÌﬂÊ‰ ’›—« «·¬‰)
        If val(ws.Cells(i, "K").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "·œÌÂ €Ì«»"
            FullData(4, Idx) = ws.Cells(i, "K").value
            FullData(5, Idx) = dt
        End If
        
        ' ›Õ’ »ﬁÌ… «·Õ«·«  (‰›ﬁ…° ”·›…° ﬂ«‘... ≈·Œ)
        ' „·«ÕŸ…: Â–Â «·Õ«·«  ” »ﬁÏ  ŸÂ— ·√‰‰« ·„ ‰’›— √⁄„œ Â«
        If val(ws.Cells(i, "AJ").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "·œÌÂ ‰›ﬁ…"
            FullData(4, Idx) = ws.Cells(i, "AJ").value
            FullData(5, Idx) = dt
        End If
        
        ' ... [√ﬂ„· »ﬁÌ… ‘—Êÿ «·›Õ’ «·Œ«’… »ﬂ Â‰« ﬂ„« ÂÌ] ...
    Next i
    
    ' „‰«œ«… «·”Ê» "«·„ÊÃÊœ ⁄‰œﬂ" ·„·¡ «·ﬁ«∆„… »«·‰ «∆Ã «·ÃœÌœ…
    Call PopulateListView(FullData, Idx)
    
    '  ÕœÌÀ «·⁄œ«œ ›Ì «· ﬂ”  »Êﬂ”
    Me.TextBox3.value = Me.ListView1.ListItems.count

    Application.ScreenUpdating = True
    
    MsgBox " „  ’›Ì— «·€Ì«»«  Ê ÕœÌÀ —«œ«— «·„—«ﬁ»… »‰Ã«Õ.", vbInformation, "‰Ÿ«„ ⁄„«œ «·‰⁄Ì„Ì"
End Sub

Private Sub CommandButton4_Click()
Unload Me
UserForm78.Show

End Sub

Private Sub UserForm_Activate()
Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long: lastRow = ws.Cells(ws.rowS.count, "A").End(xlUp).row
    Dim i As Long
    Dim hasAbsence As Boolean: hasAbsence = False

    ' ›Õ’ ”—Ì⁄ ··⁄„Êœ K ﬁ»·  ⁄»∆… «·ﬁ«∆„…
    For i = 9 To lastRow
        If val(ws.Cells(i, "K").value) > 0 Then
            hasAbsence = True
            Exit For ' ‰Œ—Ã »„Ã—œ ≈ÌÃ«œ √Ê· €Ì«» · Ê›Ì— «·Êﬁ 
        End If
    Next i

    ' ÷»ÿ Õ«·… «··Ì»·
    If hasAbsence Then
        With Label8
            .BackColor = vbYellow ' ›”›Ê—Ì
            .Caption = " ‰»ÌÂ: ÌÊÃœ €Ì«»«  €Ì— „’›¯—…"
            ' Â‰« Ì„ﬂ‰ﬂ „‰«œ«… ﬂÊœ «·Ê„Ì÷ ≈–« ﬂ‰   ” Œœ„Â
        End With
    Else
        With Label8
            .BackColor = RGB(0, 150, 0) ' √Œ÷— «” ﬁ—«—
            .Caption = "·« ÌÊÃœ €Ì«»« "
        End With
    End If
    
    ' À„  ﬂ„· »ﬁÌ… ﬂÊœ  ⁄»∆… «·‹ ListView «·Œ«’ »ﬂ...
End Sub

' ===============================
' ÕœÀ «· Õ„Ì·
' ===============================
   
Private Sub UserForm_Initialize()
On Error Resume Next
 
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dt As Variant
    Dim Idx As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.rowS.count, "A").End(xlUp).row
    
    '  ÂÌ∆… ListView
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideSelection = False
        .ColumnHeaders.Clear
        .ListItems.Clear
        .ColumnHeaders.Add , , "«·—ﬁ„ «·ÊŸÌ›Ì", 80
        .ColumnHeaders.Add , , "«·«”„", 150
        .ColumnHeaders.Add , , "”»» «·≈Õ«·… ≈·Ï ‰Ÿ«„ –ﬂ—‰Ì", 260
        .ColumnHeaders.Add , , "«·ﬁÌ„…", 80
        .ColumnHeaders.Add , , " «—ÌŒ «·≈Õ«·…", 100
        .Font.Name = "Times New Roman"
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' =========================
    '  ⁄»∆… fullData
    ' =========================
    Idx = 0
    For i = 9 To lastRow
        dt = ws.Cells(i, "BN").value
        If IsNumeric(dt) And dt <> "" Then dt = CDate(dt)
        
        ' «·€Ì«» K
        If val(ws.Cells(i, "K").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "·œÌÂ €Ì«»"
            FullData(4, Idx) = ws.Cells(i, "K").value
            FullData(5, Idx) = dt
        End If
        
        ' «·‰›ﬁ… AJ
        If val(ws.Cells(i, "AJ").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "·œÌÂ ‰›ﬁ…"
            FullData(4, Idx) = ws.Cells(i, "AJ").value
            FullData(5, Idx) = dt
        End If
        
        ' «·”·›… AX
        If val(ws.Cells(i, "AX").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "·œÌÂ ”·›…"
            FullData(4, Idx) = ws.Cells(i, "AX").value
            FullData(5, Idx) = dt
        End If
        
        ' ﬂ«‘ AY
        If Trim(ws.Cells(i, "AY").value) = "ﬂ«‘" Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "ÿ—Ìﬁ… «·œ›⁄ ﬂ«‘"
            FullData(4, Idx) = ws.Cells(i, "aw").value
            FullData(5, Idx) = dt
        End If
        
        ' BF Õ«·«  Œ«’…
        If ws.Cells(i, "BF").value <> "" Then
            If ws.Cells(i, "BF").value Like "*—« » Ã“∆Ì*" _
            Or ws.Cells(i, "BF").value Like "*—›⁄ Ìœ*" _
            Or ws.Cells(i, "BF").value Like "*«Ã«“… »œÊ‰ —« »*" _
            Or ws.Cells(i, "BF").value Like "*«„Ê„… À«‰Ì 6 «‘Â—*" Then
                Idx = Idx + 1
                ReDim Preserve FullData(1 To 5, 1 To Idx)
                FullData(1, Idx) = ws.Cells(i, "A").value
                FullData(2, Idx) = ws.Cells(i, "E").value
                FullData(3, Idx) = ws.Cells(i, "BF").value
                FullData(4, Idx) = ws.Cells(i, "aw").value
                FullData(5, Idx) = dt
            End If
        End If
        
        ' BH
       '==================================
         If val(ws.Cells(i, "BH").value) > 1 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "«·«÷«›… «·„«·Ì… «·„ƒﬁ Â"
            FullData(4, Idx) = CInt(ws.Cells(i, "BH"))
            FullData(5, Idx) = dt
        End If
         '=============================«” ﬁÿ«⁄„ ›—ﬁ
        If val(ws.Cells(i, "ar").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "·œÌÂ «” ﬁÿ«⁄ „ ›—ﬁ"
            FullData(4, Idx) = ws.Cells(i, "ar").value
            FullData(5, Idx) = dt
        End If
        
        
        
        
        '===============================
        
        
        '=============================«” ﬁÿ«⁄„ ›—ﬁ
       
        
        
        '===============================
        
        
        ' BJ
        If val(ws.Cells(i, "BJ").value) > 1 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "«·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ "
            FullData(4, Idx) = CInt(ws.Cells(i, "BJ"))
            FullData(5, Idx) = dt
        End If
       
       
       
       
       
       
       '============================================
        
        ' CE + «·œ—Ã…
        If val(ws.Cells(i, "CE").value) > 0 And val(ws.Cells(i, "H").value) < 8 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = "«·œ—Ã… √ﬁ· „‰ 8 „⁄ ÊÃÊœ «” Õﬁ«ﬁ CE"
            FullData(4, Idx) = ws.Cells(i, "CE").value
            FullData(5, Idx) = dt
        End If
        
        ' AN > 0
        If val(ws.Cells(i, "AN").value) > 0 Then
            Idx = Idx + 1
            ReDim Preserve FullData(1 To 5, 1 To Idx)
            FullData(1, Idx) = ws.Cells(i, "A").value
            FullData(2, Idx) = ws.Cells(i, "E").value
            FullData(3, Idx) = ws.Cells(i, "AO").value
            FullData(4, Idx) = ws.Cells(i, "AN").value
            FullData(5, Idx) = dt
        End If
    Next i
    
    ' ⁄—÷ ﬂ· «·»Ì«‰«  √Ê· „—…
    Call PopulateListView(FullData, Idx)
    
    '  ÕœÌÀ ⁄œœ «·’›Ê›
    Me.TextBox3.value = Me.ListView1.ListItems.count
End Sub

' ===============================
' ≈⁄«œ… »‰«¡ ListView
' ===============================
Private Sub PopulateListView(dataArray() As Variant, rowCount As Long)
    Dim i As Long, itm As ListItem
    Me.ListView1.ListItems.Clear
    
    For i = 1 To rowCount
        Set itm = Me.ListView1.ListItems.Add(, , dataArray(1, i))
        itm.ListSubItems.Add , , dataArray(2, i)
        itm.ListSubItems.Add , , dataArray(3, i)
        itm.ListSubItems.Add , , dataArray(4, i)
        If IsDate(dataArray(5, i)) Then
            itm.ListSubItems.Add , , Format(dataArray(5, i), "dd/mm/yyyy")
        Else
            itm.ListSubItems.Add , , dataArray(5, i)
        End If
    Next i
    
    '  ÕœÌÀ ⁄œœ «·’›Ê›
    Me.TextBox3.value = Me.ListView1.ListItems.count
End Sub

' ===============================
' ›· —… Õ”» TextBox2 («·‘Â—)
' ===============================
Private Sub TextBox2_Change()
    Dim i As Long, Idx As Long
    Dim filtered() As Variant
    Dim monthValue As Long
    
    If Trim(Me.TextBox2.value) = "" Then
        Call PopulateListView(FullData, UBound(FullData, 2))
        Exit Sub
    End If
    
    If Not IsNumeric(Me.TextBox2.value) Then Exit Sub
    monthValue = CLng(Me.TextBox2.value)
    
    ReDim filtered(1 To 5, 1 To UBound(FullData, 2))
    Idx = 0
    For i = 1 To UBound(FullData, 2)
        If IsDate(FullData(5, i)) Then
            If Month(FullData(5, i)) = monthValue Then
                Idx = Idx + 1
                filtered(1, Idx) = FullData(1, i)
                filtered(2, Idx) = FullData(2, i)
                filtered(3, Idx) = FullData(3, i)
                filtered(4, Idx) = FullData(4, i)
                filtered(5, Idx) = FullData(5, i)
            End If
        End If
    Next i
    
    If Idx > 0 Then
        Call PopulateListView(filtered, Idx)
    Else
        Me.ListView1.ListItems.Clear
        Me.TextBox3.value = 0
    End If
End Sub

' ===============================
' ›· —… Õ”» TextBox4 («·«”„)
' ===============================
Private Sub TextBox4_Change()
    Dim i As Long, Idx As Long
    Dim filtered() As Variant
    Dim searchName As String
    
    searchName = Trim(Me.TextBox4.value)
    
    If searchName = "" Then
        Call PopulateListView(FullData, UBound(FullData, 2))
        Exit Sub
    End If
    
    ReDim filtered(1 To 5, 1 To UBound(FullData, 2))
    Idx = 0
    For i = 1 To UBound(FullData, 2)
        If InStr(1, FullData(2, i), searchName, vbTextCompare) > 0 Then
            Idx = Idx + 1
            filtered(1, Idx) = FullData(1, i)
            filtered(2, Idx) = FullData(2, i)
            filtered(3, Idx) = FullData(3, i)
            filtered(4, Idx) = FullData(4, i)
            filtered(5, Idx) = FullData(5, i)
        End If
    Next i
    
    If Idx > 0 Then
        Call PopulateListView(filtered, Idx)
    Else
        Me.ListView1.ListItems.Clear
        Me.TextBox3.value = 0
    End If
End Sub

Private Sub TextBox5_Change()
    Dim i As Long, Idx As Long
    Dim filtered() As Variant
    Dim searchStatus As String
    
    searchStatus = Trim(Me.TextBox5.value)
    
    ' ≈–« ›«—€° ⁄—÷ ﬂ· «·»Ì«‰« 
    If searchStatus = "" Then
        Call PopulateListView(FullData, UBound(FullData, 2))
        Exit Sub
    End If
    
    ReDim filtered(1 To 5, 1 To UBound(FullData, 2))
    Idx = 0
    
    For i = 1 To UBound(FullData, 2)
        ' «·⁄„Êœ 3 ÌÕ ÊÌ ”»» «·≈Õ«·…
        If InStr(1, FullData(3, i), searchStatus, vbTextCompare) > 0 Then
            Idx = Idx + 1
            filtered(1, Idx) = FullData(1, i)
            filtered(2, Idx) = FullData(2, i)
            filtered(3, Idx) = FullData(3, i)
            filtered(4, Idx) = FullData(4, i)
            filtered(5, Idx) = FullData(5, i)
        End If
    Next i
    
    If Idx > 0 Then
        Call PopulateListView(filtered, Idx)
    Else
        Me.ListView1.ListItems.Clear
        Me.TextBox3.value = 0
    End If
End Sub

