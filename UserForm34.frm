VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm34 
   Caption         =   "UserForm34"
   ClientHeight    =   9645.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17808
   OleObjectBlob   =   "UserForm34.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fullData() As Variant ' „’›Ê›… · Œ“Ì‰ ﬂ· «·»Ì«‰«  «·√’·Ì…

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
    
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Font.Size = 12
    
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

' ===============================
' ÕœÀ «· Õ„Ì·
' ===============================
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dt As Variant
    Dim idx As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
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
    idx = 0
    For i = 9 To lastRow
        dt = ws.Cells(i, "BN").value
        If IsNumeric(dt) And dt <> "" Then dt = CDate(dt)
        
        ' «·€Ì«» K
        If val(ws.Cells(i, "K").value) > 0 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "·œÌÂ €Ì«»"
            fullData(4, idx) = ws.Cells(i, "K").value
            fullData(5, idx) = dt
        End If
        
        ' «·‰›ﬁ… AJ
        If val(ws.Cells(i, "AJ").value) > 0 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "·œÌÂ ‰›ﬁ…"
            fullData(4, idx) = ws.Cells(i, "AJ").value
            fullData(5, idx) = dt
        End If
        
        ' «·”·›… AX
        If val(ws.Cells(i, "AX").value) > 0 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "·œÌÂ ”·›…"
            fullData(4, idx) = ws.Cells(i, "AX").value
            fullData(5, idx) = dt
        End If
        
        ' ﬂ«‘ AY
        If Trim(ws.Cells(i, "AY").value) = "ﬂ«‘" Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "ÿ—Ìﬁ… «·œ›⁄ ﬂ«‘"
            fullData(4, idx) = ws.Cells(i, "aw").value
            fullData(5, idx) = dt
        End If
        
        ' BF Õ«·«  Œ«’…
        If ws.Cells(i, "BF").value <> "" Then
            If ws.Cells(i, "BF").value Like "*—« » Ã“∆Ì*" _
            Or ws.Cells(i, "BF").value Like "*—›⁄ Ìœ*" _
            Or ws.Cells(i, "BF").value Like "*«Ã«“… »œÊ‰ —« »*" _
            Or ws.Cells(i, "BF").value Like "*«„Ê„… À«‰Ì 6 «‘Â—*" Then
                idx = idx + 1
                ReDim Preserve fullData(1 To 5, 1 To idx)
                fullData(1, idx) = ws.Cells(i, "A").value
                fullData(2, idx) = ws.Cells(i, "E").value
                fullData(3, idx) = ws.Cells(i, "BF").value
                fullData(4, idx) = ws.Cells(i, "aw").value
                fullData(5, idx) = dt
            End If
        End If
        
        ' BH
       '==================================
         If val(ws.Cells(i, "BH").value) > 1 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "«·«÷«›… «·„«·Ì… «·„ƒﬁ Â"
            fullData(4, idx) = CInt(ws.Cells(i, "BH"))
            fullData(5, idx) = dt
        End If
         '=============================«” ﬁÿ«⁄„ ›—ﬁ
        If val(ws.Cells(i, "ar").value) > 0 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "·œÌÂ «” ﬁÿ«⁄ „ ›—ﬁ"
            fullData(4, idx) = ws.Cells(i, "ar").value
            fullData(5, idx) = dt
        End If
        
        
        
        
        '===============================
        
        
        '=============================«” ﬁÿ«⁄„ ›—ﬁ
       
        
        
        '===============================
        
        
        ' BJ
        If val(ws.Cells(i, "BJ").value) > 1 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "«·«” ﬁÿ«⁄ «·„«·Ì «·„ƒﬁ "
            fullData(4, idx) = CInt(ws.Cells(i, "BJ"))
            fullData(5, idx) = dt
        End If
       
       
       
       
       
       
       '============================================
        
        ' CE + «·œ—Ã…
        If val(ws.Cells(i, "CE").value) > 0 And val(ws.Cells(i, "H").value) < 8 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = "«·œ—Ã… √ﬁ· „‰ 8 „⁄ ÊÃÊœ «” Õﬁ«ﬁ CE"
            fullData(4, idx) = ws.Cells(i, "CE").value
            fullData(5, idx) = dt
        End If
        
        ' AN > 0
        If val(ws.Cells(i, "AN").value) > 0 Then
            idx = idx + 1
            ReDim Preserve fullData(1 To 5, 1 To idx)
            fullData(1, idx) = ws.Cells(i, "A").value
            fullData(2, idx) = ws.Cells(i, "E").value
            fullData(3, idx) = ws.Cells(i, "AO").value
            fullData(4, idx) = ws.Cells(i, "AN").value
            fullData(5, idx) = dt
        End If
    Next i
    
    ' ⁄—÷ ﬂ· «·»Ì«‰«  √Ê· „—…
    Call PopulateListView(fullData, idx)
    
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
    Dim i As Long, idx As Long
    Dim filtered() As Variant
    Dim monthValue As Long
    
    If Trim(Me.TextBox2.value) = "" Then
        Call PopulateListView(fullData, UBound(fullData, 2))
        Exit Sub
    End If
    
    If Not IsNumeric(Me.TextBox2.value) Then Exit Sub
    monthValue = CLng(Me.TextBox2.value)
    
    ReDim filtered(1 To 5, 1 To UBound(fullData, 2))
    idx = 0
    For i = 1 To UBound(fullData, 2)
        If IsDate(fullData(5, i)) Then
            If Month(fullData(5, i)) = monthValue Then
                idx = idx + 1
                filtered(1, idx) = fullData(1, i)
                filtered(2, idx) = fullData(2, i)
                filtered(3, idx) = fullData(3, i)
                filtered(4, idx) = fullData(4, i)
                filtered(5, idx) = fullData(5, i)
            End If
        End If
    Next i
    
    If idx > 0 Then
        Call PopulateListView(filtered, idx)
    Else
        Me.ListView1.ListItems.Clear
        Me.TextBox3.value = 0
    End If
End Sub

' ===============================
' ›· —… Õ”» TextBox4 («·«”„)
' ===============================
Private Sub TextBox4_Change()
    Dim i As Long, idx As Long
    Dim filtered() As Variant
    Dim searchName As String
    
    searchName = Trim(Me.TextBox4.value)
    
    If searchName = "" Then
        Call PopulateListView(fullData, UBound(fullData, 2))
        Exit Sub
    End If
    
    ReDim filtered(1 To 5, 1 To UBound(fullData, 2))
    idx = 0
    For i = 1 To UBound(fullData, 2)
        If InStr(1, fullData(2, i), searchName, vbTextCompare) > 0 Then
            idx = idx + 1
            filtered(1, idx) = fullData(1, i)
            filtered(2, idx) = fullData(2, i)
            filtered(3, idx) = fullData(3, i)
            filtered(4, idx) = fullData(4, i)
            filtered(5, idx) = fullData(5, i)
        End If
    Next i
    
    If idx > 0 Then
        Call PopulateListView(filtered, idx)
    Else
        Me.ListView1.ListItems.Clear
        Me.TextBox3.value = 0
    End If
End Sub

Private Sub TextBox5_Change()
    Dim i As Long, idx As Long
    Dim filtered() As Variant
    Dim searchStatus As String
    
    searchStatus = Trim(Me.TextBox5.value)
    
    ' ≈–« ›«—€° ⁄—÷ ﬂ· «·»Ì«‰« 
    If searchStatus = "" Then
        Call PopulateListView(fullData, UBound(fullData, 2))
        Exit Sub
    End If
    
    ReDim filtered(1 To 5, 1 To UBound(fullData, 2))
    idx = 0
    
    For i = 1 To UBound(fullData, 2)
        ' «·⁄„Êœ 3 ÌÕ ÊÌ ”»» «·≈Õ«·…
        If InStr(1, fullData(3, i), searchStatus, vbTextCompare) > 0 Then
            idx = idx + 1
            filtered(1, idx) = fullData(1, i)
            filtered(2, idx) = fullData(2, i)
            filtered(3, idx) = fullData(3, i)
            filtered(4, idx) = fullData(4, i)
            filtered(5, idx) = fullData(5, i)
        End If
    Next i
    
    If idx > 0 Then
        Call PopulateListView(filtered, idx)
    Else
        Me.ListView1.ListItems.Clear
        Me.TextBox3.value = 0
    End If
End Sub

