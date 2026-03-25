VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm56 
   Caption         =   "UserForm56"
   ClientHeight    =   12615
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13560
   OleObjectBlob   =   "UserForm56.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm56"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












































Private Sub ComboBox3_Change()
If Me.ComboBox3.value <> "" Then
        Me.TextBox3.Text = Me.ComboBox3.value
    End If
End Sub

Private Sub ComboBox4_Change()
If Me.ComboBox4.value <> "" Then
        Me.TextBox4.Text = Me.ComboBox4.value
    End If
End Sub

Private Sub CommandButton1_Click()
On Error Resume Next
 Dim ws As Worksheet
    Dim dataArr As Variant
    Dim tempArr() As Variant
    Dim finalArr() As Variant
    Dim selectedCols As Variant
    Dim colNums() As Long
    Dim lastRow As Long, LastCol As Long
    Dim i As Long, j As Long, k As Long
    Dim BW As Long
    Dim filterValue As String
    Dim filterCol As Long
    Dim found As Boolean
    Dim cellValue As Variant

    ' ===== «· Õﬁﬁ „‰ «·√⁄„œ… =====
    If TextBox1.Text = "" Then
        MsgBox "«Œ «—Ì «·√⁄„œ… «· Ì  —ÌœÌ‰  —ÕÌ·Â«", vbExclamation
        Exit Sub
    End If

    ' ===== «· Õﬁﬁ „‰ «”„ «· ﬁ—Ì— =====
    If Trim(TextBox3.Text) = "" Then
        MsgBox "Ì—ÃÏ ﬂ «»… «”„ «· ﬁ—Ì— ›Ì TextBox3", vbExclamation
        Exit Sub
    End If

    ' ===== «· Õﬁﬁ „‰ «”„ «·ÃÂ… «·„«‰Õ… =====
    If Trim(TextBox4.Text) = "" Then
        MsgBox "Ì—ÃÏ ﬂ «»… «”„ «·ÃÂ… «·„«‰Õ… ›Ì TextBox4", vbExclamation
        Exit Sub
    End If

    ' Ê—ﬁ… «·»Ì«‰«  «·√’·Ì…
    Set ws = ThisWorkbook.Sheets("Sheet3")
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    '  Õ„Ì· «·»Ì«‰« 
    dataArr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, LastCol)).value

    ' ===== √⁄„œ… «· —ÕÌ· =====
    selectedCols = Split(TextBox1.Text, ",")
    ReDim colNums(LBound(selectedCols) To UBound(selectedCols))
    For i = LBound(selectedCols) To UBound(selectedCols)
        colNums(i) = ws.Range(Trim(selectedCols(i)) & "1").Column
    Next i

    ' ===== ›· —… Õ”» «·⁄„Êœ «·„Œ «— Ê«·ﬁÌ„… =====
    filterValue = Trim(TextBox2.Text)
    BW = 0
    ReDim tempArr(1 To lastRow - 1, 1 To UBound(colNums) + 1) ' -1 ·«Œ– «·»Ì«‰«  »œÊ‰ ’› «·⁄‰«ÊÌ‰

    If filterValue = "" Then
        ' »œÊ‰ ›· —…
        For i = 2 To lastRow
            BW = BW + 1
            For k = LBound(colNums) To UBound(colNums)
                tempArr(BW, k + 1) = dataArr(i, colNums(k))
            Next k
        Next i
    Else
        ' ›· —… Õ”» «·⁄„Êœ «·„Õœœ
        filterCol = Application.Match(ComboBox2.value, ws.Rows(1), 0)
        For i = 2 To lastRow
            found = False
            cellValue = dataArr(i, filterCol)
            If IsNumeric(cellValue) Then
                If cellValue = val(filterValue) Then found = True
            Else
                If InStr(1, CStr(cellValue), filterValue, vbTextCompare) > 0 Then found = True
            End If

            If found Then
                BW = BW + 1
                For k = LBound(colNums) To UBound(colNums)
                    tempArr(BW, k + 1) = dataArr(i, colNums(k))
                Next k
            End If
        Next i
    End If

    If BW = 0 Then
        MsgBox "·«  ÊÃœ »Ì«‰«  „ÿ«»ﬁ… ··›· —…", vbExclamation
        Exit Sub
    End If

    ' „’›Ê›… ‰Â«∆Ì… »«·ÕÃ„ «·ÕﬁÌﬁÌ
    ReDim finalArr(1 To BW, 1 To UBound(colNums) + 1)
    For i = 1 To BW
        For j = 1 To UBound(colNums) + 1
            finalArr(i, j) = tempArr(i, j)
        Next j
    Next i

    ' =====  —ÕÌ· «· ﬁ—Ì— =====
    Call ExportReport(finalArr, colNums, TextBox3.Text, TextBox4.Text)

    MsgBox " „ ≈‰‘«¡ «· ﬁ—Ì— Ê —ÕÌ·Â »‰Ã«Õ", vbInformation

End Sub

Function ColumnLetter(ColumnNumber As Long) As String
    Dim n As Long, s As String
    n = ColumnNumber
    s = ""
    Do While n > 0
        s = Chr(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetter = s
End Function





Private Sub CommandButton2_Click()
 Unload Me
UserForm10.Show
End Sub

Private Sub ListBox1_AfterUpdate()

End Sub

Private Sub ListBox1_Change()
Call UpdateListBox2
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  Dim selIndex As Long
    Dim colLetter As String
    Dim arr() As String
    Dim i As Long, newTxt As String

    If Me.ListBox1.ListIndex = -1 Then Exit Sub
    selIndex = Me.ListBox1.ListIndex

    ' Õ—› «·⁄„Êœ «·ÕﬁÌﬁÌ
    colLetter = Trim(Me.ListBox1.List(selIndex, 1))

    ' Õ–› «·Õ—› „‰ TextBox1
    If Trim(Me.TextBox1.Text) <> "" Then
        arr = Split(Me.TextBox1.Text, ",")
        newTxt = ""
        For i = LBound(arr) To UBound(arr)
            If Trim(arr(i)) <> colLetter Then
                newTxt = newTxt & Trim(arr(i)) & ","
            End If
        Next i
        If newTxt <> "" Then newTxt = Left(newTxt, Len(newTxt) - 1)
        Me.TextBox1.Text = newTxt
    End If

    ' Õ–› „‰ ListBox1 ›ﬁÿ
    Me.ListBox1.RemoveItem selIndex

    ' ?? ·« ‰·„” ListBox2 Â‰«
    ' ‰⁄Ìœ  Õ„Ì·Â »«·ﬂ«„·
    Call ReloadListBox2

End Sub



Private Sub UserForm_Initialize()
 Dim ws As Worksheet
    Dim LastCol As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Sheet3")
    LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' ComboBox1 ·√Œ Ì«— «·√⁄„œ… ·· —ÕÌ·
    ComboBox1.Clear
    For i = 1 To LastCol
        ComboBox1.AddItem ws.Cells(1, i).value
    Next i

    ' ComboBox2 ·√Œ Ì«— ⁄„Êœ «·›· —…
    ComboBox2.Clear
    For i = 1 To LastCol
        ComboBox2.AddItem ws.Cells(1, i).value
    Next i
    
    
     With Me.ComboBox3
        .AddItem " ﬁ—Ì— «·„‰ﬁÊ·Ì‰"
        .AddItem " ﬁ—Ì— «·„—Õ·Ì‰"
        .AddItem " ﬁ—Ì— «·„‰ﬁÊ·Ì‰ Ê«·„—Õ·Ì‰"
        .AddItem " ﬁ—Ì— «·„Õ«·Ì‰ ⁄·Ï «· ﬁ«⁄œ"
        
    End With
    
    '  ⁄»∆… ComboBox4 »√”„«¡ «·ÃÂ«  «·„«‰Õ…
    With Me.ComboBox4
        .AddItem "„ﬁ— ÂÌ∆… «·’‰«⁄«  «·Õ—»ÌÂ"
        .AddItem "„ﬁ— ‘—ﬂ… «·’‰«⁄«  «·Õ—»ÌÂ «·⁄«„…"
        .AddItem "„’‰⁄ «·ﬂ—«„… Ê«·Õ«—À"
        .AddItem "„’‰⁄ ⁄ﬁ»Â Ê»œ—"
        .AddItem "„’‰⁄ «·ﬁ«œ”Ì…"
        .AddItem "„’‰⁄ ÕÿÌ‰"
        .AddItem "„’‰⁄ «·Ì—„Êﬂ"
        .AddItem "„’‰⁄ «·‰Â—Ê«‰"
        .AddItem "„’‰⁄ «·—‘Ìœ"
        .AddItem "„’‰⁄ «·—»Ì⁄"
        .AddItem "„’‰⁄  »Êﬂ"
        .AddItem "„’‰⁄ Ã«»— »‰ ÕÌ«‰"
        .AddItem "„’‰⁄ «·ﬁ⁄ﬁ«⁄"
        .AddItem "„’‰⁄ Õ„Ê—«»Ì"
    End With
    
    
    
    
    
End Sub
Private Sub ComboBox1_Change()
Dim ws As Worksheet
    Dim colName As String
    Dim colLetter As String
    Dim colIndex As Long
    Dim i As Long
    Dim exists As Boolean
    Dim txtArr() As String
    Dim j As Long

    Set ws = ThisWorkbook.Sheets(3)

    colName = Trim(Me.ComboBox1.value)
    If colName = "" Then Exit Sub

    ' «·»ÕÀ ⁄‰ «·⁄„Êœ Õ”» «·«”„ ›Ì «·’› «·√Ê·
    colIndex = 0
    For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        If Trim(ws.Cells(1, i).value) = colName Then
            colIndex = i
            colLetter = Split(ws.Cells(1, i).Address, "$")(1)
            Exit For
        End If
    Next i

    If colIndex = 0 Then
        MsgBox "·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·⁄„Êœ ›Ì «·‘Ì ", vbExclamation
        Exit Sub
    End If

    ' =========================
    ' ≈÷«›… «”„ «·⁄„Êœ ≈·Ï ListBox1
    ' =========================
    exists = False
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.List(i, 0) = colName Then
            exists = True
            Exit For
        End If
    Next i

    If Not exists Then
        Me.ListBox1.AddItem colName
        Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = colLetter
    End If

    ' =========================
    ' ≈÷«›… Õ—› «·⁄„Êœ ≈·Ï TextBox1 „⁄ ›«’·…
    ' =========================
    If Trim(Me.TextBox1.Text) = "" Then
        Me.TextBox1.Text = colLetter
    Else
        txtArr = Split(Me.TextBox1.Text, ",")
        exists = False
        For j = LBound(txtArr) To UBound(txtArr)
            If Trim(txtArr(j)) = colLetter Then
                exists = True
                Exit For
            End If
        Next j

        If Not exists Then
            Me.TextBox1.Text = Me.TextBox1.Text & ", " & colLetter
        End If
    End If

   
    
    
    
    
   Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets(3)
    
    '  Õ„Ì· «·√⁄„œ… «·„Œ «—… ›Ì ListBox2 „⁄ «”„ «·⁄„Êœ ›Êﬁ «·»Ì«‰« 
    Call LoadSelectedColumnsWithHeaderAbove(ws1, Me.ListBox1, Me.ListBox2)
  
End Sub
Public Function IsSummableColumn(rng As Range) As Boolean
    Dim c As Range
    Dim numCount As Long, textCount As Long
    
    '  ÂÌ∆… «·⁄œ«œ« 
    numCount = 0
    textCount = 0
    
    ' ›Õ’ ﬂ· Œ·Ì… ›Ì «·⁄„Êœ
    For Each c In rng
        If c.value <> "" Then
            If IsNumeric(c.value) Then
                numCount = numCount + 1
            Else
                textCount = textCount + 1
            End If
        End If
    Next c
    
    ' «·ﬁ—«—: ≈–« ﬂ«‰ ⁄œœ «·ﬁÌ„ «·—ﬁ„Ì… √ﬂ»— „‰ «·‰’Ì… => Ì„ﬂ‰ Ã„⁄Â«
    If numCount > textCount Then
        IsSummableColumn = True
    Else
        IsSummableColumn = False
    End If
End Function

'==================================
Public Sub LoadSelectedColumnsWithHeaderAbove(ws As Worksheet, lstCols As MSForms.ListBox, lstData As MSForms.ListBox)
    On Error Resume Next
    Dim colCount As Long
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim colName As String
    Dim colIndex As Long
    Dim rowValues() As Variant
    Dim rowNum As Long
    Dim maxRows As Long
    
    ' ≈–« ·« ÌÊÃœ √⁄„œ… „Œ «—…
    If lstCols.ListCount = 0 Then Exit Sub
    
    colCount = lstCols.ListCount
    
    ' ≈ÌÃ«œ ¬Œ— ’› »‰«¡ ⁄·Ï ﬂ· «·√⁄„œ… «·„Œ «—…
    lastRow = 0
    For j = 0 To colCount - 1
        colName = lstCols.List(j)
        colIndex = 0
        For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
            If Trim(ws.Cells(1, i).value) = Trim(colName) Then
                colIndex = i
                Exit For
            End If
        Next i
        If colIndex > 0 Then
            lastRow = Application.WorksheetFunction.Max(lastRow, ws.Cells(ws.Rows.count, colIndex).End(xlUp).row)
        End If
    Next j
    
    ' ⁄œœ «·’›Ê› «·»Ì«‰«  (»«” À‰«¡ «”„ «·⁄„Êœ)
    maxRows = lastRow
    
    '  ›—Ì€ ListBox2
    lstData.Clear
    
    ' ÷»ÿ ⁄œœ «·√⁄„œ…
    lstData.ColumnCount = colCount
    
    ' ≈‰‘«¡ ’›Ê› ›Ì ListBox2 »ÕÌÀ ÌﬂÊ‰ «·’› «·√Ê· ·ﬂ· ⁄„Êœ «”„ «·⁄„Êœ
    For rowNum = 1 To maxRows ' rowNum = 1 -> «”„ «·⁄„Êœ
        ReDim rowValues(0 To colCount - 1)
        For j = 0 To colCount - 1
            colName = lstCols.List(j)
            
            ' ≈ÌÃ«œ —ﬁ„ «·⁄„Êœ
            colIndex = 0
            For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
                If Trim(ws.Cells(1, i).value) = Trim(colName) Then
                    colIndex = i
                    Exit For
                End If
            Next i
            
            If colIndex > 0 Then
                If rowNum = 1 Then
                    rowValues(j) = colName ' «·’› «·√Ê· = «”„ «·⁄„Êœ
                Else
                    rowValues(j) = ws.Cells(rowNum, colIndex).value
                End If
            Else
                rowValues(j) = ""
            End If
        Next j
        
        ' ≈÷«›… «·’› ≈·Ï ListBox2
        lstData.AddItem rowValues(0)
        For j = 1 To colCount - 1
            lstData.List(lstData.ListCount - 1, j) = rowValues(j)
        Next j
    Next rowNum
End Sub
Private Sub btnLoadColumns_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(3)
    
    '  Õ„Ì· «·√⁄„œ… «·„Œ «—… ›Ì ListBox2 „⁄ «”„ «·⁄„Êœ ›Êﬁ «·»Ì«‰« 
    Call LoadSelectedColumnsWithHeaderAbove(ws, Me.ListBox1, Me.ListBox2)
End Sub

Sub ExportReport(arr As Variant, colNums() As Long, reportName As String, deptName As String)

    Dim wb As Workbook, ws As Worksheet
    Dim startRow As Long, lastDataRow As Long, footerRow As Long
    Dim saveFolder As String, savePath As String
    Dim k As Long, headerArr() As Variant
    Dim fileNum As Long
    Dim totalCols As Long
    Dim colLetter As String
    Dim i As Long
    Dim excludeCols As Variant
    Dim tableStartRow As Long
    Dim numericCount As Long
    Dim cellVal As Variant
    Dim majorityNumeric As Boolean

    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)

    ' ===== ≈⁄œ«œ «·’›Õ… A4 =====
    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(1)
        .RightMargin = Application.CentimetersToPoints(1)
        .TopMargin = Application.CentimetersToPoints(1.5)
        .BottomMargin = Application.CentimetersToPoints(1.5)
    End With
    ws.DisplayRightToLeft = True

    totalCols = UBound(colNums) + 1
    startRow = 1

    ' ===== «· —ÊÌ”… 3 ’›Ê› „œ„Ã… =====
    ws.Range("A" & startRow & ":" & ColumnLetter(totalCols) & (startRow + 4)).Merge
    ws.Range("A" & startRow).value = Trim(reportName) & "          ‘Â— " & Month(Date) & " / ”‰… " & Year(Date) & "          " & Trim(deptName)
    ws.Range("A" & startRow).Font.Bold = True
    ws.Range("A" & startRow).Font.Size = 18
    ws.Range("A" & startRow).HorizontalAlignment = xlDistributed
    ws.Range("A" & startRow).VerticalAlignment = xlCenter
    ws.Range("A" & startRow).WrapText = True

    ' ===== √”„«¡ «·√⁄„œ… =====
    tableStartRow = 8
    ReDim headerArr(1 To 1, 1 To totalCols)
    For k = LBound(colNums) To UBound(colNums)
        headerArr(1, k + 1) = ThisWorkbook.Sheets("Sheet3").Cells(1, colNums(k)).value
    Next k
    ws.Range("A" & tableStartRow).Resize(1, totalCols).value = headerArr
    ws.Rows(tableStartRow).Font.Bold = True
    ws.Rows(tableStartRow).HorizontalAlignment = xlCenter

    ' ===== ﬂ «»… «·»Ì«‰«  =====
    ws.Range("A" & tableStartRow + 1).Resize(UBound(arr, 1), totalCols).value = arr
    ws.Columns.AutoFit

    ' ===== ﬂ·„… «·„Ã„Ê⁄ «·ﬂ·Ì =====
    lastDataRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    ws.Cells(lastDataRow + 1, 1).value = "«·„Ã„Ê⁄ «·ﬂ·Ì"
    ws.Cells(lastDataRow + 1, 1).Font.Bold = True
    ws.Cells(lastDataRow + 1, 1).Font.Size = 18

    ' ===== «·√⁄„œ… «·„” À‰«… =====
    excludeCols = Array("A", "B", "H", "I", "CA", "K", "L")

    ' ===== «·Ã„⁄ «·–ﬂÌ «·„Õ”‰ „⁄ «·√€·»Ì… =====
    For k = LBound(colNums) To UBound(colNums)
        colLetter = ColumnLetter(colNums(k))
        
        '  Õﬁﬁ „‰ «·√⁄„œ… «·„” À‰«…
        If Not IsError(Application.Match(ColumnLetter(colNums(k)), excludeCols, 0)) Then GoTo SkipColumn
        
        numericCount = 0
        For i = 1 To UBound(arr, 1)
            cellVal = arr(i, k + 1)
            If IsNumeric(cellVal) Then
                ' «· Õﬁﬁ „‰ «·√—ﬁ«„ «·’ÕÌÕ… √Ê ⁄‘—Ì… ? 2 „‰«“·
                If InStr(1, CStr(cellVal), ".") > 0 Then
                    If Len(Split(CStr(cellVal), ".")(1)) <= 2 Then
                        numericCount = numericCount + 1
                    End If
                Else
                    numericCount = numericCount + 1
                End If
            End If
        Next i
        
        ' ≈–« «·√€·»Ì… >50% ’«·Õ… ··Ã„⁄
        If numericCount >= (UBound(arr, 1) / 2) Then
            ws.Cells(lastDataRow + 1, colNums(k)).Formula = "=SUM(" & colLetter & tableStartRow + 1 & ":" & colLetter & lastDataRow & ")"
        End If

SkipColumn:
    Next k

    ' ===== «· –ÌÌ· =====
    footerRow = lastDataRow + 3
    ws.Range("A" & footerRow & ":" & ColumnLetter(totalCols) & (footerRow + 4)).Merge
    ws.Range("A" & footerRow).value = "«·„‰Ÿ„          «·„Õ«”»          «·„œﬁﬁ          —∆Ì” «·„œﬁﬁÌ‰          „œÌ— «·ﬁ”„          «·„œÌ—"
    ws.Range("A" & footerRow).Font.Size = 18
    ws.Range("A" & footerRow).Font.Bold = True
    ws.Range("A" & footerRow).HorizontalAlignment = xlDistributed
    ws.Range("A" & footerRow).VerticalAlignment = xlCenter
    ws.Range("A" & footerRow).WrapText = True

    ' ===== Õ›Ÿ «·„·› =====
    saveFolder = "D:\export1\"
    If Dir(saveFolder, vbDirectory) = "" Then MkDir saveFolder
    If Right(reportName, 5) <> ".xlsx" Then reportName = reportName & ".xlsx"
    savePath = saveFolder & reportName
    fileNum = 1
    Do While Dir(savePath) <> ""
        savePath = saveFolder & Left(reportName, Len(reportName) - 5) & "_" & fileNum & ".xlsx"
        fileNum = fileNum + 1
    Loop

    wb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
    wb.Close False

End Sub

' ===== œ«·… „”«⁄œ… · ÕÊÌ· —ﬁ„ «·⁄„Êœ ≈·Ï Õ—› =====

Private Sub AutoFitListBox(lb As MSForms.ListBox)

    Dim col As Long, row As Long
    Dim maxLen As Long
    Dim txt As String
    Dim widths As String

    For col = 0 To lb.ColumnCount - 1
        maxLen = 0

        For row = 0 To lb.ListCount - 1
            txt = CStr(lb.List(row, col))
            If Len(txt) > maxLen Then maxLen = Len(txt)
        Next row

        '  ﬁ—Ì» ⁄—÷ „‰«”» (7 ‰ﬁ«ÿ ·ﬂ· Õ—›)
        widths = widths & (maxLen * 7 + 10) & ";"
    Next col

    lb.ColumnWidths = widths

End Sub


Public Sub AddHeadersToListBox2()
    Dim i As Long
    Dim colCount As Long
    
    '  √ﬂœ √‰ srcArr „ÊÃÊœ… Ê„’›Ê›…
    If Not IsArray(srcArr) Then Exit Sub
    
    '  ÕœÌœ ⁄œœ «·√⁄„œ…
    colCount = UBound(srcArr, 2)
    
    ' „”Õ √Ì „Õ ÊÏ ”«»ﬁ
    Me.ListBox2.Clear
    
    ' ÷»ÿ ⁄œœ «·√⁄„œ… ›Ì ListBox2
    Me.ListBox2.ColumnCount = colCount
    
    ' ≈÷«›… ’› ··—ƒÊ”
    Me.ListBox2.AddItem
    For i = 1 To colCount
        Me.ListBox2.List(0, i - 1) = srcArr(1, i)  ' «·’› «·√Ê· ›Ì srcArr = —ƒÊ” «·√⁄„œ…
    Next i
    
    ' ≈÷«›… ’› ›«—€ √Ê ›«’· ≈–«  Õ» («Œ Ì«—Ì)
    ' Me.ListBox2.AddItem
    ' For i = 1 To colCount
    '     Me.ListBox2.List(1, i - 1) = String(12, "-")
    ' Next i
End Sub


Public Sub UpdateListBox2()
    Dim totalCols As Long
    Dim rowCount As Long
    Dim arr() As Variant
    Dim i As Long, j As Long
    
    ' ‰› —÷ √‰ ListBox1 ·œÌÂ« »Ì«‰« 
    rowCount = Me.ListBox1.ListCount
    totalCols = Me.ListBox1.ColumnCount
    
    ' «·Õœ «·√ﬁ’Ï 80 ⁄„Êœ
    If totalCols > 80 Then totalCols = 80
    
    ' ≈⁄œ«œ „’›Ê›… À‰«∆Ì… «·√»⁄«œ
    ReDim arr(0 To rowCount - 1, 0 To totalCols - 1)
    
    ' ‰”Œ «·»Ì«‰«  „‰ ListBox1 ≈·Ï «·„’›Ê›…
    For i = 0 To rowCount - 1
        For j = 0 To totalCols - 1
            arr(i, j) = Me.ListBox1.List(i, j)
        Next j
    Next i
    
    '  ⁄ÌÌ‰ «·„’›Ê›… ·‹ ListBox2
    Me.ListBox2.Clear
    Me.ListBox2.ColumnCount = totalCols
    Me.ListBox2.List = arr
    
    ' ÷»ÿ ⁄—÷ «·√⁄„œ… (80 ‰ﬁÿ… ·ﬂ· ⁄„Êœ ﬂ„À«·)
    Dim widths As String
    widths = String(totalCols * 4, " ") ' ›ﬁÿ „ﬂ«‰ ›«—€ „ƒﬁ 
    widths = ""
    For j = 1 To totalCols
        widths = widths & "80;"
    Next j
    widths = Left(widths, Len(widths) - 1)
    Me.ListBox2.ColumnWidths = widths
End Sub



Private Sub TextBox2_Change()
On Error Resume Next
   Dim ws As Worksheet
    Dim dataArr As Variant
    Dim resultArr() As Variant
    Dim headerArr() As Variant
    Dim lastRow As Long, LastCol As Long
    Dim filterCol As Long
    Dim filterValue As String
    Dim i As Long, j As Long, r As Long
    Dim cellValue As Variant

    ' ???? ?????
    If Me.ComboBox2.value = "" Then Exit Sub

    filterValue = Trim(Me.TextBox2.Text)

    Set ws = ThisWorkbook.Sheets("Sheet3")   ' ????? ??????

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' ????? ???????? ????? (???????? + ????????)
    dataArr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, LastCol)).value

    ' ????? ???? ??????? ?? ??? ??????
    filterCol = Application.Match(Me.ComboBox2.value, ws.Rows(1), 0)
    If IsError(filterCol) Then Exit Sub

    ' ????? ListBox2
    With Me.ListBox2
        .Clear
        .ColumnCount = LastCol
        .ColumnWidths = String(LastCol, "80;")
    End With

    ' ===== ????? ???? ??????? =====
    ReDim headerArr(0 To 0, 0 To LastCol - 1)
    For j = 1 To LastCol
        headerArr(0, j - 1) = dataArr(1, j)
    Next j
    Me.ListBox2.List = headerArr

    ' ===== ??? ???? ????? ????? ???? ?? ???????? =====
    If filterValue = "" Then
        ReDim resultArr(0 To lastRow - 2, 0 To LastCol - 1)
        For i = 2 To lastRow
            For j = 1 To LastCol
                resultArr(i - 2, j - 1) = dataArr(i, j)
            Next j
        Next i

        Me.ListBox2.List = CombineArrays(headerArr, resultArr)
        Exit Sub
    End If

    ' ===== ????? ????? ??? ?????? ?????? ??? =====
    ReDim resultArr(1 To lastRow - 1, 1 To LastCol)
    r = 0

    For i = 2 To lastRow
        cellValue = dataArr(i, filterCol)

        If IsNumeric(cellValue) And IsNumeric(filterValue) Then
            If CDbl(cellValue) = CDbl(filterValue) Then
                r = r + 1
                For j = 1 To LastCol
                    resultArr(r, j) = dataArr(i, j)
                Next j
            End If
        Else
            If InStr(1, CStr(cellValue), filterValue, vbTextCompare) > 0 Then
                r = r + 1
                For j = 1 To LastCol
                    resultArr(r, j) = dataArr(i, j)
                Next j
            End If
        End If
    Next i

    ' ===== ??? ??????? =====
    If r > 0 Then
        Dim finalArr() As Variant
        ReDim finalArr(0 To r, 0 To LastCol - 1)

        ' ????
        For j = 1 To LastCol
            finalArr(0, j - 1) = dataArr(1, j)
        Next j

        ' ??????
        For i = 1 To r
            For j = 1 To LastCol
                finalArr(i, j - 1) = resultArr(i, j)
            Next j
        Next i

        Me.ListBox2.List = finalArr
    End If

End Sub

' œ«·… „”«⁄œ… ·œ„Ã „’›Ê› Ì‰
Private Function JoinArrays(arr1 As Variant, arr2 As Variant) As Variant
On Error Resume Next
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long
    r1 = UBound(arr1, 1)
    c1 = UBound(arr1, 2)
    r2 = UBound(arr2, 1)
    c2 = UBound(arr2, 2)
    
    Dim result() As Variant
    ReDim result(0 To r1 + r2, 0 To c1)
    
    Dim i As Long, j As Long
    
    ' ‰”Œ arr1
    For i = 0 To r1
        For j = 0 To c1
            result(i, j) = arr1(i, j)
        Next j
    Next i
    
    ' ‰”Œ arr2
    For i = 0 To r2
        For j = 0 To c2
            result(i + r1 + 1, j) = arr2(i, j)
        Next j
    Next i
    
    JoinArrays = result
End Function

Public Function CombineArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim r1 As Long, r2 As Long, c As Long
    Dim i As Long, j As Long
    r1 = UBound(arr1, 1)
    r2 = UBound(arr2, 1)
    c = UBound(arr1, 2)

    Dim res() As Variant
    ReDim res(0 To r1 + r2 + 1, 0 To c)

    For i = 0 To r1
        For j = 0 To c
            res(i, j) = arr1(i, j)
        Next j
    Next i

    For i = 0 To r2
        For j = 0 To c
            res(i + r1 + 1, j) = arr2(i, j)
        Next j
    Next i

    CombineArrays = res
End Function


Private Sub ReloadListBox2()

    Dim ws As Worksheet
    Dim srcArr As Variant
    Dim headers() As Variant
    Dim dataArr() As Variant
    Dim colIndex() As Long
    Dim i As Long, j As Long, r As Long
    Dim lastRow As Long, LastCol As Long
    Dim colCount As Long
    Dim m As Variant

    ' ===== ????? ?????? =====
    Set ws = ThisWorkbook.Sheets("Sheet3")

    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    LastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    ' ??? ???? ??? ??????
    If lastRow < 1 Then
        Me.ListBox2.Clear
        Exit Sub
    End If

    srcArr = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, LastCol)).value

    ' ??? ??????? ????????
    colCount = Me.ListBox1.ListCount
    If colCount = 0 Then
        Me.ListBox2.Clear
        Exit Sub
    End If

    ' ===== ??? ????? ??????? ?? ???????? =====
    ReDim colIndex(1 To colCount)
    For i = 1 To colCount
        m = Application.Match(Me.ListBox1.List(i - 1, 0), ws.Rows(1), 0)
        If IsError(m) Then
            MsgBox "«·⁄„Êœ €Ì— „ÊÃÊœ : " & Me.ListBox1.List(i - 1, 0), vbCritical
            Exit Sub
        End If
        colIndex(i) = CLng(m)
    Next i

    ' ===== ???? ??????? (??????) =====
    ReDim headers(0 To 0, 0 To colCount - 1)
    For i = 1 To colCount
        headers(0, i - 1) = srcArr(1, colIndex(i))
    Next i

    ' ===== ??? ???? ?????? (??? ????) =====
    If lastRow = 1 Then
        With Me.ListBox2
            .Clear
            .ColumnCount = colCount
            .List = headers
        End With
        Exit Sub
    End If

    ' ===== ???????? =====
    ReDim dataArr(0 To lastRow - 2, 0 To colCount - 1)
    r = 0
    For i = 2 To lastRow
        For j = 1 To colCount
            dataArr(r, j - 1) = srcArr(i, colIndex(j))
        Next j
        r = r + 1
    Next i

    ' ===== ????? ListBox2 =====
    With Me.ListBox2
        .Clear
        .ColumnCount = colCount
        .List = JoinArrays(headers, dataArr)
    End With

End Sub
