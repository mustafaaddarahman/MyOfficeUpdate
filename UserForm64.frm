VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm64 
   Caption         =   "UserForm64"
   ClientHeight    =   11940
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13188
   OleObjectBlob   =   "UserForm64.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ws As Worksheet              ' «·‘Ì  «·„’œ—
Dim srcArr As Variant            ' „’›Ê›… «·»Ì«‰« 
Dim selectedCols() As Long       ' √—ﬁ«„ «·√⁄„œ… «·„Œ «—…
Dim selectedHeaders() As String  ' √”„«¡ «·√⁄„œ… «·„Œ «—…

'=================== √Õœ«À ComboBox ===================
Private Sub ComboBox1_Change()
    Dim colName As String
    colName = Me.ComboBox1.value
    If colName <> "" Then
        Dim exists As Boolean: exists = False
        Dim i As Long
        For i = 0 To Me.ListBox1.ListCount - 1
            If Me.ListBox1.List(i) = colName Then exists = True: Exit For
        Next i
        If Not exists Then
            Me.ListBox1.AddItem colName
            
            ' «·Õ’Ê· ⁄·Ï Õ—› «·⁄„Êœ
            Dim colLetter As String
            Dim LastCol As Long
            LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
            For i = 1 To LastCol
                If ws.Cells(8, i).value = colName Then
                    colLetter = Split(ws.Cells(8, i).Address(True, False), "$")(0)
                    Exit For
                End If
            Next i
            
            ' ≈÷«›… «·Õ—› ≈·Ï TextBox1 „⁄ ›«’·…
            If Me.TextBox1.Text = "" Then
                Me.TextBox1.Text = colLetter
            Else
                Me.TextBox1.Text = Me.TextBox1.Text & "," & colLetter
            End If
        End If
    End If
End Sub

Private Sub ComboBox2_Change()
    ' ·«  ⁄œÌ· Â‰«° Ì” Œœ„ ··›· —…
End Sub

Private Sub ComboBox3_Change()
    If Me.ComboBox3.value <> "" Then Me.TextBox3.Text = Me.ComboBox3.value
End Sub

Private Sub ComboBox4_Change()
    If Me.ComboBox4.value <> "" Then Me.TextBox4.Text = Me.ComboBox4.value
End Sub

Private Sub CommandButton1_Click()
 Dim wsSrc As Worksheet
    Dim dataArr As Variant
    Dim tempArr() As Variant
    Dim finalArr() As Variant
    Dim selectedCols As Variant
    Dim colNums() As Long
    Dim lastRow As Long, LastCol As Long
    Dim i As Long, j As Long, k As Long
    Dim cz As Long
    Dim filterValue As String
    Dim filterCol As Long
    Dim found As Boolean
    Dim cellValue As Variant

    Const HEADER_ROW As Long = 8
    Const DATA_START_ROW As Long = 9

    If TextBox1.Text = "" Then
        MsgBox "«Œ «—Ì «·√⁄„œ… «·„—«œ  —ÕÌ·Â«", vbExclamation
        Exit Sub
    End If

    If Trim(TextBox3.Text) = "" Or Trim(TextBox4.Text) = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «”„ «· ﬁ—Ì— Ê«·ÃÂ…", vbExclamation
        Exit Sub
    End If

    Set wsSrc = ThisWorkbook.Sheets("tw6een")

    lastRow = wsSrc.Cells(wsSrc.Rows.count, 1).End(xlUp).row
    LastCol = wsSrc.Cells(HEADER_ROW, wsSrc.Columns.count).End(xlToLeft).Column

    dataArr = wsSrc.Range( _
        wsSrc.Cells(HEADER_ROW, 1), _
        wsSrc.Cells(lastRow, LastCol) _
    ).value

    selectedCols = Split(TextBox1.Text, ",")
    ReDim colNums(LBound(selectedCols) To UBound(selectedCols))

    For i = LBound(selectedCols) To UBound(selectedCols)
        colNums(i) = wsSrc.Range(Trim(selectedCols(i)) & HEADER_ROW).Column
    Next i

    filterValue = Trim(TextBox2.Text)
    cz = 0

    ReDim tempArr(1 To UBound(dataArr, 1) - 1, 1 To UBound(colNums) + 1)

    If filterValue = "" Then
        For i = 8 To UBound(dataArr, 1)
            cz = cz + 1
            For k = LBound(colNums) To UBound(colNums)
                tempArr(cz, k + 1) = dataArr(i, colNums(k))
            Next k
        Next i
    Else
        filterCol = Application.Match(ComboBox2.value, wsSrc.Rows(HEADER_ROW), 0)

        For i = 8 To UBound(dataArr, 1)
            found = False
            cellValue = dataArr(i, filterCol)

            If IsNumeric(cellValue) And IsNumeric(filterValue) Then
                found = (val(cellValue) = val(filterValue))
            Else
                found = InStr(1, CStr(cellValue), filterValue, vbTextCompare) > 0
            End If

            If found Then
                cz = cz + 1
                For k = LBound(colNums) To UBound(colNums)
                    tempArr(cz, k + 1) = dataArr(i, colNums(k))
                Next k
            End If
        Next i
    End If

    If cz = 0 Then
        MsgBox "·«  ÊÃœ »Ì«‰«  „ÿ«»ﬁ…", vbExclamation
        Exit Sub
    End If

    ReDim finalArr(1 To cz, 1 To UBound(colNums) + 1)
    For i = 1 To cz
        For j = 1 To UBound(colNums) + 1
            finalArr(i, j) = tempArr(i, j)
        Next j
    Next i

    Call ExportReportFiltered(finalArr, colNums, TextBox3.Text, TextBox4.Text)

    MsgBox " „ «· —ÕÌ· »‰Ã«Õ", vbInformation

End Sub

Private Sub CommandButton2_Click()
 Unload Me
UserForm10.Show
End Sub

'=================== ÕœÀ ≈“«·… «·√⁄„œ… „‰ ListBox1 ===================
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim idx As Long
    idx = Me.ListBox1.ListIndex
    If idx >= 0 Then
        Dim colName As String
        colName = Me.ListBox1.List(idx)
        Me.ListBox1.RemoveItem idx
        
        ' ≈“«·… «·Õ—› «·„ﬁ«»· „‰ TextBox1
        Dim colLetter As String
        Dim LastCol As Long
        LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
        Dim i As Long
        For i = 1 To LastCol
            If ws.Cells(8, i).value = colName Then
                colLetter = Split(ws.Cells(8, i).Address(True, False), "$")(0)
                Exit For
            End If
        Next i
        
        '  ÕœÌÀ TextBox1
        Dim letters() As String, newText As String
        letters = Split(Me.TextBox1.Text, ",")
        newText = ""
        For i = LBound(letters) To UBound(letters)
            If letters(i) <> colLetter Then
                If newText = "" Then
                    newText = letters(i)
                Else
                    newText = newText & "," & letters(i)
                End If
            End If
        Next i
        Me.TextBox1.Text = newText
    End If
End Sub

'===================  Õ„Ì· «·„’›Ê›… «·„’œ— ===================
Private Sub LoadSrcArray()
    Set ws = ThisWorkbook.Worksheets("tw6een")
    Dim lastRow As Long, LastCol As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    srcArr = ws.Range(ws.Cells(8, 1), ws.Cells(lastRow, LastCol)).value
End Sub

'=================== UserForm Initialize ===================
Private Sub UserForm_Initialize()
    LoadSrcArray
    
    ' ÷»ÿ ComboBox2 »√”„«¡ «·√⁄„œ… (›· —…)
    Me.ComboBox2.Clear
    Dim c As Long
    For c = 1 To UBound(srcArr, 2)
        Me.ComboBox2.AddItem srcArr(1, c)
    Next c
    
    ' ComboBox1 («Œ Ì«— «·√⁄„œ… ·· —ÕÌ·)
    Me.ComboBox1.Clear
    Dim LastCol As Long
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To LastCol
        Me.ComboBox1.AddItem ws.Cells(8, c).value
    Next c
    
    ' ComboBox3 («”„ «· ﬁ—Ì—)
    With Me.ComboBox3
        .AddItem " ﬁ—Ì— «·«÷«›…"
        .AddItem " ﬁ—Ì— «·ﬁ«∆„… «·„›—œÂ"
        .AddItem " ﬁ—Ì— ’—›Ì… «·—« »"
        .AddItem " ﬁ—Ì— «·„«” —"
        .AddItem " ﬁ—Ì— «· ’«—ÌÕ «·«„‰Ì…"
    End With
    
    ' ComboBox4 («·ÃÂ… «·„«‰Õ…)
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

'=================== ›· —… ListBox2 ⁄‰œ  €ÌÌ— TextBox2 ===================
Private Sub TextBox2_Change()
On Error Resume Next
    Dim filterValue As String, r As Long, i As Long, j As Long
    Dim validCols() As Long, colCount As Long, rowIndex As Long
    Dim filterColIndex As Long, cellValue As Variant, isNumericCol As Boolean
    
    If Not IsArray(srcArr) Then Exit Sub
    If Me.ListBox1.ListCount = 0 Then Exit Sub
    
    filterValue = Trim(Me.TextBox2.Text)
    colCount = Me.ListBox1.ListCount
    ReDim validCols(1 To colCount)
    
    ' Ã„⁄ «·√⁄„œ… «·„Œ «—…
    For i = 1 To colCount
        validCols(i) = -1
        For j = 1 To UBound(srcArr, 2)
            If srcArr(1, j) = Me.ListBox1.List(i - 1) Then
                validCols(i) = j
                Exit For
            End If
        Next j
    Next i
    
    ' «·⁄„Êœ ··›· —…
    filterColIndex = 0
    isNumericCol = False
    If Me.ComboBox2.ListIndex >= 0 Then
        For i = 1 To colCount
            If validCols(i) > 0 Then
                If srcArr(1, validCols(i)) = Me.ComboBox2.value Then
                    filterColIndex = validCols(i)
                    If IsNumeric(srcArr(2, filterColIndex)) Then isNumericCol = True
                    Exit For
                End If
            End If
        Next i
    End If
    
    ' ≈⁄œ«œ ListBox2
    Me.ListBox2.Clear
    Me.ListBox2.ColumnCount = colCount
    Me.ListBox2.AddItem "" ' —ƒÊ”
    For i = 1 To colCount
        Me.ListBox2.List(0, i - 1) = srcArr(1, validCols(i))
    Next i
    Me.ListBox2.AddItem "" ' ›«’·
    For i = 1 To colCount
        Me.ListBox2.List(1, i - 1) = String(12, "-")
    Next i
    
    ' ≈÷«›… «·»Ì«‰«  «·„›· —…
    rowIndex = 2
    For r = 2 To UBound(srcArr, 1)
        Dim AddRow As Boolean: AddRow = False
        If filterColIndex = 0 Or Len(filterValue) = 0 Then
            AddRow = True
        Else
            cellValue = srcArr(r, filterColIndex)
            If isNumericCol Then
                If val(cellValue) = val(filterValue) Then AddRow = True
            Else
                If InStr(1, CStr(cellValue), filterValue, vbTextCompare) > 0 Then AddRow = True
            End If
        End If
        
        If AddRow Then
            Me.ListBox2.AddItem CStr(srcArr(r, validCols(1)))
            For i = 2 To colCount
                Me.ListBox2.List(rowIndex, i - 1) = CStr(srcArr(r, validCols(i)))
            Next i
            rowIndex = rowIndex + 1
        End If
    Next r
End Sub

'===================  Õ„Ì· «·√⁄„œ… «·„Œ «—… ›Ì ListBox2 ===================
Public Sub LoadSelectedColumnsWithHeaderAbove(ws As Worksheet, lstCols As MSForms.ListBox, lstData As MSForms.ListBox)
    Dim colCount As Long, lastRow As Long, i As Long, j As Long
    Dim colName As String, colIndex As Long, rowValues() As Variant, rowNum As Long, maxRows As Long
    
    If lstCols.ListCount = 0 Then Exit Sub
    colCount = lstCols.ListCount
    lastRow = 0
    
    For j = 0 To colCount - 1
        colName = lstCols.List(j)
        colIndex = 0
        For i = 1 To ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
            If Trim(ws.Cells(8, i).value) = Trim(colName) Then
                colIndex = i
                Exit For
            End If
        Next i
        If colIndex > 0 Then
            lastRow = Application.WorksheetFunction.Max(lastRow, ws.Cells(ws.Rows.count, colIndex).End(xlUp).row)
        End If
    Next j
    
    maxRows = lastRow - 8 + 1
    lstData.Clear
    lstData.ColumnCount = colCount
    
    For rowNum = 1 To maxRows
        ReDim rowValues(0 To colCount - 1)
        For j = 0 To colCount - 1
            colName = lstCols.List(j)
            colIndex = 0
            For i = 1 To ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
                If Trim(ws.Cells(8, i).value) = Trim(colName) Then
                    colIndex = i
                    Exit For
                End If
            Next i
            If colIndex > 0 Then
                If rowNum = 1 Then
                    rowValues(j) = colName
                Else
                    rowValues(j) = ws.Cells(rowNum + 8 - 1, colIndex).value
                End If
            Else
                rowValues(j) = ""
            End If
        Next j
        lstData.AddItem rowValues(0)
        For j = 1 To colCount - 1
            lstData.List(lstData.ListCount - 1, j) = rowValues(j)
        Next j
    Next rowNum
End Sub

'=================== “— «· —ÕÌ· ===================


'=================== Õ›Ÿ «· ﬁ—Ì— ===================
Public Sub SaveWorkbookWithUniqueName(wb As Workbook, reportName As String)
    Dim saveFolder As String, savePath As String, fileNum As Long
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
    MsgBox " „ Õ›Ÿ «· ﬁ—Ì— »‰Ã«Õ ›Ì: " & savePath, vbInformation
End Sub


Public Sub ExportReportFiltered(arr As Variant, colNums() As Long, reportName As String, deptName As String)
    Dim wb As Workbook, ws As Worksheet, wsSrc As Worksheet
    Dim totalCols As Long, tableStartRow As Long, lastDataRow As Long, footerRow As Long
    Dim saveFolder As String, savePath As String, fileNum As Long
    Dim k As Long, i As Long, colLetter As String, cellVal As Variant
    Dim isNumericCol As Boolean, excludeCols As Variant
    
    Const HEADER_ROW As Long = 8
    Const DATA_START_ROW As Long = 9
    
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    Set wsSrc = ThisWorkbook.Sheets("tw6een")
    
    ws.DisplayRightToLeft = True
    totalCols = UBound(colNums) + 1
    tableStartRow = HEADER_ROW
    
    ' ===== «· —ÊÌ”… 5 ’›Ê› „œ„Ã… =====
    ws.Range("A1:" & ColumnLetter(totalCols) & "5").Merge
    ws.Range("A1").value = reportName & "    " & Format(Date, "MMMM yyyy") & "    " & deptName
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    ws.Range("A1").Borders.LineStyle = xlContinuous
    
    ' ===== —ƒÊ” «·√⁄„œ… =====
    For k = LBound(colNums) To UBound(colNums)
        ws.Cells(tableStartRow, k + 1).value = wsSrc.Cells(HEADER_ROW, colNums(k)).value
        ws.Cells(tableStartRow, k + 1).Font.Bold = True
        ws.Cells(tableStartRow, k + 1).Interior.Color = RGB(200, 200, 200)
        ws.Cells(tableStartRow, k + 1).Borders.LineStyle = xlContinuous
        ws.Cells(tableStartRow, k + 1).HorizontalAlignment = xlCenter
    Next k
    
    ' ===== «·»Ì«‰«  =====
    ws.Range(ws.Cells(tableStartRow + 1, 1), ws.Cells(tableStartRow + UBound(arr, 1), totalCols)).value = arr
    ws.Columns.AutoFit
    
    lastDataRow = tableStartRow + UBound(arr, 1)
    
    ' ===== ’› «·„Ã„Ê⁄ «·ﬂ·Ì =====
   
    
    ' «” À‰«¡ «·√⁄„œ… 2 Ê3 „‰ «·Ã„⁄
    excludeCols = Array(2, 3)
    
    For k = LBound(colNums) To UBound(colNums)
        If IsError(Application.Match(k + 1, excludeCols, 0)) Then
            isNumericCol = True
            '  Õﬁﬁ „‰ «·⁄„Êœ ≈–« ﬂ· «·ﬁÌ„ —ﬁ„Ì…
            For i = 1 To UBound(arr, 1)
                cellVal = arr(i, k + 1)
                If Not IsNumeric(cellVal) Then
                    isNumericCol = False
                    Exit For
                End If
            Next i
            ' ≈–« «·⁄„Êœ —ﬁ„Ì «Ã„⁄Â
            If isNumericCol Then
                colLetter = ColumnLetter(k + 1)
                ws.Cells(lastDataRow + 1, k + 1).Formula = "=SUM(" & colLetter & DATA_START_ROW & ":" & colLetter & lastDataRow & ")"
                ws.Cells(lastDataRow + 1, k + 1).Font.Bold = True
            End If
        End If
    Next k
    
     ws.Cells(lastDataRow + 1, 1).value = "«·„Ã„Ê⁄ «·ﬂ·Ì"
    ws.Cells(lastDataRow + 1, 1).Font.Bold = True
    ws.Cells(lastDataRow + 1, 1).Font.Size = 16
    ws.Cells(lastDataRow + 1, 1).HorizontalAlignment = xlCenter
    ' ===== «· –ÌÌ· 5 ’›Ê› „œ„Ã… =====
    footerRow = lastDataRow + 3
    ws.Range("A" & footerRow & ":" & ColumnLetter(totalCols) & (footerRow + 4)).Merge
    ws.Range("A" & footerRow).value = "«·„‰Ÿ„          «·„Õ«”»          «·„œﬁﬁ          —∆Ì” «·„œﬁﬁÌ‰          „œÌ— «·ﬁ”„          «·„œÌ—"
    ws.Range("A" & footerRow).Font.Size = 18
    ws.Range("A" & footerRow).Font.Bold = True
    ws.Range("A" & footerRow).HorizontalAlignment = xlCenter
    ws.Range("A" & footerRow).VerticalAlignment = xlCenter
    ws.Range("A" & footerRow).Borders.LineStyle = xlContinuous
    
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
    
    MsgBox " „ Õ›Ÿ «· ﬁ—Ì— »‰Ã«Õ ›Ì: " & savePath, vbInformation
End Sub

' œ«·…  ÕÊÌ· —ﬁ„ «·⁄„Êœ ≈·Ï Õ—›
Public Function ColumnLetter(ColumnNumber As Long) As String
    Dim n As Long, s As String
    n = ColumnNumber
    s = ""
    Do While n > 0
        s = Chr(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    ColumnLetter = s
End Function




