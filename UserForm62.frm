VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm62 
   Caption         =   "UserForm62"
   ClientHeight    =   11880
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   25080
   OleObjectBlob   =   "UserForm62.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm62"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fullEmployeeList() As String
Dim empMasterArr As Variant
Dim fullNameMasterArr() As String
Dim dataLoaded As Boolean
Dim lvInitialized As Boolean


Private Sub CommandButton1_Click()
         BackupAndPrepareTw6een

    Dim wsS As Worksheet, wsD As Worksheet
    Dim srcArr As Variant, outArr() As Variant
    Dim lastRowS As Long, i As Long, j As Long
    Dim colMap As Variant
    Dim colNumbers() As Long
    Dim nameArr As Variant
    Dim hazardTotal As Double
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    Set wsS = Sheets("Sheet1")
    Set wsD = Sheets("tw6een")
    
    ' ======= —ƒÊ” «·√⁄„œ… ›Ì «·’› 8 =======
    Dim headers As Variant
    headers = Array("«·—ﬁ„ «·ÊŸÌ›Ì", "«·«”„", "«·«»", "«·Ãœ", "«··ﬁ»", "‰Ê⁄ «·Œœ„…", _
                    "Õ«·… «·Œœ„…", "«·œ—Ã… «·ÊŸÌ›Ì…", "«·⁄‰Ê«‰ «·ÊŸÌ›Ì", "«·‘Â«œ…", _
                    "«·—« »", "«·÷—Ì»…", "„Ã„Ê⁄ «·«” ﬁÿ«⁄« ", "«·„‰’»", "„Œ’’«  “ÊÃÌ…", _
                    "„Œ’’«  «·«ÿ›«·", "„Œ’’«  «·‘Â«œ…", "„Œ’’«  «·„‰’»", "«·ŒÿÊ—…", "«” À‰«∆Ì…", _
                    "—ﬁ«»Ì…", "«·Œœ„… «·Ã«„⁄Ì…", "„ﬁÿÊ⁄…", "«·„Êﬁ⁄ «·Ã€—«›Ì", "ÿ»«»…", _
                    "Õ—›…", "„Â‰Ì…", "—∆«”Ì…", "Â‰œ”Ì…", "Œœ„… Œ«—ÃÌ…", "«Œ—Ï", _
                    "„Ã„Ê⁄ «·„Œ’’« ", "«·—« » «·’«›Ì", "ÿ—Ìﬁ… «·œ›⁄", "IBAN")
    
    For j = 0 To UBound(headers)
        wsD.Cells(8, j + 1).value = headers(j)
    Next j
    
    ' =======  ·ÊÌ‰ «·’› 8 =======
    With wsD.Rows(8)
        .Interior.Color = RGB(220, 220, 220) ' —„«œÌ ›« Õ
        .Font.Color = RGB(0, 0, 128) ' ‰Ì·Ì / √“—ﬁ €«„ﬁ
        .Font.Bold = True
    End With
    
    ' ======= „”Õ «·»Ì«‰«  «·ﬁœÌ„… =======
    wsD.Rows("9:" & wsD.Rows.count).ClearContents
    
    ' ======= ﬁ—«¡… »Ì«‰«  Sheet1 =======
    lastRowS = wsS.Cells(wsS.Rows.count, "A").End(xlUp).row
    If lastRowS < 9 Then Exit Sub
    srcArr = wsS.Range("A9", wsS.Cells(lastRowS, wsS.Cells(8, wsS.Columns.count).End(xlToLeft).Column)).value
    
    ' ======= ≈⁄œ«œ „’›Ê›… «·≈Œ—«Ã =======
    ReDim outArr(1 To UBound(srcArr), 1 To 35)
    
    ' ======= Œ—Ìÿ… «·√⁄„œ… =======
    colMap = Array("B", "E", "E", "E", "E", "BF", "BF", "H", "F", "⁄„Êœ ÃœÌœ", _
                   "J", "X", "AV", "⁄„Êœ ÃœÌœ", "U", "T", "N", "R", "AA", _
                   "⁄„Êœ ÃœÌœ", "⁄„Êœ ÃœÌœ", "⁄„Êœ ÃœÌœ", "⁄„Êœ ÃœÌœ", "S", "⁄„Êœ ÃœÌœ", _
                   "⁄„Êœ ÃœÌœ", "AA", "⁄„Êœ ÃœÌœ", "P", "⁄„Êœ ÃœÌœ", "⁄„Êœ ÃœÌœ", "AD", "AW", "AY", "AZ")
    
    ReDim colNumbers(LBound(colMap) To UBound(colMap))
    For j = LBound(colMap) To UBound(colMap)
        If colMap(j) = "⁄„Êœ ÃœÌœ" Then
            colNumbers(j) = 0
        Else
            colNumbers(j) = wsS.Columns(colMap(j)).Column
        End If
    Next j
    
    ' ======= ‰”Œ «·»Ì«‰«  „⁄  ﬁ”Ì„ «·«”„ ÊÕ”«» «·ŒÿÊ—… =======
    For i = 1 To UBound(srcArr)
        '  ﬁ”Ì„ «·«”„
        nameArr = SplitFullNameSmart(CStr(srcArr(i, 5)))
        
        ' Ã„⁄ «·ŒÿÊ—…: Y + AC
        hazardTotal = 0
        If UBound(srcArr, 2) >= 25 And IsNumeric(srcArr(i, 25)) Then hazardTotal = hazardTotal + srcArr(i, 25) ' Y
        If UBound(srcArr, 2) >= 29 And IsNumeric(srcArr(i, 29)) Then hazardTotal = hazardTotal + srcArr(i, 29) ' AC
        
        For j = 1 To 35
            Select Case colMap(j - 1)
                Case "⁄„Êœ ÃœÌœ"
                    outArr(i, j) = ""
                Case "E"
                    Select Case j
                        Case 2: outArr(i, j) = nameArr(1)
                        Case 3: outArr(i, j) = nameArr(2)
                        Case 4: outArr(i, j) = nameArr(3)
                        Case 5: outArr(i, j) = nameArr(4)
                    End Select
                Case Else
                    If j = 19 Then
                        outArr(i, j) = hazardTotal ' «·⁄„Êœ S ›Ì tw6een
                    Else
                        If colNumbers(j - 1) > 0 And colNumbers(j - 1) <= UBound(srcArr, 2) Then
                            outArr(i, j) = srcArr(i, colNumbers(j - 1))
                        Else
                            outArr(i, j) = ""
                        End If
                    End If
            End Select
        Next j
    Next i
    
    ' ======= ﬂ «»… «·»Ì«‰«  œ›⁄… Ê«Õœ… =======
    wsD.Range("A9").Resize(UBound(outArr), 35).value = outArr
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox " „ ‰ﬁ· «·»Ì«‰«  Ê ·ÊÌ‰ —ƒÊ” «·√⁄„œ… ÊÕ”«» „Ã„Ê⁄ (Y + AC) ›Ì ⁄„Êœ S »‰Ã«Õ!", vbInformation
    ReloadTw6eenForm
End Sub

' ======= œ«·…  ﬁ”Ì„ «·«”„ =======
Function SplitFullNameSmart(fullName As String) As Variant
    Dim Parts() As String
    Dim result(1 To 4) As String
    Dim i As Long
    
   Parts = Split(Trim(fullName), " ")
    
    ' √Ê· 3 √Ã“«¡
   For i = 0 To 2
        If i <= UBound(Parts) Then
          result(i + 1) = Parts(i)
        Else
           result(i + 1) = ""
        End If
    Next i
    
    ' √Ì ﬂ·„… »⁄œ «·À·«À…  –Â» ··⁄„Êœ «·—«»⁄
    result(4) = ""
    For i = 3 To UBound(Parts)
        If result(4) = "" Then
            result(4) = Parts(i)
        Else
            result(4) = result(4) & " " & Parts(i)
        End If
    Next i
    
    SplitFullNameSmart = result
End Function

Private Sub CommandButton10_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton2_Click()
 Dim ctrl As Control
    '  „— ⁄·Ï ﬂ· ⁄‰«’— «·›Ê—„
    For Each ctrl In Me.Controls
        ' ≈–« ﬂ«‰  Ìﬂ” »Êﬂ”° ‰›—€Â
        If TypeName(ctrl) = "TextBox" Then
            ctrl.Text = ""
        End If
    Next ctrl
    
    ' Ê÷⁄ «· —ﬂÌ“ ⁄·Ï √Ê·  Ìﬂ” »Êﬂ” ··«œŒ«·
    Me.TextBox1.SetFocus
    
    MsgBox "Ã«Â“ ·≈œŒ«· ”Ã· ÃœÌœ.", vbInformation
End Sub

Private Sub CommandButton3_Click()
Dim wsTarget As Worksheet
    Dim newRow As Long
    Dim LastCol As Long
    Dim colIndex As Long
    
    Set wsTarget = ThisWorkbook.Sheets("tw6een")
    
    ' ≈ÌÃ«œ «·’› «·ÃœÌœ ›Ì ‰Â«Ì… «·»Ì«‰« 
    newRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row + 1
    
    ' ¬Œ— ⁄„Êœ „” Œœ„ ·—ƒÊ” «·√⁄„œ…
    LastCol = wsTarget.Cells(8, wsTarget.Columns.count).End(xlToLeft).Column
    
    ' ‰”Œ «·»Ì«‰«  „‰ TextBoxes ≈·Ï «·‘Ì 
    For colIndex = 1 To LastCol
        On Error Resume Next
        wsTarget.Cells(newRow, colIndex).value = Me.Controls("TextBox" & colIndex).Text
        On Error GoTo 0
    Next colIndex
    
    ' ≈⁄«œ… „·¡ ListBox »⁄œ ≈÷«›… «·”Ã· «·ÃœÌœ
    Call UserForm_Initialize ' Ì⁄Ìœ  ⁄»∆… fullEmployeeList Ê ListBox1
    
    MsgBox " „ ≈÷«›… «·”Ã· «·ÃœÌœ »‰Ã«Õ!", vbInformation
End Sub

Private Sub CommandButton4_Click()
Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim found As Boolean: found = False
    Dim r As Long
    Dim LastCol As Long
    Dim colIndex As Long
    Dim EmpID As String
    
    Set wsTarget = ThisWorkbook.Sheets("tw6een")
    
    EmpID = Trim(Me.TextBox1.Text) ' «› —÷ √‰ TextBox1 ÌÕ ÊÌ ⁄·Ï «·—ﬁ„ «·ÊŸÌ›Ì
    
    If EmpID = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «·—ﬁ„ «·ÊŸÌ›Ì · ÕœÌÀ «·»Ì«‰« .", vbExclamation
        Exit Sub
    End If
    
    ' «·»ÕÀ ⁄‰ «·—ﬁ„ «·ÊŸÌ›Ì ›Ì «·⁄„Êœ B »œ¡« „‰ «·’› 9
    lastRow = wsTarget.Cells(wsTarget.Rows.count, "A").End(xlUp).row
    
    For r = 9 To lastRow
        If Trim(wsTarget.Cells(r, "A").value) = EmpID Then
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        MsgBox "«·—ﬁ„ «·ÊŸÌ›Ì €Ì— „ÊÃÊœ ›Ì «·»Ì«‰« .", vbExclamation
        Exit Sub
    End If
    
    '  ÕœÌÀ «·√⁄„œ… Ê›ﬁ TextBoxes
    LastCol = wsTarget.Cells(8, wsTarget.Columns.count).End(xlToLeft).Column
    
    For colIndex = 1 To LastCol
        On Error Resume Next
        wsTarget.Cells(r, colIndex).value = Me.Controls("TextBox" & colIndex).Text
        On Error GoTo 0
    Next colIndex
    
    MsgBox " „  ÕœÌÀ «·»Ì«‰«  »‰Ã«Õ ··„ÊŸ› —ﬁ„: " & EmpID, vbInformation
    
    '  ÕœÌÀ ListBox »⁄œ «· ÕœÌÀ
    Call UserForm_Initialize
End Sub

Private Sub CommandButton6_Click()
Dim ws As Worksheet
    Dim lastRow As Long
    Dim found As Boolean: found = False
    Dim r As Long
    Dim LastCol As Long
    Dim colIndex As Long
    Dim EmpID As String
    
    Set ws = ThisWorkbook.Sheets("tw6een")
    
    EmpID = Trim(Me.TextBox1.Text) ' «·—ﬁ„ «·ÊŸÌ›Ì
    
    If EmpID = "" Then
        MsgBox "Ì—ÃÏ ≈œŒ«· «·—ﬁ„ «·ÊŸÌ›Ì.", vbExclamation
        Exit Sub
    End If
    
    ' «·»ÕÀ ⁄‰ «·—ﬁ„ «·ÊŸÌ›Ì ›Ì «·⁄„Êœ B
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    For r = 9 To lastRow
        If Trim(ws.Cells(r, "B").value) = EmpID Then
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then
        ' ≈÷«›… ”Ã· ÃœÌœ
        r = lastRow + 1
    End If
    
    ' ⁄œœ «·√⁄„œ… »‰«¡ ⁄·Ï —ƒÊ” «·’› 8
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' Õ›Ÿ «·»Ì«‰«  „‰ TextBoxes ≈·Ï «·‘Ì 
    For colIndex = 1 To LastCol
        On Error Resume Next
        ws.Cells(r, colIndex).value = Me.Controls("TextBox" & colIndex).Text
        On Error GoTo 0
    Next colIndex
    
    '  ÕœÌÀ ListBox √Ê ListView »⁄œ «·Õ›Ÿ
    Call UserForm_Initialize
    
    If found Then
        MsgBox " „  ÕœÌÀ »Ì«‰«  «·„ÊŸ› —ﬁ„: " & EmpID, vbInformation
    Else
        MsgBox " „ ≈÷«›… ”Ã· ÃœÌœ ··„ÊŸ› —ﬁ„: " & EmpID, vbInformation
    End If
End Sub

Private Sub CommandButton7_Click()
   
    Dim wsS As Worksheet, wsD As Worksheet
    Dim lastS As Long, lastD As Long
    Dim srcArr As Variant, destArr As Variant
    Dim dict As Object
    Dim i As Long, r As Long
    Dim EmpID As String
    Dim nameParts As Variant
    Dim hz As Double
    
    Set wsS = Sheets("Sheet1")
    Set wsD = Sheets("tw6een")
    
    lastS = wsS.Cells(wsS.Rows.count, "B").End(xlUp).row
    lastD = wsD.Cells(wsD.Rows.count, "A").End(xlUp).row
    
    If lastS < 9 Or lastD < 2 Then Exit Sub
    
    ' ?? «· ’ÕÌÕ «·„Â„:  Õ„Ì· «·„’›Ê›… ·€«Ì… BF
    srcArr = wsS.Range("A9:BF" & lastS).value
    destArr = wsD.Range("A2:AI" & lastD).value
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' ›Â—”… «·—ﬁ„ «·ÊŸÌ›Ì ›Ì tw6een
    For i = 1 To UBound(destArr, 1)
        EmpID = Trim(destArr(i, 1))
        If EmpID <> "" Then dict(EmpID) = i
    Next i
    
    ' «· ÕœÌÀ
    For i = 1 To UBound(srcArr, 1)
        EmpID = Trim(srcArr(i, 2)) ' B
        If dict.exists(EmpID) Then
            r = dict(EmpID)
            
            ' «·—ﬁ„ «·ÊŸÌ›Ì
            destArr(r, 1) = srcArr(i, 2)
            
            '  ﬁ”Ì„ «·«”„
            nameParts = SplitName4(srcArr(i, 5)) ' E
            destArr(r, 2) = nameParts(0)
            destArr(r, 3) = nameParts(1)
            destArr(r, 4) = nameParts(2)
            destArr(r, 5) = nameParts(3)
            
            ' «·‰”Œ »œÊ‰  ‘›Ì  ( —„Ì“ À«» )
            destArr(r, 6) = srcArr(i, 58) ' BF ‰Ê⁄ «·Œœ„…
            destArr(r, 7) = srcArr(i, 58) ' BF Õ«·… «·Œœ„…
            destArr(r, 8) = srcArr(i, 8)  ' H
            destArr(r, 9) = srcArr(i, 6)  ' F
            destArr(r, 11) = srcArr(i, 10) ' J
            destArr(r, 12) = srcArr(i, 24) ' X
            destArr(r, 13) = srcArr(i, 48) ' AV
            destArr(r, 15) = srcArr(i, 21) ' U
            destArr(r, 16) = srcArr(i, 20) ' T
            destArr(r, 17) = srcArr(i, 14) ' N
            destArr(r, 18) = srcArr(i, 18) ' R
            destArr(r, 24) = srcArr(i, 19) ' S
            destArr(r, 27) = srcArr(i, 27) ' AA
            destArr(r, 29) = srcArr(i, 16) ' P
            destArr(r, 32) = srcArr(i, 30) ' AD
            destArr(r, 33) = srcArr(i, 49) ' AW
            destArr(r, 34) = srcArr(i, 51) ' AY
            destArr(r, 35) = srcArr(i, 52) ' AZ
            
            ' «·ŒÿÊ—… = Y + AC
            hz = 0
            If IsNumeric(srcArr(i, 25)) Then hz = hz + srcArr(i, 25) ' Y
            If IsNumeric(srcArr(i, 29)) Then hz = hz + srcArr(i, 29) ' AC
            destArr(r, 19) = hz
            
        End If
    Next i
    
    wsD.Range("A2:AI" & lastD).value = destArr
    
    MsgBox " „ «· ÕœÌÀ »‰Ã«Õ »œÊ‰ Out of Range.", vbInformation
End Sub

' ===== œ«·…  ﬁ”Ì„ «·«”„ »√„«‰ =====
Function SplitName4(ByVal fullName As String) As Variant
    Dim tmp() As String
    Dim res(0 To 3) As String
    Dim i As Long
    
    fullName = Trim(fullName)
    If fullName = "" Then
        SplitName4 = res
        Exit Function
    End If
    
    tmp = Split(fullName, " ")
    
    For i = 0 To 2
        If i <= UBound(tmp) Then res(i) = tmp(i)
    Next i
    
    If UBound(tmp) >= 3 Then
        For i = 3 To UBound(tmp)
            res(3) = res(3) & IIf(res(3) = "", "", " ") & tmp(i)
        Next i
    End If
    
    SplitName4 = res
End Function

Private Sub CommandButton8_Click()
Dim ws As Worksheet
    Dim lastRow As Long
    Dim EmpID As String
    Dim i As Long
    
    Set ws = Sheets("tw6een")
    
    ' Ã·» «·—ﬁ„ «·ÊŸÌ›Ì
    EmpID = Trim(TextBox1.value) ' €Ì¯— «”„ «· ﬂ” »Êﬂ” ≈–« ·“„
    
    If EmpID = "" Then
        MsgBox "√œŒ· «·—ﬁ„ «·ÊŸÌ›Ì √Ê·«", vbExclamation
        Exit Sub
    End If
    
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).value) = EmpID Then
            ws.Rows(i).Delete
            MsgBox " „ Õ–› «·ﬁÌœ »‰Ã«Õ", vbInformation
            Exit Sub
        End If
    Next i
    
    MsgBox "«·—ﬁ„ «·ÊŸÌ›Ì €Ì— „ÊÃÊœ", vbCritical
End Sub

Private Sub CommandButton9_Click()
       BackupAndPrepareTw6een

  Dim wsDest As Worksheet
    Dim wsSource As Worksheet
    Dim lastRowSource As Long
    Dim lastColSource As Long
    Dim lastRowDest As Long
    Dim sourceData As Variant
    Dim destRange As Range
    
    ' ===  ÕœÌœ «·‘Ì  «·ÊÃÂ… ===
    Set wsDest = ThisWorkbook.Worksheets("tw6een")
    
    ' › Õ ‰«›–… «Œ Ì«— «·‘Ì  «·Œ«—ÃÌ
    Dim fName As Variant
    fName = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    If fName = False Then Exit Sub ' ≈–« √·€Ï «·„” Œœ„
    
    ' › Õ «·„·› «·Œ«—ÃÌ
    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(fName)
    Set wsSource = wbSource.Sheets(1) ' ‰› —÷ √‰ «·»Ì«‰«  »«·‘Ì  «·√Ê·
    
    ' ===  ÕœÌœ ¬Œ— ’› Ê⁄„Êœ ›Ì «·‘Ì  «·Œ«—ÃÌ ===
    lastRowSource = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    lastColSource = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column
    
    ' === «· Õﬁﬁ ≈–«  ÊÃœ »Ì«‰«  ===
    If lastRowSource < 2 Then
        MsgBox "«·‘Ì  «·Œ«—ÃÌ ·« ÌÕ ÊÌ ⁄·Ï »Ì«‰« .", vbExclamation, " ‰»ÌÂ"
        wbSource.Close False
        Exit Sub
    End If
    
    ' === ‰”Œ «·»Ì«‰«  „‰ «·’› 2 ≈·Ï ¬Œ— ’› ===
    sourceData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, lastColSource)).value
    
    ' ===  ÕœÌœ „ﬂ«‰ «··’ﬁ ›Ì ‘Ì  tw6een ===
    lastRowDest = wsDest.Cells(wsDest.Rows.count, 1).End(xlUp).row
    If lastRowDest < 8 Then lastRowDest = 8 ' «·’› 8 ÌÕ ÊÌ —ƒÊ” «·√⁄„œ…
    
    ' ‰ÿ«ﬁ «·ÊÃÂ… Ì»œ√ „‰ «·’› 9
    Set destRange = wsDest.Cells(9, 1).Resize(UBound(sourceData, 1), UBound(sourceData, 2))
    
    ' === ·’ﬁ «·»Ì«‰«  ›Ì ‘Ì  tw6een ===
    destRange.value = sourceData
    
    ' === Ê÷⁄ ⁄·«„… Ã«Â“Ì… «·»Ì«‰«  ===
    wsDest.Range("A1").value = "Data Ready" ' ⁄·«„… ÷„‰Ì… ·· Õ„Ì·
    
    ' === €·ﬁ «·„·› «·Œ«—ÃÌ ===
    wbSource.Close False
    
    ' === —”«·…  √ﬂÌœ ===
    MsgBox " „ «” Ì—«œ «·»Ì«‰«  »‰Ã«Õ „‰ «·‘Ì  «·Œ«—ÃÌ." & vbCrLf & _
           "Ê „  ÕœÌÀ Õ«·… «· Õ„Ì·.", vbInformation, " „"
    
    ' === «” œ⁄«¡ œ«·…  ÕœÌÀ UserForm ≈–« „ÊÃÊœ… ===
    On Error Resume Next
    Call RefreshTw6eenUserFormData ' Â‰« «› —÷ œ«·… „ÊÃÊœ… · ÕœÌÀ ListView √Ê TextBoxes
    ReloadTw6eenForm
    On Error GoTo 0
End Sub

Private Sub Label27_Click()

End Sub

Private Sub ListBox1_Click()
    Dim wsTarget As Worksheet
    Dim selectedRow As Long
    Dim LastCol As Long
    Dim colIndex As Long
    Dim clickedName As String
    
    Set wsTarget = ThisWorkbook.Sheets("tw6een")
    
    If Me.ListBox1.ListIndex = -1 Then Exit Sub
    
    clickedName = Me.ListBox1.List(Me.ListBox1.ListIndex)
    
    ' «·»ÕÀ ⁄‰ «·’› «·„ÿ«»ﬁ ›Ì ‘Ì  tw6een
    Dim r As Long
    Dim found As Boolean: found = False
    LastCol = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row
    
    For r = 9 To LastCol
        Dim fullNameInSheet As String
        fullNameInSheet = wsTarget.Cells(r, "B").value & " " & _
                          wsTarget.Cells(r, "C").value & " " & _
                          wsTarget.Cells(r, "D").value & " " & _
                          wsTarget.Cells(r, "E").value
        If fullNameInSheet = clickedName Then
            selectedRow = r
            found = True
            Exit For
        End If
    Next r
    
    If Not found Then Exit Sub
    
    ' ‰”Œ «·»Ì«‰«  „‰ «·‘Ì  ≈·Ï TextBoxes
    LastCol = wsTarget.Cells(8, wsTarget.Columns.count).End(xlToLeft).Column
    
    For colIndex = 1 To LastCol
        On Error Resume Next
        Me.Controls("TextBox" & colIndex).Text = wsTarget.Cells(selectedRow, colIndex).Text
        On Error GoTo 0
    Next colIndex
End Sub



Private Sub TextBox40_Change()
    Dim i As Long
    Dim searchText As String
    
    searchText = Trim(Me.TextBox40.Text)
    
    Me.ListBox1.Clear
    
    ' «” Œœ«„ «·„’›Ê›… String ›ﬁÿ -> Ì„‰⁄ Type mismatch
    For i = LBound(fullEmployeeList) To UBound(fullEmployeeList)
        If searchText = "" Or InStr(1, fullEmployeeList(i), searchText, vbTextCompare) > 0 Then
            Me.ListBox1.AddItem fullEmployeeList(i)
        End If
    Next i
    
 
'==========================================›· —… «··”  ›ÌÊ
      BuildAndFilter_ListView_All
    
End Sub




Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
On Error Resume Next

  Dim wsTw6een As Worksheet
    Dim lastDataRow1 As Long
    Dim totalColsCount As Long
    Dim employeeData As Variant
    
    '  ⁄ÌÌ‰ «·‘Ì 
    Set wsTw6een = ThisWorkbook.Worksheets("tw6een")
    
    '  ÕœÌœ ¬Œ— ’› ÌÕ ÊÌ »Ì«‰« 
    lastDataRow1 = wsTw6een.Cells(wsTw6een.Rows.count, 1).End(xlUp).row
    
    ' ⁄œœ «·√⁄„œ… ⁄·Ï √”«” —ƒÊ” «·√⁄„œ… ›Ì «·’› 8
    totalColsCount = wsTw6een.Cells(8, wsTw6een.Columns.count).End(xlToLeft).Column
    
    ' «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  „‰ «·’› 9 Ê„« »⁄œÂ
    If lastDataRow1 < 9 Then
        MsgBox "«·‰Ÿ«„ ·« ÌÕ ÊÌ ⁄·Ï √Ì »Ì«‰«  ›Ì ‘Ì  «· ÊÿÌ‰." & vbCrLf & _
               "Ì„ﬂ‰ﬂ «” Ì—«œ »Ì«‰«  „‰ ‘Ì  Œ«—ÃÌ.", vbExclamation, " ‰»ÌÂ"
        Exit Sub
    End If
    
    ' Ã·» «·»Ì«‰«  „‰ «·’› 9 ≈·Ï ¬Œ— ’›
    employeeData = wsTw6een.Range(wsTw6een.Cells(9, 1), wsTw6een.Cells(lastDataRow1, totalColsCount)).value
    
    '  ⁄»∆… ListView √Ê TextBoxes „‰ «·„’›Ê›…
    Call PopulateListView_Tw6een(employeeData)



















'==============================================================
   BuildAndFilter_ListView_All
 
 
Call FillListView
     
 
  
  
  
  
 
  
    Dim ws As Worksheet
    Dim LastCol As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("tw6een")
    
    ' ≈ÌÃ«œ ¬Œ— ⁄„Êœ „” Œœ„ ›Ì «·’› 8
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    '  ”·”· √”„«¡ «·√⁄„œ… ⁄·Ï «··Ì»·
    For i = 1 To LastCol
        On Error Resume Next ' ·Ê √Ì Label €Ì— „ÊÃÊœ Ì Ã«Ê“
        Me.Controls("Label" & i).Caption = ws.Cells(8, i).value
        On Error GoTo 0
    Next i
    
    '================================== ⁄»∆… «··”  »Êﬂ”
     Dim wsTarget As Worksheet
    Dim lastDataRow As Long
    Dim rowIndex As Long
    Dim employeeFullName As String
    
    Set wsTarget = ThisWorkbook.Sheets("tw6een")
    
    lastDataRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row
    
    '  ÕœÌœ ÕÃ„ «·„’›Ê›…
    ReDim fullEmployeeList(1 To lastDataRow - 8)
    
    Me.ListBox1.Clear
    
    For rowIndex = 9 To lastDataRow
        employeeFullName = wsTarget.Cells(rowIndex, "B").value & " " & _
                           wsTarget.Cells(rowIndex, "C").value & " " & _
                           wsTarget.Cells(rowIndex, "D").value & " " & _
                           wsTarget.Cells(rowIndex, "E").value
        Me.ListBox1.AddItem employeeFullName
        fullEmployeeList(rowIndex - 8) = employeeFullName '  Œ“Ì‰ «·«”„«¡ ›Ì «·„’›Ê›…
    Next rowIndex
End Sub

Private Sub FillEmployeeList()
    Dim wsTarget As Worksheet
    Dim lastDataRow As Long
    Dim rowIndex As Long
    Dim employeeFullName As String
    Dim searchText As String
    
    searchText = Trim(Me.TextBox40.Text)
    
    Set wsTarget = ThisWorkbook.Sheets("tw6een")
    lastDataRow = wsTarget.Cells(wsTarget.Rows.count, "B").End(xlUp).row
    
    ' „”Õ ListBox ﬁ»· ≈⁄«œ… „·¡
    Me.ListBox1.Clear
    
    For rowIndex = 9 To lastDataRow
        employeeFullName = wsTarget.Cells(rowIndex, "B").value & " " & _
                           wsTarget.Cells(rowIndex, "C").value & " " & _
                           wsTarget.Cells(rowIndex, "D").value & " " & _
                           wsTarget.Cells(rowIndex, "E").value
        If searchText = "" Or InStr(1, employeeFullName, searchText, vbTextCompare) > 0 Then
            Me.ListBox1.AddItem employeeFullName
        End If
    Next rowIndex
End Sub

Private Sub FillListView()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim LastCol As Long
    Dim r As Long, c As Long
    Dim itm As ListItem
    
    Set ws = ThisWorkbook.Sheets("tw6een")
    
    '  ÕœÌœ ¬Œ— ’› Ê¬Œ— ⁄„Êœ
    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    ' ≈⁄œ«œ ListView
    With Me.ListView1
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .ColumnHeaders.Clear
        
        ' ≈‰‘«¡ —ƒÊ” «·√⁄„œ…
        For c = 1 To LastCol
            .ColumnHeaders.Add Text:=ws.Cells(8, c).value, Width:=100
        Next c
        
        .ListItems.Clear
        
        ' ≈÷«›… «·»Ì«‰«  „‰ «·’› 9 Ê„« »⁄œÂ
        For r = 9 To lastRow
            Set itm = .ListItems.Add(Text:=ws.Cells(r, 1).value) ' «·⁄„Êœ «·√Ê·
            ' ≈÷«›… »ﬁÌ… «·√⁄„œ…
            For c = 2 To LastCol
                itm.ListSubItems.Add Text:=ws.Cells(r, c).value
            Next c
        Next r
    End With
End Sub


'============================== œ«·…  ﬁ”Ì„ «·«”„
'Function SplitFullNameSmart(fullName As String) As Variant
  '  Dim arr() As String, res(0 To 3) As String
 '   Dim i As Long, j As Long
    
  '  fullName = Trim(fullName)
  '  If fullName = "" Then
        ' ≈–« «·«”„ ›«—€° ‰—Ã⁄ „’›Ê›… 4 ⁄‰«’— ›«—€…
   '     For i = 0 To 3
       '     res(i) = ""
     '   Next i
     '   SplitFullNameSmart = res
     '   Exit Function
  '  End If
    
  '  arr = Split(fullName)
    
    ' √Ê· 3 √”„«¡ (√Ê √ﬁ· ≈–« «·«”„ ﬁ’Ì—)
  '  For i = 0 To 2
    '    If i <= UBound(arr) Then
      '      res(i) = arr(i)
     '   Else
      '      res(i) = ""
     '   End If
  ' Next i
    
    ' √Ì √”„«¡ „ »ﬁÌ… ? «··ﬁ»
  ' res(3) = ""
  '  If UBound(arr) >= 3 Then
     '   For j = 3 To UBound(arr)
       '     If res(3) = "" Then
            '    res(3) = arr(j)
         '   Else
         '       res(3) = res(3) & " " & arr(j)
         '   End If
      '  Next j
  ' End If
    
   ' SplitFullNameSmart = res
'End Function

Public Sub BuildAndFilter_ListView_All()

    Dim ws As Worksheet
    Dim empArr As Variant
    Dim lastRow As Long
    Dim totalCols As Long
    Dim SearchTxt As String
    Dim i2 As Long, c As Long
    Dim itm As ListItem
    Dim rowText As String
    Dim maxSubItems As Long
    Dim startRow As Long
    
    ' ’› «·»œ«Ì… «·ÕﬁÌﬁÌ… ··»Ì«‰« 
    startRow = 9

    ' === «· Õﬁﬁ „‰ ÊÃÊœ «·‘Ì  ===
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("tw6een")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "‘Ì  «· ÊÿÌ‰ 'tw6een' €Ì— „ÊÃÊœ!", vbCritical
        Exit Sub
    End If

    ' === «· Õﬁﬁ „‰ ⁄œœ «·√⁄„œ… ===
    totalCols = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    If totalCols = 0 Then
        MsgBox "’› «·⁄‰«ÊÌ‰ (’› 8) ›«—€!", vbExclamation
        Exit Sub
    End If

    ' === ¬Œ— ’› ===
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < startRow Then Exit Sub

    ' ===  Õ„Ì· «·»Ì«‰«  ≈·Ï „’›Ê›… »œ¡« „‰ «·’› 9 ===
    empArr = ws.Range(ws.Cells(startRow, 1), ws.Cells(lastRow, totalCols)).value

    ' ===  ÂÌ∆… ListView («·√⁄„œ… „—… Ê«Õœ… ›ﬁÿ) ===
    Static lvInitialized As Boolean
    If Not lvInitialized Then
        With ListView1
            .View = lvwReport
            .FullRowSelect = True
            .Gridlines = True
            .HideColumnHeaders = False
            .ColumnHeaders.Clear
            .ListItems.Clear
        End With

        For c = 1 To totalCols
            ListView1.ColumnHeaders.Add , , ws.Cells(8, c).value, 90
        Next c

        lvInitialized = True
    End If

    ' === „”Õ «·»Ì«‰«  ›ﬁÿ (»œÊ‰ ·„” «·√⁄„œ…) ===
    ListView1.ListItems.Clear

    ' === ‰’ «·»ÕÀ ===
    SearchTxt = Trim(TextBox40.Text)
    maxSubItems = ListView1.ColumnHeaders.count - 1

    ' ===  ⁄»∆… «·»Ì«‰«  „⁄ «·›· —… ===
    For i2 = 1 To UBound(empArr, 1)

        rowText = ""
        For c = 1 To UBound(empArr, 2)
            If Not IsError(empArr(i2, c)) And Not IsNull(empArr(i2, c)) Then
                rowText = rowText & " " & CStr(empArr(i2, c))
            End If
        Next c

        If SearchTxt = "" Or InStr(1, rowText, SearchTxt, vbTextCompare) > 0 Then
            Set itm = ListView1.ListItems.Add(, , SafeText(empArr(i2, 1)))
            For c = 2 To UBound(empArr, 2)
                If c - 1 <= maxSubItems Then
                    itm.SubItems(c - 1) = SafeText(empArr(i2, c))
                End If
            Next c
        End If

    Next i2

End Sub

Private Function SafeText(v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        SafeText = ""
    Else
        SafeText = CStr(v)
    End If
End Function
Public Sub PopulateListView_Tw6een(empArr As Variant)
    Dim lvItem As ListItem
    Dim r As Long, c As Long
    Dim totalColsCount As Long
    
    ' „”Õ √Ì »Ì«‰«  ﬁœÌ„… ›Ì «·‹ ListView
    With Tw6eenUserForm.ListView1
        .Clear
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        ' ≈÷«›… —ƒÊ” «·√⁄„œ… „‰ «·’› 8 ›Ì «·‘Ì 
        totalColsCount = UBound(empArr, 2)
        For c = 1 To totalColsCount
            .ColumnHeaders.Add Text:=Tw6eenUserForm.wsHeaders.Cells(8, c).value, Width:=100
        Next c
    End With
    
    '  ⁄»∆… «·‹ ListView »«·»Ì«‰«  „‰ «·’› 9 Ê„« »⁄œÂ
    For r = LBound(empArr, 1) To UBound(empArr, 1)
        Set lvItem = Tw6eenUserForm.ListView1.ListItems.Add(Text:=empArr(r, 1)) ' «·⁄„Êœ «·√Ê·
        For c = 2 To UBound(empArr, 2)
            lvItem.ListSubItems.Add Text:=empArr(r, c)
        Next c
    Next r
End Sub
Public Sub RefreshTw6eenUserFormData()
    ' ‰› —÷ √‰ UserForm «”„Â Tw6eenUserForm
    '  ﬁÊ„ Â–Â «·œ«·… » ⁄»∆… ListView √Ê TextBoxes „‰ ‘Ì  tw6een
    If UserFormIsLoaded("Tw6eenUserForm") Then
        Tw6eenUserForm.PopulateListView_Tw6een
    End If
End Sub

Public Function UserFormIsLoaded(ByVal formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formName Then
            UserFormIsLoaded = True
            Exit Function
        End If
    Next
    UserFormIsLoaded = False
End Function

'================================== ⁄“Ì“ «··”  »Êﬂ”


' ==== „’›Ê›…  Œ“Ì‰ «·»Ì«‰«  «·œ«Œ·Ì… ====


' ==== «” Ì—«œ «·»Ì«‰«  «·Œ«—ÃÌ… ====
Public Sub ImportTw6eenExternalData()
    Dim wsDest As Worksheet
    Dim wsSource As Worksheet
    Dim wbSource As Workbook
    Dim fName As Variant
    Dim lastRowSource As Long
    Dim lastColSource As Long
    Dim sourceData As Variant
    Dim destRange As Range
    
    ' ‘Ì  «· ÊÿÌ‰
    Set wsDest = ThisWorkbook.Worksheets("tw6een")
    
    ' «Œ Ì«— «·„·› «·Œ«—ÃÌ
    fName = Application.GetOpenFilename("Excel Files (*.xls*), *.xls*")
    If fName = False Then Exit Sub
    
    ' › Õ «·„·›
    Set wbSource = Workbooks.Open(fName)
    Set wsSource = wbSource.Sheets(1)
    
    ' ¬Œ— ’› Ê⁄„Êœ
    lastRowSource = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).row
    lastColSource = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column
    
    ' «· Õﬁﬁ „‰ «·»Ì«‰« 
    If lastRowSource < 2 Then
        MsgBox "«·‘Ì  «·Œ«—ÃÌ ·« ÌÕ ÊÌ ⁄·Ï »Ì«‰« .", vbExclamation
        wbSource.Close False
        Exit Sub
    End If
    
    ' Ã·» «·»Ì«‰«  „‰ «·’› 2 Õ Ï ¬Œ— ’›
    sourceData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRowSource, lastColSource)).value
    
    ' ·’ﬁ «·»Ì«‰«  ›Ì tw6een »œ¡« „‰ «·’› 9
    Set destRange = wsDest.Cells(9, 1).Resize(UBound(sourceData, 1), UBound(sourceData, 2))
    destRange.value = sourceData
    
    ' ≈€·«ﬁ «·„·› «·Œ«—ÃÌ
    wbSource.Close False
    
    ' ⁄·«„… ÷„‰Ì… »√‰ «·»Ì«‰«  Ã«Â“…
    wsDest.Range("A1").value = "Data Ready"
    
    '  ÕœÌÀ «·„’›Ê›… «·œ«Œ·Ì…
    Call UpdateTw6eenInternalArray
    
    '  ⁄»∆… ListView / ListBox
    Call PopulateListView_Tw6een
    
    MsgBox " „ «” Ì—«œ «·»Ì«‰«  «·Œ«—ÃÌ… Ê ÕœÌÀ «·ﬁ«∆„… »‰Ã«Õ.", vbInformation
End Sub

