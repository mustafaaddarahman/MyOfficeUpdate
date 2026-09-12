VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm80 
   Caption         =   "UserForm80"
   ClientHeight    =   11028
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   15324
   OleObjectBlob   =   "UserForm80.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Dim Zoomer As New clsZoomManager
Public TempData As New Collection
Public CurrentTop As Double
Private ControlTop As Long

Private Sub btnCalcAll_Click()
    Dim ctrl As Control
    Dim empID As String
    Dim baseSalary As Double, unitDiff As Double, months As Double
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Dim ws9 As Worksheet: Set ws9 = ThisWorkbook.Sheets("Sheet9")
    Dim r As Variant

    For Each ctrl In Frame1.Controls
        If TypeName(ctrl) = "TextBox" And InStr(ctrl.Name, "CI_") > 0 Then
            empID = Split(ctrl.Name, "_")(1)
            months = val(ctrl.value)
            
            r = Application.match(val(empID), ws1.Range("B:B"), 0)
            If Not IsError(r) Then
                ws9.Range("G18").value = ws1.Cells(r, "H").value
                ws9.Range("F18").value = ws1.Cells(r, "I").value
                Application.CalculateFull
                DoEvents
                
                baseSalary = val(ws9.Range("H18").value)
                unitDiff = val(ws9.Range("S17").value)
                
                If months = 0 Then
                    UpdateTextBoxInFrame empID, "BV", 0
                    UpdateTextBoxInFrame empID, "J", baseSalary
                Else
                    UpdateTextBoxInFrame empID, "BV", unitDiff * months
                    UpdateTextBoxInFrame empID, "J", baseSalary + (unitDiff * months)
                End If
            End If
        End If
    Next ctrl
    MsgBox " „  ÕœÌÀ «·—Ê« »!", vbInformation
End Sub

Private Sub CommandButton1_Click()

Dim Item As Variant, empData As Variant
    Dim r As Range, ws1 As Worksheet
    Dim ID As String
    
    Set ws1 = ThisWorkbook.Sheets(1)
    If TempData.count = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    
    For Each Item In TempData
        empData = Item
        ID = CStr(empData(1))
        Set r = ws1.Columns("B").Find(What:=ID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not r Is Nothing Then
            '  ÕœÌÀ «·»Ì«‰«  ›ﬁÿ œÊ‰ «·„”«” »«·„—Õ·… («· Ì  „   —ﬁÌ Â« »«·“— «·Œ«’)
            UpdateCellIfValueExists ws1, r.row, "J", GetVal(ID, "J")
            UpdateCellIfValueExists ws1, r.row, "H", GetVal(ID, "H")
            '  ÕœÌÀ »ﬁÌ… «·√⁄„œ…...
            UpdateCellIfValueExists ws1, r.row, "BV", GetVal(ID, "BV")
            UpdateCellIfValueExists ws1, r.row, "CI", GetVal(ID, "CI")
            
            Dim dateVal As String
            dateVal = GetValStr(ID, "CJ")
            If dateVal <> "" Then ws1.Cells(r.row, "CJ").value = DateValue(dateVal)
            
            ws1.Calculate
        End If
    Next Item
    
    Application.ScreenUpdating = True
    MsgBox " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ.", vbInformation
    
   Sheets(55).Range("a9:bz8000").ClearContents
Application.ScreenUpdating = False
Sheets(55).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select
Selection.Copy Sheets(55).Range("A1:bz8000")

Application.ScreenUpdating = True
[a8].Select
Sheets(55).Select
Sheets(55).Activate
Application.ScreenUpdating = True
 
    
End Sub



Private Sub CommandButton2_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Long
    ListBox1.Clear
    For i = 9 To ws.Cells(ws.rowS.count, "E").End(xlUp).row
        If val(ws.Cells(i, "BV").value) > 0 Then
            ListBox1.AddItem ws.Cells(i, "B").value
            ListBox1.List(ListBox1.ListCount - 1, 1) = ws.Cells(i, "E").value
        End If
    Next i
End Sub

Private Sub CommandButton3_Click()
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Dim ws9 As Worksheet: Set ws9 = ThisWorkbook.Sheets("Sheet9")
    Dim i As Long, lastRow As Long, empID As String
    Dim baseSalary As Double
    lastRow = ws1.Cells(ws1.rowS.count, "B").End(xlUp).row
    For i = 9 To lastRow
        If val(ws1.Cells(i, "BV").value) > 0 Then
            ws9.Range("G18").value = ws1.Cells(i, "H").value
            ws9.Range("F18").value = ws1.Cells(i, "I").value
            Application.CalculateFull
            baseSalary = val(ws9.Range("H18").value)
            ws1.Cells(i, "BV").value = ""
            ws1.Cells(i, "CI").value = ""
            ws1.Cells(i, "CJ").value = ""
            ws1.Cells(i, "J").value = baseSalary
            empID = ws1.Cells(i, "B").value
            UpdateTextBoxInFrame empID, "BV", ""
            UpdateTextBoxInFrame empID, "CI", ""
            UpdateTextBoxInFrame empID, "CJ", ""
            UpdateTextBoxInFrame empID, "J", Format(baseSalary, "0.00")
        End If
    Next i
    MsgBox " „ «· ’›Ì—.", vbInformation
End Sub

Private Sub CommandButton4_Click()
    Call btnResetAll_Click
End Sub

Private Sub CommandButton5_Click()
    Unload Me
End Sub

Private Sub btnTransfer_Click()
    Dim Item As Variant, r As Range
    Dim tRow As Long, sRow As Long
    If TempData.count = 0 Then Exit Sub
    Application.ScreenUpdating = False
    For Each Item In TempData
        Set r = Sheets(1).Columns("B").Find(Item(1), LookIn:=xlValues, LookAt:=xlWhole)
        If Not r Is Nothing Then
            sRow = r.row
            tRow = Sheets(35).Cells(Sheets(35).rowS.count, "B").End(xlUp).row + 1
            If tRow < 9 Then tRow = 9
            Sheets(1].Range("A" & sRow & ":BU" & sRow).Copy
            Sheets(35).Range("A" & tRow).PasteSpecial xlPasteValues
            With Sheets(35)
                .Cells(tRow, "BW").value = Item(3)
                .Cells(tRow, "BX").value = Item(4)
                .Cells(tRow, "BY").value = GetVal(Item(1), "J")
                .Cells(tRow, "BV").value = GetVal(Item(1), "BV")
            End With
            Sheets(1).Cells(sRow, "J").value = GetVal(Item(1), "J")
            Sheets(1).Cells(sRow, "BV").value = GetVal(Item(1), "BV")
        End If
    Next Item
    Call btnResetAll_Click
    Application.ScreenUpdating = True
    MsgBox " „ «· —ÕÌ·."
    Sheets(34).Activate
End Sub

Private Sub CommandButton6_Click()

Dim Item As Variant, empData As Variant
    Dim r As Range, ws1 As Worksheet, ws35 As Worksheet, wsPast As Worksheet
    Dim tRow As Long, pRow As Range
    Dim ID As String, empName As String

    ' 1.  ‰‘Ìÿ «·‘Ì «  ·÷„«‰  ÕœÌÀ «·„⁄«œ·« 
    
    
    Set ws1 = ThisWorkbook.Sheets(1)
    Set ws35 = ThisWorkbook.Sheets(35)
    Set wsPast = ThisWorkbook.Sheets("past_month55")

    If TempData.count = 0 Then Exit Sub
    Application.ScreenUpdating = False

    For Each Item In TempData
        empData = Item
        ID = CStr(empData(1))
        empName = empData(0)

        ' «·»ÕÀ ⁄‰ «·„ÊŸ› ›Ì «·‘Ì « 
        Set r = ws1.Columns("B").Find(What:=ID, LookIn:=xlValues, LookAt:=xlWhole)
        Set pRow = wsPast.Columns("B").Find(What:=ID, LookIn:=xlValues, LookAt:=xlWhole)

        If Not r Is Nothing And Not pRow Is Nothing Then
            If ws1.Cells(r.row, "E").value = empName Then
                ws1.Calculate
                
                tRow = ws35.Cells(ws35.rowS.count, "B").End(xlUp).row + 1
                If tRow < 9 Then tRow = 9

                ' √-  —ÕÌ· «·»Ì«‰«  «·√”«”Ì… A ≈·Ï E
                ws35.Range("A" & tRow & ":E" & tRow).value = wsPast.Range("A" & pRow.row & ":E" & pRow.row).value

                ' »-  —ÕÌ· «·»Ì«‰«  „‰ F ≈·Ï BU
                wsPast.Range("F" & pRow.row & ":BU" & pRow.row).Copy
                ws35.Range("F" & tRow).PasteSpecial xlPasteValues

                ' Ã-  ⁄»∆… «·√⁄„œ… „‰ BV ≈·Ï CE
                With ws35
                    .Cells(tRow, "BV").value = ws1.Cells(r.row, "BV").value
                    .Cells(tRow, "BW").value = ws1.Cells(r.row, "H").value    '  „  »œÌ·Â« „‰ BU ≈·Ï BW
                    .Cells(tRow, "BX").value = ws1.Cells(r.row, "I").value
                    .Cells(tRow, "BY").value = ws1.Cells(r.row, "J").value
                    .Cells(tRow, "BZ").value = ws1.Cells(r.row, "AE").value
                    .Cells(tRow, "CA").value = ws1.Cells(r.row, "AV").value
                    .Cells(tRow, "CB").value = ws1.Cells(r.row, "AW").value
                    
                    ' œ- «·›—ﬁ (BZ - AV)
                    .Cells(tRow, "CC").value = val(.Cells(tRow, "BZ").value) - val(.Cells(tRow, "CA").value)
                    
                    ' Â‹- «·‰’Ê’ «·À«» …
                    .Cells(tRow, "CD").value = "≈÷«›… «” Õﬁ«ﬁ«  «·⁄·«Ê«  »«À— —Ã⁄Ì"
                    .Cells(tRow, "CE").value = "≈÷«›…"
                End With
            End If
        End If
    Next Item

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Call btnResetAll_Click
    MsgBox " „  —ÕÌ· «·»Ì«‰«  ÊÕ”«» «·›—ﬁ »‰Ã«Õ", vbInformation
    Sheets(35).Activate

End Sub

' 2. œ«·… «·›· —… (TextBox_Search)
Sub LoadEmployeeList(filterText As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Long
    ListBox1.Clear
    For i = 9 To ws.Cells(ws.rowS.count, "E").End(xlUp).row
        If InStr(1, ws.Cells(i, "E").value, filterText, vbTextCompare) > 0 Or _
           InStr(1, ws.Cells(i, "B").value, filterText, vbTextCompare) > 0 Or filterText = "" Then
            ListBox1.AddItem ws.Cells(i, "B").value
            ListBox1.List(ListBox1.ListCount - 1, 1) = ws.Cells(i, "E").value
        End If
    Next i
End Sub

' 3. œ«·… «·Õ–› «·¬„‰… ( „ ≈’·«ÕÂ« »—„ÃÌ«)
Sub btnResetAll_Click()
    Dim i As Integer
    ' ≈Ã»«— «·›Ê—„ ⁄·Ï  ÕœÌÀ «·Õ«·… ﬁ»· «·Õ–›
    Me.CommandButton5.SetFocus
    For i = Me.Frame1.Controls.count - 1 To 0 Step -1
        On Error Resume Next
        Me.Frame1.Controls.Remove Me.Frame1.Controls(i).Name
        On Error GoTo 0
    Next i
    Set TempData = New Collection
    CurrentTop = 33
    Me.Frame1.ScrollHeight = 100
End Sub

' 4. œ«·… Ã·» «·ﬁÌ„ „‰ «·‹ TextBox
Function GetVal(ID As String, col As String) As Double
    Dim c As Control
    For Each c In Frame1.Controls
        If c.Name = col & "_" & ID Then
            GetVal = val(Replace(c.value, ",", ""))
            Exit Function
        End If
    Next c
    GetVal = 0




End Function

Private Sub CommandButton7_Click()

Dim Item As Variant, empData As Variant
    Dim r As Range, ws1 As Worksheet
    Dim ID As String
    Dim oldVal As Variant, newVal As Long
    
    Set ws1 = ThisWorkbook.Sheets(1)
    
    If TempData.count = 0 Then
        MsgBox "·«  ÊÃœ »Ì«‰«  ›Ì «·ﬁ«∆„… «·„ƒﬁ …!", vbExclamation
        Exit Sub
    End If
    
    If MsgBox("Â· √‰  „ √ﬂœ „‰ ≈÷«›… „—Õ·… ÃœÌœ… ·ﬂ· «·„ÊŸ›Ì‰ ›Ì «·ﬁ«∆„…ø", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False
    
    For Each Item In TempData
        empData = Item
        ID = CStr(empData(1))
        
        ' «·»ÕÀ «·œﬁÌﬁ ⁄‰ «·„ÊŸ›
        Set r = ws1.Columns("B").Find(What:=ID, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not r Is Nothing Then
            oldVal = ws1.Cells(r.row, "I").value
            
            '  ÕÊÌ· ÌœÊÌ ¬„‰ ··ﬁÌ„… ﬁ»· «·Ã„⁄
            If IsNumeric(oldVal) Then
                newVal = CLng(oldVal) + 1
                ws1.Cells(r.row, "I").value = newVal
                Debug.Print "ID: " & ID & " | Old: " & oldVal & " | New: " & newVal
            Else
                Debug.Print "ID: " & ID & " | «·ﬁÌ„… ›Ì «·⁄„Êœ I ·Ì”  —ﬁ„«! ﬁÌ„ Â« ÂÌ: " & oldVal
            End If
        Else
            Debug.Print "ID: " & ID & " | ·„ Ì „ «·⁄ÀÊ— ⁄·Ï Â–« «·—ﬁ„ ›Ì «·‘Ì  1."
        End If
    Next Item
    
    ws1.Calculate
    Application.ScreenUpdating = True
    MsgBox " „  «·⁄„·Ì….  √ﬂœ „‰ ‰ «∆Ã «· ÕœÌÀ ›Ì «·⁄„Êœ I.", vbInformation



End Sub

Private Sub TextBox_Search_Change()
LoadEmployeeList TextBox_Search.Text
End Sub



Private Sub TextBox_Search_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
TextBox_Search.Text = ""
    
    ' 2. «” œ⁄«¡ «·œ«·… «· Ì  Õ„· ﬂ· «·„ÊŸ›Ì‰ (»œÊ‰ ›· —…)
    LoadEmployeeList ""
    
    ' 3. ≈·€«¡  √ÀÌ— «·œ»· ﬂ·Ìﬂ «·«› —«÷Ì
    Cancel = True
End Sub

Private Sub UserForm_Initialize()
Zoomer.Bind Me, Me.SpinButton1, Me.az
    CurrentTop = 33
    With ListBox1
        .ColumnCount = 2
        .ColumnWidths = "50;180"
    End With
    LoadEmployeeList ""
End Sub





Private Sub ListBox1_Click()
    If ListBox1.ListIndex = -1 Then Exit Sub
    Dim empID As String: empID = ListBox1.value
    Dim empName As String: empName = ListBox1.List(ListBox1.ListIndex, 1)
    Dim c As Control
    For Each c In Frame1.Controls
        If c.Tag = empID Then Exit Sub
    Next c
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Dim ws9 As Worksheet: Set ws9 = ThisWorkbook.Sheets("Sheet9")
    Dim r As Variant: r = Application.match(val(empID), ws1.Range("B:B"), 0)
    If Not IsError(r) Then
        ws9.Range("G18").value = ws1.Cells(r, "H").value
        ws9.Range("F18").value = ws1.Cells(r, "I").value
        Application.CalculateFull
        AddEntryToFrame empName, empID, ws9.Range("H18").Text, ws9.Range("S17").Text
    End If
End Sub

Sub AddEntryToFrame(Name As String, ID As String, Sal As String, Diff As String)
    If Me.Frame1.Controls.count = 0 Then CurrentTop = 33
    Dim lbl As MSForms.Label
    Set lbl = Frame1.Controls.Add("Forms.Label.1")
    With lbl
        .Caption = "ï " & Name: .Top = CurrentTop: .Left = 10: .Tag = ID
    End With
    Dim rTop As Double: rTop = CurrentTop + 20
    AddTextBox ID, "CJ", Date, 10, 70, rTop
    AddTextBox ID, "CI", "0", 110, 35, rTop
    AddTextBox ID, "BV", Diff, 160, 50, rTop
    AddTextBox ID, "J", Sal, 290, 60, rTop
    Dim empInfo As Variant
    empInfo = Array(Name, ID, Sal, Diff, "", "", "", "", "", "")
    TempData.Add empInfo, CStr(ID)
    CurrentTop = CurrentTop + 55
    Frame1.ScrollHeight = CurrentTop + 20
End Sub

Sub AddTextBox(ID As String, namePart As String, val As String, leftPos As Double, widthPos As Double, topPos As Double)
    Dim txt As MSForms.TextBox
    Set txt = Frame1.Controls.Add("Forms.TextBox.1", namePart & "_" & ID)
    With txt
        .value = val: .Top = topPos: .Left = leftPos: .Width = widthPos: .Tag = ID
    End With
End Sub





Sub UpdateTextBoxInFrame(ID As String, col As String, newVal As Variant)
    Dim c As Control
    For Each c In Frame1.Controls
        If c.Name = col & "_" & ID Then c.value = newVal: Exit Sub
    Next c
End Sub

Sub UpdateCellIfValueExists(ws As Worksheet, row As Long, col As String, val As Double)
    If val <> 0 Then
        ws.Cells(row, col).value = val
    End If
End Sub

' œ«·… ·Ã·» «·‰’Ê’ (·· Ê«—ÌŒ √Ê «·‰’Ê’)
Function GetValStr(ID As String, col As String) As String
    Dim c As Control
    For Each c In Me.Frame1.Controls
        If c.Name = col & "_" & ID Then
            GetValStr = c.value
            Exit Function
        End If
    Next c
    GetValStr = ""
End Function
