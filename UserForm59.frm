VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm59 
   Caption         =   "UserForm59"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9876.001
   OleObjectBlob   =   "UserForm59.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm59"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Label1.Caption = "Ã«—Ú  ‰›Ì– «·ŒÊ«—“„Ì…..."
    Label1.ForeColor = vbBlue
    Me.Repaint

    Generate_EmplMove_Final9

    Label1.Caption = "·ﬁœ  „ »‰«¡ ÂÌﬂ· Õ—ﬂ… «·»Ì«‰«  »ﬂ·  ›«’Ì·Â« ·—Ê« » Â–« «·‘Â—!"
    Label1.ForeColor = vbBlue
End Sub

Sub Generate_EmplMove_Final9()
    Dim wsCur As Worksheet, wsPrev As Worksheet, wsOut As Worksheet
    Dim lastRowCur As Long, lastRowPrev As Long, maxCol As Long, outRow As Long
    Dim empDictCur As Object, empDictPrev As Object
    Dim r As Long, c As Variant
    Dim EmpID As Variant, EmpName As String
    Dim valCur As Double, valPrev As Double, diff As Double
    Dim colName As String, empType As String
    Dim empIDCol As Long, empNameCol As Long, headerRow As Long, dataStartRow As Long
    Dim colTotalAllowance As Long, colTotalDeduction As Long

    empIDCol = 2
    empNameCol = 5
    headerRow = 8
    dataStartRow = 9

    Set wsCur = ThisWorkbook.Sheets("Sheet1")
    Set wsPrev = ThisWorkbook.Sheets("past_month55")
    maxCol = wsCur.Range("CE8").Column

    ' √⁄„œ… «·«” Õﬁ«ﬁ Ê«·«” ﬁÿ«⁄ «·ﬂ·Ì
    colTotalAllowance = Application.Match("„Ã„Ê⁄ «·«” Õﬁ«ﬁ« ", wsCur.Rows(headerRow), 0)
    colTotalDeduction = Application.Match("„Ã„Ê⁄ «·«” ﬁÿ«⁄« ", wsCur.Rows(headerRow), 0)

    ' ≈‰‘«¡ √Ê  ›—Ì€ ‘Ì  «·‰ «∆Ã
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets("emplmove")
    If wsOut Is Nothing Then
        Set wsOut = ThisWorkbook.Sheets.Add
        wsOut.Name = "emplmove"
    End If
    On Error GoTo 0
    wsOut.Cells.ClearContents

    ' ⁄‰«ÊÌ‰ «·‰ «∆Ã
    wsOut.Range("A1:I1").value = Array( _
        "«·—ﬁ„ «·ÊŸÌ›Ì", "«·«”„", _
        "≈÷«›… «” Õﬁ«ﬁ", "Õ–› «” Õﬁ«ﬁ", _
        "≈÷«›… «” ﬁÿ«⁄", "Õ–› «” ﬁÿ«⁄", _
        "‰Ê⁄ «·Õ—ﬂ…", "«·Õ«·… «·ÊŸÌ›Ì…", "«·⁄„Êœ «·„ √À—")
    outRow = 2

    lastRowCur = wsCur.Cells(wsCur.Rows.count, empIDCol).End(xlUp).row
    lastRowPrev = wsPrev.Cells(wsPrev.Rows.count, empIDCol).End(xlUp).row

    Set empDictCur = CreateObject("Scripting.Dictionary")
    Set empDictPrev = CreateObject("Scripting.Dictionary")

    ' »‰«¡ ﬁ«„Ê” «·„ÊŸ›Ì‰
    For r = dataStartRow To lastRowCur
        If wsCur.Cells(r, empIDCol).value <> "" Then empDictCur(CStr(wsCur.Cells(r, empIDCol).value)) = r
    Next r
    For r = dataStartRow To lastRowPrev
        If wsPrev.Cells(r, empIDCol).value <> "" Then empDictPrev(CStr(wsPrev.Cells(r, empIDCol).value)) = r
    Next r

    ' ======== «·„ÊŸ›Ê‰ «·„” „—Ê‰ ========
    For Each EmpID In empDictCur.Keys
        r = empDictCur(EmpID)
        EmpName = wsCur.Cells(r, empNameCol).value
        If empDictPrev.exists(EmpID) Then
            empType = "„” „—"
            For c = 1 To maxCol
                colName = wsCur.Cells(headerRow, c).value
                If colName = "" Then GoTo NextCol

                ' «” »⁄«œ «·√⁄„œ… €Ì— «·„—€Ê»…
                If colName = "⁄œœ «·√ÿ›«·" Or colName = "„Ã„Ê⁄ «·„Œ’’« " _
                   Or colName = "„Ã„Ê⁄ «·«” Õﬁ«ﬁ« " Or colName Like "*’«›Ì*" _
                   Or colName = "‰”»… „Â‰Ì…" Or colName = "„Ã„Ê⁄ «·«” ﬁÿ«⁄« " _
                   Or colName = "«· ”·”·" Or colName = "√Ì«„ «·€Ì«»" Or colName = "«Ì«„ «·œÊ«„" _
                   Or colName = "‰”»… «·Â‰œ”Ì…" Then GoTo NextCol

                valCur = NzSafe(wsCur.Cells(r, c).value)
                valPrev = NzSafe(wsPrev.Cells(empDictPrev(EmpID), c).value)
                diff = valCur - valPrev
                If diff <> 0 Then
                    If Not colName Like "*«” ﬁÿ«⁄*" Then
                        WriteMove wsOut, outRow, EmpID, EmpName, diff, "Allowance", empType, colName
                    Else
                        wsOut.Cells(outRow, 1).value = EmpID
                        wsOut.Cells(outRow, 2).value = EmpName
                        If diff > 0 Then
                            wsOut.Cells(outRow, 5).value = CLng(diff)
                        Else
                            wsOut.Cells(outRow, 6).value = CLng(-Abs(diff))
                        End If
                        wsOut.Cells(outRow, 7).value = empType
                        wsOut.Cells(outRow, 8).value = empType
                        wsOut.Cells(outRow, 9).value = colName
                        outRow = outRow + 1
                    End If
                End If
NextCol:
            Next c
        End If
    Next EmpID

    ' ======== «·„ÊŸ›Ê‰ «·Ãœœ ========
    For Each EmpID In empDictCur.Keys
        If Not empDictPrev.exists(EmpID) Then
            r = empDictCur(EmpID)
            EmpName = wsCur.Cells(r, empNameCol).value
            wsOut.Cells(outRow, 1).value = EmpID
            wsOut.Cells(outRow, 2).value = EmpName
            wsOut.Cells(outRow, 3).value = NzSafe(wsCur.Cells(r, colTotalAllowance).value)
            wsOut.Cells(outRow, 5).value = CLng(NzSafe(wsCur.Cells(r, colTotalDeduction).value))
            wsOut.Cells(outRow, 7).value = "«÷«›… „ÊŸ›"
            wsOut.Cells(outRow, 8).value = "«÷«›… „ÊŸ›"
            wsOut.Cells(outRow, 9).value = "«·„ÊŸ› „÷«› «·Ï «·‰Ÿ«„"
            outRow = outRow + 1
        End If
    Next EmpID

    ' ======== «·„ÊŸ›Ê‰ «·„Õ–Ê›Ê‰ / «·„‰ﬁÊ·Ê‰ ========
    For Each EmpID In empDictPrev.Keys
        If Not empDictCur.exists(EmpID) Then
            r = empDictPrev(EmpID)
            EmpName = wsPrev.Cells(r, empNameCol).value
            wsOut.Cells(outRow, 1).value = EmpID
            wsOut.Cells(outRow, 2).value = EmpName
            wsOut.Cells(outRow, 4).value = -Abs(NzSafe(wsPrev.Cells(r, colTotalAllowance).value))
            wsOut.Cells(outRow, 6).value = -Abs(CLng(NzSafe(wsPrev.Cells(r, colTotalDeduction).value)))
            wsOut.Cells(outRow, 7).value = "«·„ÊŸ› „Õ–Ê›"
            wsOut.Cells(outRow, 8).value = "«·„ÊŸ› „Õ–Ê›"
            wsOut.Cells(outRow, 9).value = "«·„ÊŸ› „‰ﬁÊ· √Ê  „ ≈Õ«· Â ≈·Ï «· ﬁ«⁄œ"
            outRow = outRow + 1
        End If
    Next EmpID

    ' ======== Õ–› «·”Ã·«  «·€Ì— „—€Ê»… ========
    Dim lastRow As Long, i As Long
    lastRow = wsOut.Cells(wsOut.Rows.count, 1).End(xlUp).row
    For i = lastRow To 2 Step -1
        Dim cellVal As String
        cellVal = wsOut.Cells(i, 9).value
        If cellVal Like "*⁄œœ «·«ÿ›«·*" _
           Or cellVal Like "*«Ì«„ «·œÊ«„*" _
           Or cellVal Like "*√Ì«„ «·€Ì«»*" _
           Or cellVal Like "*‰”»… «·Â‰œ”Ì…*" _
           Or cellVal Like "*‰”»… „Â‰Ì…*" _
           Or cellVal Like "*‰”»… ŒÿÊ—…*" _
           Or cellVal Like "*«” ﬁÿ«⁄ «·—⁄«Ì… «·«Ã „«⁄Ì…*" _
           Or cellVal Like "*„Ã„Ê⁄ «·«” Õﬁ«ﬁ« *" _
           Or cellVal Like "*„Ã„Ê⁄ «·«” ﬁÿ«⁄« *" Then
           
 wsOut.Rows(i).Delete
          
           
        End If
    Next i

    ' ========  ÕœÌÀ ‰Ê⁄ «·Õ—ﬂ… ··„ €Ì—«  ========
    lastRow = wsOut.Cells(wsOut.Rows.count, 1).End(xlUp).row
    For i = 2 To lastRow
        ' ·« ‰⁄œ· «·„ÊŸ›Ì‰ «·„÷«›Ì‰ √Ê «·„Õ–Ê›Ì‰
        If wsOut.Cells(i, 8).value <> "≈÷«›… „ÊŸ›" And wsOut.Cells(i, 8).value <> "«·„ÊŸ› „Õ–Ê›" Then
            ' ≈÷«›… «” Õﬁ«ﬁ
            If wsOut.Cells(i, 3).value > 0 Then wsOut.Cells(i, 7).value = "„ €Ì—«  «÷«›…"
            ' Õ–› «” Õﬁ«ﬁ
            If wsOut.Cells(i, 4).value < 0 Then wsOut.Cells(i, 7).value = "„ €Ì—«  Õ–›"
            ' ≈÷«›… «” ﬁÿ«⁄
            If wsOut.Cells(i, 5).value > 0 Then wsOut.Cells(i, 7).value = "Õ—ﬂ… «” ﬁÿ«⁄ „÷«›…"
            ' Õ–› «” ﬁÿ«⁄
            If wsOut.Cells(i, 6).value < 0 Then wsOut.Cells(i, 7).value = "Õ–› «” ﬁÿ«⁄"
        End If
    Next i
    ' ========  ÕœÌÀ ‰Ê⁄ «·Õ—ﬂ… ··„ €Ì—«  ========
    Dim i3, lastRow3 As Integer
    lastRow3 = wsOut.Cells(wsOut.Rows.count, 1).End(xlUp).row
    For i3 = 2 To lastRow3
        ' ·« ‰⁄œ· «·„ÊŸ›Ì‰ «·„÷«›Ì‰ √Ê «·„Õ–Ê›Ì‰
        If wsOut.Cells(i3, 8).value = "«÷«›… „ÊŸ›" Or wsOut.Cells(i3, 8).value = "«·„ÊŸ› „Õ–Ê›" Then
            ' ≈÷«›… «” Õﬁ«ﬁ
           wsOut.Cells(i3, 7).value = wsOut.Cells(i3, 8).value
            ' Õ–› «” Õﬁ«ﬁ
            wsOut.Cells(i3, 7).value = wsOut.Cells(i3, 8).value
            ' ≈÷«›… «” ﬁÿ«⁄
                End If
    Next i3

    ' ======== «·„Ã«„Ì⁄ «·‰Â«∆Ì… ========
   ' Dim sumVal As Double
  '  Dim colsToSum As Variant
   ' wsOut.Cells(lastRow + 1, 1).Value = "«·„Ã«„Ì⁄ «·‰Â«∆Ì…"
   ' colsToSum = Array(3, 4, 5, 6)
   ' For Each c In colsToSum
   '     sumVal = Application.Sum(wsOut.Range(wsOut.Cells(2, c), wsOut.Cells(lastRow, c)))
    '    wsOut.Cells(lastRow + 1, c).Value = sumVal
   ' Next c

End Sub

Function NzSafe(v As Variant) As Double
    If IsNumeric(v) Then NzSafe = CDbl(v) Else NzSafe = 0
End Function

Sub WriteMove(ws As Worksheet, ByRef r As Long, EmpID As Variant, EmpName As String, _
              diff As Double, moveType As String, empType As String, colName As String)

    ws.Cells(r, 1).value = EmpID
    ws.Cells(r, 2).value = EmpName

    If moveType = "Allowance" Then
        If diff > 0 Then
            ws.Cells(r, 3).value = diff
            ws.Cells(r, 7).value = IIf(empType = "≈÷«›… „ÊŸ›", "≈÷«›… „ÊŸ›", "„” „—")
        Else
            ws.Cells(r, 4).value = -Abs(diff)
            ws.Cells(r, 7).value = "„” „—"
        End If
    End If

    If moveType = "Deduction" Then
        If diff > 0 Then
            ws.Cells(r, 5).value = CLng(diff)
            ws.Cells(r, 7).value = IIf(empType = "≈÷«›… „ÊŸ›", "≈÷«›… „ÊŸ›", "„” „—")
        Else
            ws.Cells(r, 6).value = CLng(-Abs(diff))
            ws.Cells(r, 7).value = "„” „—"
        End If
    End If

    ' «·⁄„Êœ «·„ √À—
    If colName = "" Or colName = "‰”»… „Â‰Ì…" Or colName = "„Ã„Ê⁄ «·«” Õﬁ«ﬁ« " _
       Or colName = "„Ã„Ê⁄ «·«” ﬁÿ«⁄« " Then
        ws.Cells(r, 9).value = "€Ì— „Õœœ"
    Else
        ws.Cells(r, 9).value = CStr(colName)
    End If

    r = r + 1
End Sub

Private Sub CommandButton2_Click()
Unload Me
UserForm60.Show
End Sub

Private Sub UserForm_Click()

End Sub
