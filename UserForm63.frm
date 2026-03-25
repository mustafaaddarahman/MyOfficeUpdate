VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm63 
   Caption         =   "UserForm63"
   ClientHeight    =   9255.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9996.001
   OleObjectBlob   =   "UserForm63.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm63"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim ws As Worksheet
Dim lastRow As Long
Dim LastCol As Long

'========================
' Initialize UserForm
'========================
Private Sub UserForm_Initialize()
On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    LastCol = ws.Cells(8, ws.Columns.count).End(xlToLeft).Column
    
    LoadEmployees ""
End Sub

'========================
' Load Employees into ListBox
'========================
Private Sub LoadEmployees(filterText As String)
On Error Resume Next
    Dim r As Long
    Dim EmpName As String
    
    ListBox1.Clear
    ListBox1.ColumnCount = 2
    ListBox1.ColumnWidths = "200 pt;0 pt"
    
    For r = 9 To lastRow
        EmpName = Trim(ws.Cells(r, "E").value)
        If EmpName <> "" Then
            If filterText = "" Or InStr(LCase(EmpName), LCase(filterText)) > 0 Then
                ListBox1.AddItem EmpName
                ListBox1.List(ListBox1.ListCount - 1, 1) = r
            End If
        End If
    Next r
End Sub

Private Sub TextBox1_Change()
On Error Resume Next
    LoadEmployees TextBox1.Text
End Sub

'========================
' CommandButton Click
'========================
Private Sub CommandButton1_Click()
On Error Resume Next
    Dim selectedRow As Long
    Dim question As String
    Dim ans As String
    Dim col As Long
    Dim colName As String
    Dim value As String
    Dim words() As String
    Dim i As Long
    Dim matchScore() As Double
    Dim maxScore As Double
    Dim bestCols() As Long
    Dim tolerance As Double
    
    tolerance = 0.6
    
    If ListBox1.ListIndex = -1 Then
        MsgBox "«Œ — „ÊŸ› √Ê·«"
        Exit Sub
    End If
    
    selectedRow = CLng(ListBox1.List(ListBox1.ListIndex, 1))
    question = Trim(LCase(TextBox2.Text))
    
    If question = "" Then
        MsgBox "«ﬂ » «·”ƒ«·"
        Exit Sub
    End If
    
    words = Split(question, " ")
    ReDim matchScore(1 To LastCol)
    
    '========================
    '  Õ·Ì· «·√⁄„œ…
    '========================
    For col = 1 To LastCol
        colName = LCase(ws.Cells(8, col).value)
        value = LCase(ws.Cells(selectedRow, col).Text)
        matchScore(col) = 0
        
        For i = LBound(words) To UBound(words)
            matchScore(col) = matchScore(col) + Similarity(words(i), colName) * 2
            matchScore(col) = matchScore(col) + Similarity(words(i), value)
        Next i
    Next col
    
    '========================
    ' ≈ÌÃ«œ √›÷· «·√⁄„œ…
    '========================
    maxScore = 0
    For col = 1 To LastCol
        If matchScore(col) > maxScore Then maxScore = matchScore(col)
    Next col
    
    ReDim bestCols(0)
    For col = 1 To LastCol
        If matchScore(col) >= maxScore * tolerance Then
            ReDim Preserve bestCols(UBound(bestCols) + 1)
            bestCols(UBound(bestCols)) = col
        End If
    Next col
    
    '========================
    ' »‰«¡ «·≈Ã«»…
    '========================
    ans = ""
    If UBound(bestCols) = 1 Then
        ans = "≈Ã«»… œﬁÌﬁ… ··„ÊŸ› «·„Œ «—:" & vbCrLf
        col = bestCols(0)
        ans = ans & "ï " & ws.Cells(8, col).value & ": " & ws.Cells(selectedRow, col).Text
    ElseIf UBound(bestCols) > 1 Then
        ans = "ﬁœ ÌﬂÊ‰ Â‰«ﬂ √ﬂÀ— „‰ «Õ „«· ·Â–« «·”ƒ«·:" & vbCrLf
        Dim j As Long
        For j = 1 To UBound(bestCols)
            col = bestCols(j)
            ans = ans & "ï " & ws.Cells(8, col).value & ": " & ws.Cells(selectedRow, col).Text & vbCrLf
        Next j
    Else
        ans = "«·”ƒ«· €«„÷° ≈·Ìﬂ ﬂ· »Ì«‰«  «·„ÊŸ›:" & vbCrLf
        For col = 1 To LastCol
            ans = ans & "ï " & ws.Cells(8, col).value & ": " & ws.Cells(selectedRow, col).Text & vbCrLf
        Next col
    End If
    
    TextBox3.Text = ans
End Sub

'========================
' œ«·… Similarity œ«Œ· ‰›” «·‹ UserForm
'========================
Private Function Similarity(ByVal s1 As String, ByVal s2 As String) As Double
On Error Resume Next
    Dim maxLen As Long
    maxLen = WorksheetFunction.Max(Len(s1), Len(s2))
    If maxLen = 0 Then
        Similarity = 1
    Else
        Similarity = 1 - (Levenshtein(s1, s2) / maxLen)
    End If
End Function

'========================
' œ«·… Levenshtein
'========================
Private Function Levenshtein(ByVal s As String, ByVal t As String) As Long
On Error Resume Next
    Dim i As Long, j As Long
    Dim sLen As Long, tLen As Long
    Dim cost As Long
    Dim d() As Long
    Dim min1 As Long, min2 As Long, min3 As Long
    
    sLen = Len(s)
    tLen = Len(t)
    
    ReDim d(0 To sLen, 0 To tLen)
    
    For i = 0 To sLen
        d(i, 0) = i
    Next i
    For j = 0 To tLen
        d(0, j) = j
    Next j
    
    For i = 1 To sLen
        For j = 1 To tLen
            If Mid(s, i, 1) = Mid(t, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            min1 = d(i - 1, j) + 1
            min2 = d(i, j - 1) + 1
            min3 = d(i - 1, j - 1) + cost
            d(i, j) = WorksheetFunction.Min(min1, min2, min3)
        Next j
    Next i
    
    Levenshtein = d(sLen, tLen)
End Function

