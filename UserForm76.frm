VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm76 
   Caption         =   "UserForm76"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5028
   OleObjectBlob   =   "UserForm76.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm76"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CheckBox1_Click()
If CheckBox1.value = True Then CheckBox2.value = False
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.value = True Then CheckBox1.value = False
End Sub

Private Sub CommandButton1_Click()
Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim percentage As Double
    Dim dataRange As Variant ' „’›Ê›… ·”Õ» «·»Ì«‰« 
    Dim resultsArray() As Variant ' „’›Ê›… · Œ“Ì‰ «·‰ «∆Ã
    Dim sourceColIndex As Integer
    
    ' 1. «· Õﬁﬁ „‰ «· Ìﬂ”  »Êﬂ”
    If Me.TextBox1.value = "" Then Exit Sub
    If Not IsNumeric(Me.TextBox1.value) Then Exit Sub
    
    ' 2.  ÕœÌœ —ﬁ„ «·⁄„Êœ «·„’œ—
    If Me.CheckBox1.value = True Then
        sourceColIndex = 10 ' ⁄„Êœ J ÂÊ «·⁄„Êœ —ﬁ„ 10
    ElseIf Me.CheckBox2.value = True Then
        sourceColIndex = 49 ' ⁄„Êœ AW ÂÊ «·⁄„Êœ —ﬁ„ 49
    Else
        MsgBox "Ì—ÃÏ  ÕœÌœ «·‰Ê⁄ √Ê·«", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    percentage = CDbl(Me.TextBox1.value) / 100
    lastRow = ws.Cells(ws.Rows.count, "J").End(xlUp).row
    
    If lastRow < 9 Then Exit Sub
    
    ' 3. ”Õ» «·»Ì«‰«  ≈·Ï „’›Ê›… („‰ «·’› 9 ≈·Ï ¬Œ— ’› Ê„‰ «·⁄„Êœ J ≈·Ï AW)
    ' ”Õ» «·‰ÿ«ﬁ »«·ﬂ«„· „—… Ê«Õœ… · Ê›Ì— «·Êﬁ 
    dataRange = ws.Range("A1:AW" & lastRow).value
    
    '  ÃÂÌ“ „’›Ê›… «·‰ «∆Ã (»‰›” ÕÃ„ ⁄œœ «·’›Ê›)
    ReDim resultsArray(9 To lastRow, 1 To 1)
    
    ' 4. «·„⁄«·Ã… œ«Œ· «·„’›Ê›… (›Ì «·–«ﬂ—… - ”—⁄… ›«∆ﬁ…)
    For i = 9 To lastRow
        Dim salaryValue As Variant
        salaryValue = dataRange(i, sourceColIndex)
        
        If IsNumeric(salaryValue) And salaryValue <> "" Then
            ' «·„⁄«œ·… «·„ÿ·Ê»…: («·—« » * «·‰”»…) * 100
            resultsArray(i, 1) = (salaryValue * percentage) * 100
        Else
            resultsArray(i, 1) = 0
        End If
    Next i
    
    ' 5.  ›—Ì€ «·‰ «∆Ã „‰ «·–«ﬂ—… ≈·Ï «·‘Ì  »÷—»… Ê«Õœ…
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual ' ≈Ìﬁ«› «·Õ”«» «· ·ﬁ«∆Ì ·“Ì«œ… «·”—⁄…
    
    ws.Range("AH9:AH" & lastRow).value = resultsArray
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    
    MsgBox " „  «·„⁄«·Ã… «·»Ì«‰«  Õ”» ‰”»… «·«” ﬁÿ«⁄ «‰ ÷— ·ÕÌ‰ Ì „ «⁄«œ… «Õ ”«» «·»Ì«‰«  «·ﬂ·Ì !", vbInformation
Call count_all
 MsgBox " „  ⁄„·Ì… «·«Õ ”«» !", vbInformation
End Sub
