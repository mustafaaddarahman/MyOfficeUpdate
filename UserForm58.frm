VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm58 
   Caption         =   "UserForm58"
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   15288
   OleObjectBlob   =   "UserForm58.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

 Dim ws As Worksheet
    Dim cell As Range
    Dim errType As String
    Dim details As String
    Dim lastRow As Long
    Dim rowStart As Long
    Dim colName As String
    Dim formulaText As String
    Dim col As Long

    ' „”Õ ListBox ﬁ»· «·»œ¡
    ListBox1.Clear
    ListBox1.ColumnCount = 5
    ListBox1.ColumnWidths = "100 pt;100 pt;150 pt;120 pt;200 pt"
    ListBox1.Font.Size = 12
    ListBox1.Font.Bold = True
    ListBox1.Font.Name = "Arial"

    rowStart = 8 ' «·’› «·–Ì Ì»œ√ „‰Â «·»Ì«‰« 
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' ¬Œ— ’› „” Œœ„ ›Ì «·⁄„Êœ 2
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row

    ' «·„—Ê— ⁄·Ï ﬂ· «·Œ·«Ì« „‰ «·’› 8 Õ Ï ¬Œ— ’› „” Œœ„
    For Each cell In ws.Range(ws.Cells(rowStart, 1), ws.Cells(lastRow, ws.UsedRange.Columns.count))

        errType = ""
        details = ""
        colName = ws.Cells(8, cell.Column).value ' «”„ «·⁄„Êœ „‰ «·’› 8

        ' ›Õ’ √Œÿ«¡ Excel
        If IsError(cell.value) Then
            errType = "Œÿ√ ›Ì Excel"
            details = cell.Text

        ' ›Õ’ «·ﬁÌ„ €Ì— «·’ÕÌÕ… („À«·: ‰’ »œ· —ﬁ„)
        ElseIf cell.value <> "" Then
            ' «·⁄„Êœ 2 ÌÃ» √‰ ÌﬂÊ‰ —ﬁ„
            If cell.Column = 2 And Not IsNumeric(cell.value) Then
                errType = "ﬁÌ„… €Ì— ’ÕÌÕ…"
                details = "«·„› —÷ √‰ ÌﬂÊ‰ —ﬁ„«"
            End If
        End If

        ' ≈–«  „ «·⁄ÀÊ— ⁄·Ï Œÿ√
        If errType <> "" Then

            ' ≈÷«›… ’› ÃœÌœ ›Ì ListBox
            ListBox1.AddItem ws.Name                                 ' «”„ «·‘Ì 
            ListBox1.List(ListBox1.ListCount - 1, 1) = ws.Cells(cell.row, 2).value ' «·—ﬁ„ «·ÊŸÌ›Ì
            ListBox1.List(ListBox1.ListCount - 1, 2) = ws.Cells(cell.row, 5).value ' «·«”„
            ListBox1.List(ListBox1.ListCount - 1, 3) = colName      ' «”„ «·⁄„Êœ
            ListBox1.List(ListBox1.ListCount - 1, 4) = errType & " - " & details   ' ‰Ê⁄/ ›«’Ì· «·Œÿ√

            '  ·ÊÌ‰ «·’› »«·ﬂ«„· ›Ì «·‘Ì 
            ws.Rows(cell.row).Interior.Color = RGB(255, 200, 200)

            ' ’Ì«‰… «·Œ·Ì…
            If cell.HasFormula Then
                formulaText = cell.Formula
                col = cell.Column
                '  ⁄„Ì„ «·„⁄«œ·… ⁄·Ï »«ﬁÌ «·’›Ê› ›Ì Â–« «·⁄„Êœ
                ws.Range(ws.Cells(rowStart, col), ws.Cells(lastRow, col)).Formula = formulaText
            Else
                ' ≈–« ·„  ﬂ‰ „⁄«œ·…° ÷⁄ ’›—
                cell.value = 0
            End If

        End If

    Next cell

    MsgBox " „ «·«‰ Â«¡ „‰ «·›Õ’°  ’ÕÌÕ «·√Œÿ«¡° Ê«·’›Ê› «· Ì  Õ ÊÌ √Œÿ«¡ „·Ê‰… »«··Ê‰ «·√Õ„—", vbInformation


    
  
End Sub

Private Sub CommandButton2_Click()
Unload Me
UserForm9.Show
End Sub
