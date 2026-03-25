VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm19 
   Caption         =   "UserForm19"
   ClientHeight    =   1725
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7500
   OleObjectBlob   =   "UserForm19.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
On Error GoTo emad

Dim fd As FileDialog
    Dim filePath As String

    ' ≈‰‘«¡ „—»⁄ «Œ Ì«— «·„·›
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .title = "«Œ Ì«— „·› Excel"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"

        If .Show <> -1 Then Exit Sub  ' ≈–« ÷€ÿ «·„” Œœ„ ≈·€«¡
        filePath = .SelectedItems(1)
    End With

    ' › Õ «·„·› «·„Õœœ
    Workbooks.Open filePath
emad:

End Sub

Private Sub CommandButton2_Click()
On Error Resume Next
Windows("import_data.xlsx").Activate
Sheets(1).Select
ActiveSheet.Range(RefEdit1.value).Select
Selection.Copy
ThisWorkbook.Activate
Sheets(1).Select
Dim lastR As Integer
Dim i As Integer
'lastr = Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
lastR = Application.WorksheetFunction.CountA(Sheets(1).Range("b9:b1000")) + 9
'For i = 9 To lastr
Range("a" & lastR).Select
Selection.PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False
Unload Me
MsgBox "ﬁœ  „  ‰ﬁ· «·»Ì«‰«  »‰Ã«Õ ⁄»— «·„·› «·Œ«—ÃÌ ÊÌ„ﬂ‰ﬂ„ «·„ «»⁄Â ›Ì  ÕœÌÀÂ«", vbDefaultButton2, "—”«·…  ÊÃÌÂ"
'=================== €·ﬁ „·› «·„’œ—

 Application.Visible = True
    ' ·≈Œ›«¡ «·‘—Ìÿ (Ribbon)  „«„«
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",true)"





 Dim fd As FileDialog
    Dim filePath As String
    Dim srcWB As Workbook
    Dim dstWS As Worksheet
    Dim srcRange As Range
    Dim NextRow As Long
    Dim wb As Workbook

    ' --- 1. «Œ Ì«— «·„·› «·„’œ— ---
  '  Set fd = Application.FileDialog(msoFileDialogFilePicker)
   ' With fd
     '   .title = "«Œ — „·› «·„‘—Ê⁄ «·„’œ—"
     '   .AllowMultiSelect = False
     '   .Filters.Clear
     '   .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb"
     '   If .Show <> -1 Then Exit Sub
     '   filePath = .SelectedItems(1)
  '  End With

    ' --- 2. › Õ «·„·› «·„’œ— Ê Œ“Ì‰ «·„—Ã⁄ ---
    Set srcWB = Workbooks.Open(filePath)

    ' --- 3. ‰”Œ «·»Ì«‰«  ≈·Ï „·› «·‰Ÿ«„ ---
    Set srcRange = srcWB.Sheets(1).Rows(2) ' Ì„ﬂ‰  €ÌÌ—Â Õ”» «·Õ«Ã…
    Set dstWS = Workbooks("Financial and administrative system.xlsm").Sheets("Clipboard")
    NextRow = dstWS.Cells(dstWS.Rows.count, 1).End(xlUp).row + 1

    srcRange.Copy
    dstWS.Rows(NextRow).PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    ' --- 4. ≈€·«ﬁ «·„·› «·„’œ— ---
    If Not srcWB Is Nothing Then
        If srcWB.Name <> "Financial and administrative system.xlsm" Then
            Application.DisplayAlerts = False
            srcWB.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    End If

    ' --- 5. ≈€·«ﬁ √Ì „·›«  »Ì∆… «· ÿ»Ì⁄ «·„› ÊÕ… «·√Œ—Ï ---
    For Each wb In Workbooks
        If wb.Name <> "Financial and administrative system.xlsm" Then
            Application.DisplayAlerts = False
            wb.Close SaveChanges:=False
            Application.DisplayAlerts = True
        End If
    Next wb

    ' --- 6. ≈Œ›«¡ «· ÿ»Ìﬁ Ê≈Œ›«¡ «·—Ì»Ê‰ («·Ê«ÃÂ…) ---
    Application.Visible = False
    ' ·≈Œ›«¡ «·‘—Ìÿ (Ribbon)  „«„«
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"

    MsgBox " „ «·‰”Œ Ê≈€·«ﬁ «·„·›«  «·„ƒﬁ …. «·‰Ÿ«„ «·¬‰ ›Ì Ê÷⁄ «·Ê«ÃÂ… «·„Œ›Ì….", vbInformation


Dim ws100 As Worksheet
    Dim lastRow100 As Long
    Dim i100 As Long
    
    ' «Œ Ì«— «·‘Ì  «·Õ«·Ì
    Set ws100 = ActiveSheet
    
    ' ≈ÌÃ«œ ¬Œ— ’› ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ›Ì «·⁄„Êœ B
    lastRow100 = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    
    '  ⁄»∆… «·⁄„Êœ A «» œ«¡ „‰ «·’› 9 »«” Œœ«„ ›ﬂ—… COUNTA
    For i = 9 To lastRow100
        ws.Cells(i100, "A").value = Application.WorksheetFunction.CountA(ws.Range("B9:B" & i100))
    Next i











End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub RefEdit1_BeforeDragOver(Cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal X As stdole.OLE_XPOS_CONTAINER, ByVal Y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Click()

End Sub
