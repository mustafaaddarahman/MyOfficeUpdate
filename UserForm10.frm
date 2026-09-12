VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm10 
   Caption         =   "UserForm10"
   ClientHeight    =   8808.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13452
   OleObjectBlob   =   "UserForm10.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME As Long = &H40000
Const MIN_BOX As Long = &H20000
Const MAX_BOX As Long = &H10000
Private Declare PtrSafe Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As Long

Public Sub AddToForm(ByVal Box_Type As Long)
Dim BisMask As Long
Dim Window_Handle As Long
Dim WindowStyle As Long
Dim Ret As Long
If Box_Type = MIN_BOX Or Box_Type = MAX_BOX Then
    Window_Handle = GetForegroundWindow()
    WindowStyle = GetWindowLong(Window_Handle, GWL_STYLE)
    BisMask = WindowStyle Or Box_Type
    Ret = SetWindowLong(Window_Handle, GWL_STYLE, BisMask)
    Ret = DrawMenuBar(Window_Handle)
End If
End Sub


Private Sub ComboBox1_Change()
On Error GoTo emad

 If ComboBox1.ListIndex = 0 Then
Label7.Caption = "«·ÃœÊ· «·„«·Ì «·«”«”Ì"
ListBox1.RowSource = ""
End If

If ComboBox1.ListIndex = 1 Then
Label7.Caption = "«·„’›Ê›… «·÷—Ì»ÌÂ"
ListBox1.RowSource = ""


End If

If ComboBox1.ListIndex = 3 Then
Label7.Caption = "ÃœÊ· »Ì«‰«  «Õ Ì«ÿÌ"
ListBox1.RowSource = ""
End If
If ComboBox1.ListIndex = 10 Then
Label7.Caption = "»Ì«‰«  ·Ì”  ··⁄—÷"
ListBox1.RowSource = ""
Sheets(11).Visible = False
End If
If ComboBox1.ListIndex = 11 Then
Label7.Caption = "»Ì«‰«  ·Ì”  ··⁄—÷"
ListBox1.RowSource = ""
Sheets(12).Visible = False
End If


Dim Y As String
 
 Y = ComboBox1.value
 Sheets(Y).Activate
 X = ComboBox1.ListIndex + 1
 ListBox1.RowSource = "table" & X
  ListBox1.ColumnHeads = True
 ListBox1.ColumnCount = 63
 [a2].Select

'================================



 Select Case ComboBox1.ListIndex
 Case 0
 Label7.Caption = "«·ÃœÊ· «·„«·Ì «·«”«”Ì"


Case 1
 Label7.Caption = "«·„’›Ê›Â «·÷—Ì»ÌÂ"
 
Case 2
 Label7.Caption = "»Ì«‰«  «·„Õ–Ê›Ì‰ Ê«·„—Õ·Ì‰"
Case 3
 Label7.Caption = "ÃœÊ· »Ì«‰«  «Õ Ì«ÿÌÂ"
 Sheets(4).Visible = False
 
Case 4
 Label7.Caption = "»Ì«‰«  «·⁄‰«ÊÌ‰ «·ÊŸÌ›ÌÂ"
 
Case 5
 Label7.Caption = "»Ì«‰«  «·⁄·«Ê«  Ê«· —›Ì⁄« "
Case 6
 Label7.Caption = "»Ì«‰«  «·—ﬁ„ «·ÊŸÌ›Ì"
Case 7
 Label7.Caption = "»Ì«‰«  «· —ÕÌ· ··⁄·«Ê«  Ê«· —›Ì⁄« "
Case 8
 Label7.Caption = "„’›Ê›… «·—Ê« »"
Case 9
 Label7.Caption = "„’›Ê›… «·Ê’› «·ÊŸÌ›Ì "
Case 10
 Label7.Caption = "»Ì«‰«  ·Ì”  ··⁄—÷"
Sheets(11).Visible = False
 
Case 11
 Label7.Caption = "»Ì«‰«  ·Ì”  ··⁄—÷"
Sheets(12).Visible = False
 
Case 12
 Label7.Caption = "⁄—÷  ›«’Ì· «·»Ì«‰«  «·„«·ÌÂ"

Case 13
 Label7.Caption = "»Ì«‰«  „ƒﬁ Â ·€—÷ «·ÿ»«⁄Â"
Sheets(13).Visible = True
Case 17
 Label7.Caption = "Œ·«’«  «·„»«·€ Õ”» «·œ—Ã… Ê«·«⁄œ«œ"
Sheets(18).Visible = True

End Select

ListBox1.ColumnHeads = True






emad:
End Sub






Private Sub CommandButton15_Click()
Unload Me
UserForm11.Show
End Sub

Private Sub CommandButton16_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton17_Click()
On Error GoTo emad

Sheets(15).Range("a8:L30000").ClearContents
MsgBox " „ Õ–› Ã„Ì⁄ »Ì«‰«  «‘—ÿ… «·—« » «·ﬁœÌ„Â · ÂÌ∆… «‘—ÿ… «·—« » ··‘Â— «·Õ«·Ì Êﬁœ Ì” €—ﬁ «‰‘«¡ «·«‘—ÿÂ «ﬂÀ— „‰ 20 œﬁÌﬁÂ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
emad:
End Sub



Private Sub CommandButton19_Click()
Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Sheets(18).Range("a1:m32").PrintPreview
ThisWorkbook.Application.Visible = False
UserForm10.Show
End Sub



Private Sub CommandButton20_Click()
Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Sheets(18).Range("a34:c46").PrintPreview
ThisWorkbook.Application.Visible = False
UserForm10.Show
End Sub

Private Sub CommandButton21_Click()
Call ExportDataFinalV2


Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Sheets(26).Range("a1:t16").PrintPreview
ThisWorkbook.Application.Visible = False
UserForm10.Show
End Sub

Private Sub CommandButton22_Click()
ExportDeductionsDataFinal

Unload Me
ThisWorkbook.Application.Visible = True
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
Sheets(26).Range("a20:p37").PrintPreview
ThisWorkbook.Application.Visible = False
UserForm10.Show
End Sub

Private Sub CommandButton23_Click()
UserForm15.Show
End Sub

Private Sub CommandButton24_Click()
Unload Me

UserForm55.Show
End Sub














Private Sub CommandButton25_Click()
Unload Me

UserForm54.Show
End Sub

Private Sub CommandButton26_Click()
On Error Resume Next
Unload Me
UserForm9.Show
Unload Me

End Sub

Private Sub CommandButton27_Click()
Unload Me
UserForm56.Show
End Sub

Private Sub CommandButton28_Click()
Unload Me
UserForm39.Show

End Sub

Private Sub CommandButton29_Click()

End Sub

Private Sub CommandButton30_Click()
Unload Me
UserForm45.Show

End Sub

Private Sub CommandButton31_Click()
Unload Me
UserForm46.Show
End Sub

Private Sub CommandButton32_Click()

Call copy_from_to54_1

Call export_all_data_2
End Sub

Private Sub CommandButton33_Click()
Unload Me
UserForm64.Show
End Sub

Private Sub CommandButton34_Click()
Unload Me
UserForm79.Show
End Sub

Private Sub CommandButton35_Click()
Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Application.DisplayAlerts = True
    Application.Visible = True
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",true)"
    
    
    
    ' ≈Ìﬁ«› «· ‰»ÌÂ«  ·„‰⁄ ŸÂÊ— —”«∆· «·«” »œ«· «·„“⁄Ã…
    
    
    ' 1.  ÂÌ∆… «·‘—Ìÿ: Ã⁄· ⁄—÷ Label3 Ì”«ÊÌ ’›—« ﬁ»· «·»œ¡
    Me.Label3.Width = 0
    Me.Repaint
    DoEvents
    
    On Error GoTo Cleanup
    
    ' 2. »œ¡ «·⁄„·Ì…:  ﬁœ„ »”Ìÿ ··»œ«Ì… („À·« 10%)
    UpdateProgress 0.1
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False ' „‰⁄ «· ‰»ÌÂ«  ›Ì «· ÿ»Ìﬁ «·Œ›Ì
    
    ' 3.  ‰›Ì– «· ’œÌ— (‰’· ≈·Ï 50%)
    UpdateProgress 0.5
    Call ExportSalaryTapesToSpecificFolder
    
    ' 4. ≈ﬂ„«· «·⁄„·Ì«  (‰’· ≈·Ï 80%)
    UpdateProgress 0.8
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Cells(1, 1).value = "»Ì«‰«  „’œ—…"
    
    ' 5. Õ›Ÿ «·„·› (‰’· ≈·Ï 100%)
    xlBook.SaveAs "C:\Users\Public\ExportResult.xlsx"
    UpdateProgress 1
    
    ' «· ‰ŸÌ›
    xlBook.Close SaveChanges:=False
    xlApp.Quit
    
    MsgBox " „ «· ’œÌ— »‰Ã«Õ!", vbInformation
    GoTo FinalExit

Cleanup:
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    MsgBox "ÕœÀ Œÿ√: " & Err.Description, vbCritical
Application.DisplayAlerts = False
    
    Application.Visible = False
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
FinalExit:
    
End Sub

' ≈Ã—«¡ ›—⁄Ì ÌﬁÊ„ »“Ì«œ… ⁄—÷ Label3 («·√Õ„—) »‰«¡ ⁄·Ï «·‰”»… «·„∆ÊÌ…
Private Sub UpdateProgress(Percent As Double)
    ' ‰” Œœ„ ⁄—÷ Frame1 ﬂ„—Ã⁄ √ﬁ’Ï
    ' ‰ÿ—Õ 4 ·÷„«‰ »ﬁ«¡ «·‘—Ìÿ œ«Œ· ÕœÊœ «·≈ÿ«—
    Me.Label3.Width = (Me.Frame1.Width - 4) * Percent
    Me.Repaint
    DoEvents
End Sub

Private Sub CommandButton36_Click()
    UserForm77.Show
End Sub

Private Sub OptionButton1_Click()
If OptionButton1.value = True Then
OptionButton2.value = False
Call arb
End If
End Sub

Private Sub OptionButton2_Click()
On Error Resume Next
If OptionButton2.value = True Then
OptionButton1.value = False
If Not Intersect(Target, UserForm1.ActiveControl) Is Nothing Then
Call englsh
End If
End If
End Sub







Private Sub TextBox1_Change()
Sheets(45).Cells(1, "b").value = TextBox1.Text
End Sub

Private Sub UserForm_Activate()
 
End Sub

Private Sub UserForm_Click()

End Sub
