VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm47 
   Caption         =   "UserForm47"
   ClientHeight    =   10875
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   22140
   OleObjectBlob   =   "UserForm47.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm47"
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
On Error Resume Next
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

Function rightToLeftListView()
On Error Resume Next
Dim Columns As Long
Dim i As Long
With ListView1
Columns = .ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next
End With
End Function







Private Sub CheckBox1_Change()

 
End Sub



Private Sub CheckBox2_Click()
On Error Resume Next
If CheckBox2.value = True Then
CommandButton15.Enabled = True
Me.TextBox55.Text = ""
Me.TextBox54.Text = ""
Me.TextBox59.Text = ""
Else
CommandButton15.Enabled = False
End If
End Sub

Private Sub ComboBox1_Change()
On Error Resume Next
TextBox23.value = ComboBox1.value
ComboBox1.Visible = False
End Sub

Private Sub ComboBox1_DropButtonClick()

TextBox23.value = ComboBox1.value

End Sub

Private Sub ComboBox2_Change()

End Sub

Private Sub ComboBox3_Change()
On Error Resume Next
TextBox58.Text = ComboBox3.value
End Sub

Private Sub CommandButton1_Click()
Unload Me







End Sub

Private Sub CommandButton10_Click()

End Sub

Private Sub CommandButton11_Click()



End Sub

Private Sub CommandButton12_Click()



End Sub

Private Sub CommandButton13_Click()


End Sub

Private Sub CommandButton14_Click()



End Sub

Private Sub CommandButton15_Click()


 
End Sub

Private Sub CommandButton16_Click()

End Sub

Private Sub CommandButton17_Click()

Dim maxRow_S51 As Long, maxRow_S1 As Long, maxCol_Glob As Long
    Dim idxR_S1 As Long, idxR_S51 As Long
    Dim rowHeader_Glob As Long, rowDataStart_Glob As Long
    Dim id_In_S1 As String
    Dim recordFound_S51 As Boolean
    Dim targetRow_ToUpdate As Long
    
    rowHeader_Glob = 8
    rowDataStart_Glob = 9
    
    Application.ScreenUpdating = False
    
    '  ÕœÌœ «·‘Ì « 
    ' ≈ÌÃ«œ √»⁄«œ «·»Ì«‰«  ›Ì «·‘Ì  1
    maxCol_Glob = ThisWorkbook.Sheets(1).Cells(rowHeader_Glob, ThisWorkbook.Sheets(1).Columns.count).End(xlToLeft).Column
    maxRow_S1 = ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).rowS.count, 2).End(xlUp).row
    
    ' «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰« 
    If maxRow_S1 < rowDataStart_Glob Then
        MsgBox "·«  ÊÃœ »Ì«‰«  ›Ì Sheet1 ·· —ÕÌ·!", vbExclamation
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Õ·ﬁ…  ﬂ—«—Ì… ··„—Ê— ⁄·Ï ﬂ«›… ”Ã·«  «·‘Ì  1
    For idxR_S1 = rowDataStart_Glob To maxRow_S1
        id_In_S1 = Trim(ThisWorkbook.Sheets(1).Cells(idxR_S1, 2).value)
        
        If id_In_S1 <> "" Then
            ' ≈ÌÃ«œ ¬Œ— ’› ›Ì «·‘Ì  51 ··»ÕÀ
            maxRow_S51 = ThisWorkbook.Sheets(51).Cells(ThisWorkbook.Sheets(51).rowS.count, 2).End(xlUp).row
            If maxRow_S51 < rowDataStart_Glob Then maxRow_S51 = rowDataStart_Glob - 1
            
            recordFound_S51 = False
            
            ' «·»ÕÀ ⁄‰ «·—ﬁ„ «·ÊŸÌ›Ì ›Ì «·‘Ì  51
            For idxR_S51 = rowDataStart_Glob To maxRow_S51
                If Trim(ThisWorkbook.Sheets(51).Cells(idxR_S51, 2).value) = id_In_S1 Then
                    recordFound_S51 = True
                    targetRow_ToUpdate = idxR_S51
                    Exit For
                End If
            Next idxR_S51
            
            ' ≈–« ·„ ÌÊÃœ° ‰⁄ „œ «·’› «· «·Ì ·¬Œ— ’› „ÊÃÊœ
            If Not recordFound_S51 Then
                targetRow_ToUpdate = maxRow_S51 + 1
            End If
            
            ' ‰”Œ «·»Ì«‰«  Ê·’ﬁÂ« ﬂﬁÌ„
            ThisWorkbook.Sheets(1).Range(ThisWorkbook.Sheets(1).Cells(idxR_S1, 1), ThisWorkbook.Sheets(1).Cells(idxR_S1, maxCol_Glob)).Copy
            ThisWorkbook.Sheets(51).Cells(targetRow_ToUpdate, 1).PasteSpecial Paste:=xlPasteValues
        End If
    Next idxR_S1
    
    ' --- ﬂÊœ ≈⁄«œ… «· —ﬁÌ„ «· ·ﬁ«∆Ì ›Ì «·⁄„Êœ A ·‘Ì  51 ---
    maxRow_S51 = ThisWorkbook.Sheets(51).Cells(ThisWorkbook.Sheets(51).rowS.count, 2).End(xlUp).row
    If maxRow_S51 >= rowDataStart_Glob Then
        For idxR_S51 = rowDataStart_Glob To maxRow_S51
            ThisWorkbook.Sheets(51).Cells(idxR_S51, 1).value = idxR_S51 - (rowDataStart_Glob - 1)
        Next idxR_S51
    End If
    ' -----------------------------------------------------
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    MsgBox " „  ⁄„·Ì… «· ÕœÌÀ Ê«·≈œ—«Ã Ê≈⁄«œ… «· —ﬁÌ„ »‰Ã«Õ!", vbInformation






  End Sub





Private Sub CommandButton20_Click()
On Error Resume Next
Dim lastR As Integer
lastR = Sheets(51).Cells(rowS.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(51).Cells(Y, 5) = TextBox5.Text And Sheets(51).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(51).Cells(Y, 1) = TextBox1.Text
    Sheets(51).Cells(Y, 2) = TextBox2.Text
    Sheets(51).Cells(Y, 3) = TextBox3.Text
    Sheets(51).Cells(Y, 4) = TextBox4.value
    Sheets(51).Cells(Y, 5) = TextBox5.Text
    Sheets(51).Cells(Y, 6) = TextBox6.value
    Sheets(51).Cells(Y, 7) = TextBox7.Text
    Sheets(51).Cells(Y, 8) = TextBox8.Text
    Sheets(51).Cells(Y, 9) = TextBox9.Text
    Sheets(51).Cells(Y, 10) = TextBox10.Text
    Sheets(51).Cells(Y, 11) = TextBox11.Text
    Sheets(51).Cells(Y, 12) = TextBox12.Text
     Sheets(51).Cells(Y, 13) = TextBox13.value
    Sheets(51).Cells(Y, 14) = TextBox14.Text
    Sheets(51).Cells(Y, 15) = TextBox15.value
    Sheets(51).Cells(Y, 16) = TextBox16.Text
   Sheets(51).Cells(Y, 17) = TextBox17.value
    Sheets(51).Cells(Y, 18) = TextBox18.Text
    Sheets(51).Cells(Y, 19) = TextBox19.Text
    Sheets(51).Cells(Y, 20) = TextBox20.Text
    Sheets(51).Cells(Y, 21) = TextBox21.Text
    Sheets(51).Cells(Y, 22) = TextBox22.Text
    Sheets(51).Cells(Y, 23) = TextBox23.Text
    Sheets(51).Cells(Y, 24) = TextBox24.Text
    Sheets(51).Cells(Y, 25) = TextBox25.Text
   Sheets(51).Cells(Y, 26) = TextBox26.value
    Sheets(51).Cells(Y, 27) = TextBox27.Text
   Sheets(51).Cells(Y, 28) = TextBox28.value
  Sheets(51).Cells(Y, 29) = TextBox29.Text
    Sheets(51).Cells(Y, 30) = TextBox30.Text
    Sheets(51).Cells(Y, 31) = TextBox31.Text
  Sheets(51).Cells(Y, 32) = TextBox32.Text
   Sheets(51).Cells(Y, 33) = TextBox33.Text
    Sheets(51).Cells(Y, 34) = TextBox34.Text
    Sheets(51).Cells(Y, 35) = TextBox35.Text
    Sheets(51).Cells(Y, 36) = TextBox36.Text
    Sheets(51).Cells(Y, 37) = TextBox37.Text
    Sheets(51).Cells(Y, 38) = TextBox38.Text
    Sheets(51).Cells(Y, 39) = TextBox39.Text
    Sheets(51).Cells(Y, 40) = TextBox40.Text
    Sheets(51).Cells(Y, 41) = TextBox41.Text
    Sheets(51).Cells(Y, 42) = TextBox42.Text
    Sheets(51).Cells(Y, 43) = TextBox43.Text
    Sheets(51).Cells(Y, 44) = TextBox44.Text
    Sheets(51).Cells(Y, 45) = TextBox45.Text
    Sheets(51).Cells(Y, 46) = TextBox46.Text
    Sheets(51).Cells(Y, 47) = TextBox47.Text
  Sheets(51).Cells(Y, 48) = TextBox48.Text
   Sheets(51).Cells(Y, 49) = TextBox49.Text
   Sheets(51).Cells(Y, 50) = TextBox50.Text
   Sheets(51).Cells(Y, 51) = TextBox51.value
    Sheets(51).Cells(Y, 52) = TextBox52.Text
  Sheets(51).Cells(Y, 53) = TextBox53.Text
   Sheets(51).Cells(Y, 54) = TextBox54.Text
    Sheets(51).Cells(Y, 55) = TextBox55.Text
    Sheets(51).Cells(Y, 56) = TextBox56.Text
    Sheets(51).Cells(Y, 57) = TextBox57.Text
  Sheets(51).Cells(Y, 58) = TextBox58.Text
   Sheets(51).Cells(Y, 59) = TextBox59.Text
    Sheets(51).Cells(Y, 60) = TextBox60.Text
    Sheets(51).Cells(Y, 61) = TextBox61.Text
  Sheets(51).Cells(Y, 62) = TextBox62.Text
   Sheets(51).Cells(Y, 63) = TextBox63.Text
  Sheets(51).Cells(Y, 64) = TextBox64.Text
    
    
'    Sheets(3).Cells(y, 66) = Format(TextBox82.Text, "yyyy/mm/dd")
'  Sheets(3).Cells(y, 67) = TextBox83.Text
   
'   Sheets(3).Cells(y, 68) = TextBox84.Text
' Sheets(3).Cells(y, 69) = TextBox85.Text
'  Sheets(3).Cells(y, 70) = TextBox86.Text
'   Sheets(3).Cells(y, 71) = TextBox87.Text
'   Sheets(3).Cells(y, 72) = TextBox88.Text
   Sheets(51).Cells(Y, 73) = TextBox64.Text
   Sheets(51).Cells(Y, 75) = TextBox81.Text
    
          
     MsgBox " „  ⁄„·Ì…  ÕœÌÀ «·»Ì«‰«  »‰Ã«Õ", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
emad:
End Sub

Private Sub CommandButton21_Click()


End Sub

Private Sub CommandButton22_Click()




End Sub

Private Sub CommandButton23_Click()
On Error Resume Next
Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton24_Click()
On Error Resume Next
Dim wb As Workbook, lr1 As Double, lr2 As Double
Dim fil As Variant, dat As Double
Dim mainWb As Workbook
Dim targetSheet As Worksheet

Set mainWb = ThisWorkbook

'  ÕœÌœ «”„ «·‘Ì  «·„” Âœ› »«·«”„ ·÷„«‰ œﬁ… «·‰ﬁ·
On Error Resume Next
Set targetSheet = mainWb.Sheets(51) '  √ﬂœ √‰ «·«”„ Ìÿ«»ﬁ «”„ «·‘Ì  ⁄‰œﬂ
On Error GoTo 0

If targetSheet Is Nothing Then
    MsgBox "Œÿ√: ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·‘Ì ° Ì—ÃÏ «· √ﬂœ „‰ «”„ «·‘Ì  ›Ì «·ﬂÊœ!", vbCritical
    Exit Sub
End If

' ====== «·„”«— «·„»«‘— ›Ì «·»«— ‘‰ D Ê„Ã·œ data bank ======
Dim INF As String
INF = "D:\data bank"

' 1. «· √ﬂœ „‰ √‰ «·„Ã·œ «·—∆Ì”Ì „ÊÃÊœ ›⁄·Ì«
If Dir(INF, vbDirectory) = "" Then
    MsgBox "Œÿ√: «·„Ã·œ " & INF & " €Ì— „ÊÃÊœ ›Ì «·»«— ‘‰ D!", vbCritical, "«·„Ã·œ „›ﬁÊœ"
    Exit Sub
End If

' 2. «·»ÕÀ ⁄‰ „·› data_bank  ÕœÌœ« œ«Œ· «·„Ã·œ
fil = Dir(INF & "\data_bank.xl??")

' ====== «· ⁄œÌ· «·ÃœÌœ: «· Õﬁﬁ „‰ ÊÃÊœ «·„·› ﬁ»·  ⁄ÿÌ· «·‘«‘… √Ê »œ¡ «·⁄„· ======
If fil = "" Then
    MsgBox "«·„·› €Ì— „ Ê›—° Õ«Ê· „—«Ã⁄… Ê«· √ﬂœ „‰ ÊÃÊœ «·„·› «·–Ì ÌÕ ÊÌ ⁄·Ï »‰ﬂ «·»Ì«‰«  «·ﬂ·Ì… ·Ã„Ì⁄ „‰ ”»Ì «·ÂÌ√….", vbExclamation, " ‰»ÌÂ: «·„·› „›ﬁÊœ"
    Exit Sub
End If

'  ⁄ÿÌ· «· ÕœÌÀ«  »⁄œ «· √ﬂœ „‰ ÊÃÊœ «·„·› »«·ﬂ«„·
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'  ‰ŸÌ› «·»Ì«‰«  «·ﬁœÌ„… ›Ì «·‘Ì  «·—∆Ì”Ì
lr1 = targetSheet.Cells(rowS.count, 2).End(xlUp).row
If lr1 >= 9 Then
    targetSheet.Range("A9:bv" & lr1 + 1).ClearContents
End If

Do While fil <> ""
    '  Ã‰» › Õ «·„·› «·—∆Ì”Ì √Ê «·„·›«  «·„ƒﬁ …
    If fil <> mainWb.Name And Left(fil, 2) <> "~$" Then
        
        ' › Õ «·„·› «·„Õœœ
        Set wb = Workbooks.Open(INF & "\" & fil, ReadOnly:=True)
        
        ' Õ”«» ¬Œ— ”ÿ— ›Ì «·„·› «·—∆Ì”Ì
        lr1 = targetSheet.Cells(rowS.count, 2).End(xlUp).row
        If lr1 < 8 Then lr1 = 8 ' «·»œ«Ì… „‰ «·”ÿ— 9
        
        ' Õ”«» ¬Œ— ”ÿ— ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ›Ì «·„·› «·„› ÊÕ («·⁄„Êœ A)
        lr2 = wb.ActiveSheet.Cells(rowS.count, 1).End(xlUp).row
        
        ' ‰”Œ «·»Ì«‰«  ≈–« ﬂ«‰  „ Ê›—…
        If lr2 >= 9 Then
            wb.ActiveSheet.Range("A9:bv" & lr2).Copy targetSheet.Range("A" & lr1 + 1)
        ElseIf lr2 > 1 Then
            wb.ActiveSheet.Range("A2:bv" & lr2).Copy targetSheet.Range("A" & lr1 + 1)
        End If
        
        ' ≈€·«ﬁ «·„·› ›Ê—« ·„‰⁄  ⁄·ÌﬁÂ ›Ì «·–«ﬂ—… Œ·› «· ÿ»Ìﬁ
        wb.Close SaveChanges:=False
        Set wb = Nothing
        
    End If
    fil = Dir
Loop

Application.DisplayAlerts = True
Application.ScreenUpdating = True

targetSheet.Activate
MsgBox "·ﬁœ  „ ‰ﬁ· Ê«” Ì—«œ «·»Ì«‰«  „‰ „·› data_bank »‰Ã«Õ", vbInformation, "—”«·…  ‰»ÌÂ"

End Sub



Private Sub CommandButton25_Click()
Dim targetSheet As Worksheet
    Dim lastRow As Long

    '  ⁄ÌÌ‰ «·‘Ì  «·„” Âœ› (Ì›÷· ﬂ «»… «·«”„ »Ì‰ ⁄·«„ Ì  ‰’Ì’ ·÷„«‰ «·œﬁ…)
    Set targetSheet = ThisWorkbook.Sheets(51)
    
    ' ≈Ìﬁ«›  ÕœÌÀ «·‘«‘… · ”—Ì⁄ «·⁄„·Ì…
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Õ”«» ¬Œ— ’› ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ›Ì «·‘Ì  »‰«¡ ⁄·Ï «·⁄„Êœ B (√Ê √Ì ⁄„Êœ „„ ·∆ ⁄‰œﬂ)
    lastRow = targetSheet.Cells(targetSheet.rowS.count, 2).End(xlUp).row

    ' «· Õﬁﬁ „„« ≈–« ﬂ«‰  Â‰«ﬂ »Ì«‰«  ›⁄·Ì« „‰ «·’› 9 ›„« ›Êﬁ
    If lastRow >= 9 Then
        ' Õ–› «·„Õ ÊÌ«  Ê«· ‰”Ìﬁ«  „‰ «·’› 9 ÊÕ Ï ¬Œ— ’› Ê⁄„Êœ (BV „À·« √Ê «·Œ·«Ì« »«·ﬂ«„·)
        targetSheet.rowS("9:" & lastRow).ClearContents
        ' „·«Õÿ…: ≈–« ﬂ‰   —Ìœ Õ–› «· ‰”Ìﬁ«  √Ì÷« Ê«·√·Ê«‰ «” Œœ„ .Clear »œ·« „‰ .ClearContents
    End If

    ' ≈⁄«œ…  ›⁄Ì·  ÕœÌÀ «·‘«‘…
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox " „ Õ–› Ã„Ì⁄ »Ì«‰«  ﬁ«⁄œ… »Ì«‰«  »‰ﬂ «·»Ì«‰«  »‰Ã«Õ.", vbInformation, " √ﬂÌœ «·Õ–›"
End Sub

Private Sub CommandButton3_Click()
On Error GoTo emad

End Sub

Private Sub CommandButton4_Click()
Dim ws_Source As Worksheet, ws_Target As Worksheet
    Dim maxRow_Tar As Long, idxR_Src As Long
    Dim uniqueID_Sel As String
    Dim rowHeader_Glob As Long, rowDataStart_Glob As Long
    Dim foundRow_Src As Long
    Dim lastCol_Src As Long
    
    '  Õﬁﬁ „‰ √‰ «·„” Œœ„ ﬁœ «Œ «— ”ÿ—« „‰ «·‹ ListView1
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Ì—ÃÏ  ÕœÌœ «·”Ã· «·„—«œ  —ÕÌ·Â „‰ «·ﬁ«∆„… √Ê·«!", vbExclamation, " ‰»ÌÂ"
        Exit Sub
    End If
    
    '  ⁄ÌÌ‰ «·‘Ì «  Ê«·’›Ê› «·À«» …
    On Error Resume Next
    Set ws_Source = ThisWorkbook.Sheets(51)
    If ws_Source Is Nothing Then Set ws_Source = Sheet51
    Set ws_Target = ThisWorkbook.Sheets(1)
    If ws_Target Is Nothing Then Set ws_Target = Sheet1
    On Error GoTo 0
    
    rowHeader_Glob = 8
    rowDataStart_Glob = 9
    
    ' Ã·» «·„⁄—¯› «·›—Ìœ («·⁄„Êœ B) „‰ «·”ÿ— «·„Õœœ ›Ì «·‹ ListView
    ' „·«ÕŸ…: «·⁄„Êœ B ÂÊ «·⁄„Êœ «·À«‰Ì° ›Ì «·‹ ListView Ì„À· SubItems(1) ·√‰ «·⁄„Êœ «·√Ê· ÂÊ text
    uniqueID_Sel = Trim(ListView1.SelectedItem.ListSubItems(1).Text)
    
    If uniqueID_Sel = "" Then
        MsgBox "«·”Ã· «·„Õœœ ·« ÌÕ ÊÌ ⁄·Ï —ﬁ„ ÊŸÌ›Ì ›Ì «·⁄„Êœ B!", vbCritical, "Œÿ√"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 1. «·»ÕÀ ⁄‰ «·”ÿ— «·„ÿ«»ﬁ ›Ì Sheet51 »‰«¡ ⁄·Ï «·⁄„Êœ B
    foundRow_Src = 0
    Dim lastRow_Src As Long
    lastRow_Src = ws_Source.Cells(ws_Source.rowS.count, 2).End(xlUp).row
    
    For idxR_Src = rowDataStart_Glob To lastRow_Src
        If Trim(CStr(ws_Source.Cells(idxR_Src, 2).value)) = uniqueID_Sel Then
            foundRow_Src = idxR_Src
            Exit For
        End If
    Next idxR_Src
    
    ' ≈–«  „ «·⁄ÀÊ— ⁄·Ï «·”ÿ— ›Ì «·‘Ì  «·√”«”Ì
    If foundRow_Src > 0 Then
        ' ≈ÌÃ«œ ¬Œ— ⁄„Êœ Ê¬Œ— ’› ›«—€ ›Ì «·‘Ì  «·Âœ› (Sheet1)
        lastCol_Src = ws_Source.Cells(rowHeader_Glob, ws_Source.Columns.count).End(xlToLeft).Column
        maxRow_Tar = ws_Target.Cells(ws_Target.rowS.count, 2).End(xlUp).row
        If maxRow_Tar < rowHeader_Glob Then maxRow_Tar = rowHeader_Glob
        Dim nextEmptyRow_Tar As Long
        nextEmptyRow_Tar = maxRow_Tar + 1
        If nextEmptyRow_Tar < rowDataStart_Glob Then nextEmptyRow_Tar = rowDataStart_Glob
        
        ' ‰”Œ «·”ÿ— »«·ﬂ«„· »‰›” «·‰”ﬁ ≈·Ï Sheet1
        ws_Source.Range(ws_Source.Cells(foundRow_Src, 1), ws_Source.Cells(foundRow_Src, lastCol_Src)).Copy
        ws_Target.Cells(nextEmptyRow_Tar, 1).PasteSpecial xlPasteAll
        
        ' Õ–› «·”ÿ— „‰ Sheet51
        
        
        Application.CutCopyMode = False
        Application.ScreenUpdating = True
        
        MsgBox " „  —ÕÌ· «·”Ã· ≈·Ï Sheet1 ÊÕ–›Â „‰ Sheet51 »‰Ã«Õ.", vbInformation, "‰Ã«Õ «·⁄„·Ì…"
        
        '  ÕœÌÀ «·‹ ListView  ·ﬁ«∆Ì« »⁄œ «·Õ–› ·«” »⁄«œ «·”Ã· «·„›—€
        If Cpu_Is_TextBox72_Available Then Call TextBox72_Change
    Else
        Application.ScreenUpdating = True
        MsgBox "·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·”Ã· ›Ì Œ·«Ì« «·‘Ì ° ﬁœ ÌﬂÊ‰  „  ÕœÌÀÂ „”»ﬁ«.", vbCritical, "Œÿ√"
    End If
End Sub



Private Sub CommandButton6_Click()


End Sub



Private Sub CommandButton8_Click()
On Error Resume Next
Sheets(51).Activate
For X = 9 To 1500
If Cells(X, 5) = TextBox5.Text Then
Cells(X, 5).Select
Exit For
End If
Next X
'==================================



End Sub

Private Sub CommandButton7_Click()

End Sub

Private Sub CommandButton9_Click()

On Error GoTo emad
 With Me.ListView1
.ListItems.Clear
Dim item1 As ListItem
Dim last1, frw1 As Integer


Sheets(51).Activate
last1 = Sheets(51).Cells(rowS.count, 1).End(xlUp).row



'=============================================

   Sheets(51).Activate
    
    
    For frw1 = 2 To last1
                
  If Sheets(51).Cells(frw1, "h") = TextBox67.Text And Sheets(51).Cells(frw1, "f") = ComboBox2.value And ComboBox3.Text = Sheets(51).Cells(frw1, "bf").value Then
     
'==================================================

Set item1 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(51).Cells(frw1, "b")
item1.SubItems(2) = Sheets(51).Cells(frw1, "c")
item1.SubItems(3) = Sheets(51).Cells(frw1, "d")
item1.SubItems(4) = Sheets(51).Cells(frw1, "e")
item1.SubItems(5) = Sheets(51).Cells(frw1, "f")
item1.SubItems(6) = Sheets(51).Cells(frw1, "g")
item1.SubItems(7) = Sheets(51).Cells(frw1, "h")
item1.SubItems(8) = Sheets(51).Cells(frw1, "i")
item1.SubItems(9) = Sheets(51).Cells(frw1, "j")
item1.SubItems(10) = Sheets(51).Cells(frw1, "k")
item1.SubItems(11) = Sheets(51).Cells(frw1, "l")
item1.SubItems(12) = Sheets(51).Cells(frw1, "m")
item1.SubItems(13) = Sheets(51).Cells(frw1, "n")
item1.SubItems(14) = Sheets(51).Cells(frw1, "o")
item1.SubItems(15) = Sheets(51).Cells(frw1, "p")
item1.SubItems(16) = Sheets(51).Cells(frw1, "q")
item1.SubItems(17) = Sheets(51).Cells(frw1, "r")
item1.SubItems(18) = Sheets(51).Cells(frw1, "s")
item1.SubItems(19) = Sheets(51).Cells(frw1, "t")
item1.SubItems(20) = Sheets(51).Cells(frw1, "u")
item1.SubItems(21) = Sheets(51).Cells(frw1, "v")
item1.SubItems(22) = Sheets(51).Cells(frw1, "w")
item1.SubItems(23) = Sheets(51).Cells(frw1, "x")
item1.SubItems(24) = Sheets(51).Cells(frw1, "y")
item1.SubItems(25) = Sheets(51).Cells(frw1, "z")
item1.SubItems(26) = Sheets(51).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(51).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(51).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(51).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(51).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(51).Cells(frw1, "af")
item1.SubItems(32) = Sheets(51).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(51).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(51).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(51).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(51).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(51).Cells(frw1, "al")
item1.SubItems(38) = Sheets(51).Cells(frw1, "am")
item1.SubItems(39) = Sheets(51).Cells(frw1, "an")
item1.SubItems(40) = Sheets(51).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(51).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(51).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(51).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(51).Cells(frw1, "as")
item1.SubItems(45) = Sheets(51).Cells(frw1, "at")
item1.SubItems(46) = Sheets(51).Cells(frw1, "au")
item1.SubItems(47) = Sheets(51).Cells(frw1, "av")
item1.SubItems(48) = Sheets(51).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(51).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(51).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(51).Cells(frw1, "az")
item1.SubItems(52) = Sheets(51).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(51).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(51).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(51).Cells(frw1, "bd")
item1.SubItems(56) = Sheets(51).Cells(frw1, "be")
item1.SubItems(57) = Sheets(51).Cells(frw1, "bf")
item1.SubItems(58) = Sheets(51).Cells(frw1, "bg")
item1.SubItems(59) = Sheets(51).Cells(frw1, "bh")
item1.SubItems(60) = Sheets(51).Cells(frw1, "bi")
item1.SubItems(61) = Sheets(51).Cells(frw1, "bj")
item1.SubItems(62) = Sheets(51).Cells(frw1, "bk")
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("j9:j800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ad9:ad800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ae9:ae800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("av9:av800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("aw9:aw800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("x9:x800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(51).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")






UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")




End If
Next frw1
End With

'============================================= «·Õ”«»«  «·ﬂ·ÌÂ

Dim simadd As Double
Dim simsubtract As Double
Dim item2 As ListItem
Dim last2, frw2 As Integer
last2 = Sheets(51).Range("A" & rowS.count).End(xlUp).row
For frw2 = 2 To last2



If ComboBox2.value = "«·ﬂ·" And Sheets(51).Cells(frw2, "bf") = ComboBox3.value Then

Set item2 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw2, "A"))




item2.SubItems(1) = Sheets(51).Cells(frw2, "b")
item2.SubItems(2) = Sheets(51).Cells(frw2, "c")
item2.SubItems(3) = Sheets(51).Cells(frw2, "d")
item2.SubItems(4) = Sheets(51).Cells(frw2, "e")
item2.SubItems(5) = Sheets(51).Cells(frw2, "f")
item2.SubItems(6) = Sheets(51).Cells(frw2, "g")
item2.SubItems(7) = Sheets(51).Cells(frw2, "h")
item2.SubItems(8) = Sheets(51).Cells(frw2, "i")
item2.SubItems(9) = Sheets(51).Cells(frw2, "j")
item2.SubItems(10) = Sheets(51).Cells(frw2, "k")
item2.SubItems(11) = Sheets(51).Cells(frw2, "l")
item2.SubItems(12) = Sheets(51).Cells(frw2, "m")
item2.SubItems(13) = Sheets(51).Cells(frw2, "n")
item2.SubItems(14) = Sheets(51).Cells(frw2, "o")
item2.SubItems(15) = Sheets(51).Cells(frw2, "p")
item2.SubItems(16) = Sheets(51).Cells(frw2, "q")
item2.SubItems(17) = Sheets(51).Cells(frw2, "r")
item2.SubItems(18) = Sheets(51).Cells(frw2, "s")
item2.SubItems(19) = Sheets(51).Cells(frw2, "t")
item2.SubItems(20) = Sheets(51).Cells(frw2, "u")
item2.SubItems(21) = Sheets(51).Cells(frw2, "v")
item2.SubItems(22) = Sheets(51).Cells(frw2, "w")
item2.SubItems(23) = Sheets(51).Cells(frw2, "x")
item2.SubItems(24) = Sheets(51).Cells(frw2, "y")
item2.SubItems(25) = Sheets(51).Cells(frw2, "z")
item2.SubItems(26) = Sheets(51).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(51).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(51).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(51).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(51).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(51).Cells(frw2, "af")
item2.SubItems(32) = Sheets(51).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(51).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(51).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(51).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(51).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(51).Cells(frw2, "al")
item2.SubItems(38) = Sheets(51).Cells(frw2, "am")
item2.SubItems(39) = Sheets(51).Cells(frw2, "an")
item2.SubItems(40) = Sheets(51).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(51).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(51).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(51).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(51).Cells(frw2, "as")
item2.SubItems(45) = Sheets(51).Cells(frw2, "at")
item2.SubItems(46) = Sheets(51).Cells(frw2, "au")
item2.SubItems(47) = Sheets(51).Cells(frw2, "av")
item2.SubItems(48) = Sheets(51).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(51).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(51).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(51).Cells(frw2, "az")
item2.SubItems(52) = Sheets(51).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(51).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(51).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(51).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(51).Cells(frw2, "be")
item2.SubItems(57) = Sheets(51).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(51).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(51).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(51).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(51).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(51).Cells(frw2, "bk")
'item2.SubItems(63) = Sheets(51).Cells(frw2, "bl")
simadd = Application.WorksheetFunction.Sum(Sheets(51).Range("bh9:bh" & frw2))
simsubtract = Application.WorksheetFunction.Sum(Sheets(51).Range("bj9:bj" & frw2))
TextBox73.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("j9:j" & frw2))
TextBox74.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("ad9:ad" & frw2))
TextBox79.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("ae9:ae" & frw2))
TextBox80.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("av9:av" & frw2))
TextBox66.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("aw9:aw" & frw2))
TextBox76.Text = Application.WorksheetFunction.Sum(Sheets(51).Range("X9:X" & frw1))
TextBox67.Text = ""

TextBox66.Text = (TextBox66.Text + simadd) - simsubtract
'TextBox75.Text = Application.WorksheetFunction.Count(Sheets(51).Range("a9:a" & frw2))
'=COUNTIF(BF9:BF578;BL9)

'TextBox75.Text = ListView1.ListItems.Count
UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")




ElseIf Sheets(51).Cells(frw1, "h") Like "*" & TextBox67.Text & "*" And Sheets(51).Cells(frw1, "f") Like "*" & ComboBox2.Text & "*" And ComboBox3.value = Sheets(51).Cells(frw1, "bf").value Then
  item2.SubItems(1) = Sheets(51).Cells(frw2, "b")
item2.SubItems(2) = Sheets(51).Cells(frw2, "c")
item2.SubItems(3) = Sheets(51).Cells(frw2, "d")
item2.SubItems(4) = Sheets(51).Cells(frw2, "e")
item2.SubItems(5) = Sheets(51).Cells(frw2, "f")
item2.SubItems(6) = Sheets(51).Cells(frw2, "g")
item2.SubItems(7) = Sheets(51).Cells(frw2, "h")
item2.SubItems(8) = Sheets(51).Cells(frw2, "i")
item2.SubItems(9) = Sheets(51).Cells(frw2, "j")
item2.SubItems(10) = Sheets(51).Cells(frw2, "k")
item2.SubItems(11) = Sheets(51).Cells(frw2, "l")
item2.SubItems(12) = Sheets(51).Cells(frw2, "m")
item2.SubItems(13) = Sheets(51).Cells(frw2, "n")
item2.SubItems(14) = Sheets(51).Cells(frw2, "o")
item2.SubItems(15) = Sheets(51).Cells(frw2, "p")
item2.SubItems(16) = Sheets(51).Cells(frw2, "q")
item2.SubItems(17) = Sheets(51).Cells(frw2, "r")
item2.SubItems(18) = Sheets(51).Cells(frw2, "s")
item2.SubItems(19) = Sheets(51).Cells(frw2, "t")
item2.SubItems(20) = Sheets(51).Cells(frw2, "u")
item2.SubItems(21) = Sheets(51).Cells(frw2, "v")
item2.SubItems(22) = Sheets(51).Cells(frw2, "w")
item2.SubItems(23) = Sheets(51).Cells(frw2, "x")
item2.SubItems(24) = Sheets(51).Cells(frw2, "y")
item2.SubItems(25) = Sheets(51).Cells(frw2, "z")
item2.SubItems(26) = Sheets(51).Cells(frw2, "aa")
item2.SubItems(27) = Sheets(51).Cells(frw2, "ab")
item2.SubItems(28) = Sheets(51).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(51).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(51).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(51).Cells(frw2, "af")
item2.SubItems(32) = Sheets(51).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(51).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(51).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(51).Cells(frw2, "aj")
item2.SubItems(36) = Sheets(51).Cells(frw2, "ak")
item2.SubItems(37) = Sheets(51).Cells(frw2, "al")
item2.SubItems(38) = Sheets(51).Cells(frw2, "am")
item2.SubItems(39) = Sheets(51).Cells(frw2, "an")
item2.SubItems(40) = Sheets(51).Cells(frw2, "ao")
item2.SubItems(41) = Sheets(51).Cells(frw2, "ap")
item2.SubItems(42) = Sheets(51).Cells(frw2, "aq")
item2.SubItems(43) = Sheets(51).Cells(frw2, "ar")
item2.SubItems(44) = Sheets(51).Cells(frw2, "as")
item2.SubItems(45) = Sheets(51).Cells(frw2, "at")
item2.SubItems(46) = Sheets(51).Cells(frw2, "au")
item2.SubItems(47) = Sheets(51).Cells(frw2, "av")
item2.SubItems(48) = Sheets(51).Cells(frw2, "aw")
item2.SubItems(49) = Sheets(51).Cells(frw2, "ax")
item2.SubItems(50) = Sheets(51).Cells(frw2, "ay")
item2.SubItems(51) = Sheets(51).Cells(frw2, "az")
item2.SubItems(52) = Sheets(51).Cells(frw2, "ba")
item2.SubItems(53) = Sheets(51).Cells(frw2, "bb")
item2.SubItems(54) = Sheets(51).Cells(frw2, "bc")
item2.SubItems(55) = Sheets(51).Cells(frw2, "bd")
item2.SubItems(56) = Sheets(51).Cells(frw2, "be")
item2.SubItems(57) = Sheets(51).Cells(frw2, "bf")
item2.SubItems(58) = Sheets(51).Cells(frw2, "bg")
item2.SubItems(59) = Sheets(51).Cells(frw2, "bh")
item2.SubItems(60) = Sheets(51).Cells(frw2, "bi")
item2.SubItems(61) = Sheets(51).Cells(frw2, "bj")
item2.SubItems(62) = Sheets(51).Cells(frw2, "bk")
item2.SubItems(63) = Sheets(51).Cells(frw2, "bl")





'==================================================
TextBox73.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("j9:j800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ad9:ad800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox79.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("ae9:ae800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox80.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("av9:av800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox66.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("aw9:aw800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)
TextBox76.Text = Application.WorksheetFunction.SumIfs(Sheets(51).Range("x9:x800"), Sheets(51).Range("f9:f800"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("h9:h800"), "=" & Me.TextBox67.Text)


'=SUMIFS(J9:J571,F9:F571,BH12,H9:H571,BG12)
TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(51).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*")






UserForm1.TextBox73.value = Format(UserForm1.TextBox73.value, "#,## IQD")
UserForm1.TextBox74.value = Format(UserForm1.TextBox74.value, "#,## IQD")
UserForm1.TextBox79.value = Format(UserForm1.TextBox79.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
UserForm1.TextBox76.value = Format(UserForm1.TextBox76.value, "#,## IQD")
End If
Next






'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("h9:h571"), "=" & Me.TextBox67.Text, Sheets(51).Range("f9:f571"), "*" & Me.ComboBox2.Text & "*", Sheets(51).Range("BF9:BF571"), "=" & Me.ComboBox3.Value)


'TextBox75.Text = Application.WorksheetFunction.CountIfs(Sheets(51).Range("BF9:BF" & frw2), "=" & Me.ComboBox3.Text)
TextBox75.Text = ListView1.ListItems.count




emad:



End Sub



Private Sub Frame_Click()

End Sub

Private Sub ListBox1_Click()
On Error Resume Next
Sheets(51).Activate
If TextBox2.Text = "" And TextBox5.Text = "" Then
Image1.Picture = LoadPicture("D:\employ_pic\shdow_pic.jpg")
End If
For i = 0 To ListBox1.ListCount
    If ListBox1.Selected(i) = True Then
        For j = 1 To 64
        Controls("TextBox" & j).Text = Cells(ListBox1.List(i, 1), j)
        Next j
    End If
Next i
ComboBox1.Visible = False
'===================================
Dim last1, frw1 As Integer
last1 = Sheets(51).Range("A" & rowS.count).End(xlUp).row
For frw1 = 1 To last1

If Sheets(51).Cells(frw1, "b") = TextBox2.Text And Sheets(51).Cells(frw1, "e") = TextBox5.Text Then
TextBox81.Text = Sheets(51).Cells(frw1, "bw")
Image1.Picture = LoadPicture(Sheets(51).Cells(frw1, 64))

'TextBox58Text = Sheets(51).Cells(frw1, "bg")

End If
Next frw1
If TextBox49.Text <> "" Then
TextBox77.Text = ConvertNumberToText(TextBox49.Text, "œÌ‰«—", "")
Else
TextBox77.Text = 0
End If

TextBox53.value = Format(Date, "yyyy/mm/dd")
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ComboBox1.Visible = False
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
Call arb

 
End Sub



Private Sub TextBox10_Change()


End Sub

Private Sub TextBox11_Change()


End Sub

Private Sub TextBox12_Change()


End Sub

Private Sub TextBox13_Change()


End Sub

Private Sub TextBox14_Change()

End Sub

Private Sub TextBox15_Change()


End Sub

Private Sub TextBox16_Change()

End Sub

Private Sub TextBox17_Change()


End Sub

Private Sub TextBox18_Change()

End Sub

Private Sub TextBox19_Change()


End Sub

Private Sub TextBox2_Change()


End Sub

Private Sub TextBox21_Change()

End Sub

Private Sub TextBox22_Change()

End Sub

Private Sub TextBox23_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next

ComboBox1.Visible = True
ComboBox1.List = Sheets(4).Range("f1:f11").value

End Sub

Private Sub TextBox25_Change()

End Sub

Private Sub TextBox26_Change()

End Sub

Private Sub TextBox28_Change()

End Sub

Private Sub TextBox3_Change()



End Sub

Private Sub TextBox34_Change()

End Sub

Private Sub TextBox35_Change()

End Sub

Private Sub TextBox36_Change()

End Sub

Private Sub TextBox37_Change()

End Sub

Private Sub TextBox38_Change()

End Sub

Private Sub TextBox4_Change()
arb



End Sub

Private Sub TextBox40_Change()

End Sub

Private Sub TextBox41_Change()

End Sub

Private Sub TextBox42_Change()

End Sub

Private Sub TextBox43_Change()

End Sub

Private Sub TextBox44_Change()

End Sub

Private Sub TextBox45_Change()

End Sub

Private Sub TextBox46_Change()

End Sub

Private Sub TextBox5_Change()
arb
End Sub

Private Sub TextBox5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next

TextBox27.value = ""
ListBox1.Clear
End Sub



Private Sub TextBox6_Change()
arb

End Sub

Private Sub TextBox60_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Sheets(51).Activate
CommandButton5.Locked = True
TextBox50.SetFocus
For i = 9 To Sheets(51).Cells(rowS.count, 5).End(xlUp).row
ListBox1.AddItem
 ListBox1.List(i - 9, 0) = Cells(i, 5).value
  ListBox1.List(i - 9, 1) = i
      Next i
End Sub

Private Sub TextBox65_AfterUpdate()
On Error Resume Next
If TextBox72.Text = "#N/A" Then

 Beep
 Beep
Application.Speech.Speak " al    rakum   gair     mau juod"
End If
End Sub

Private Sub TextBox65_Change()

On Error GoTo emad
Sheets(51).Cells(3, 64) = TextBox65.Text
TextBox72.Text = Sheets(51).Cells(3, 65).Text
ListBox1.Text = TextBox72.Text
If ListBox1.Text <> "" Then
2 Beep
Application.Speech.Speak " al    rakum       mau juod"

End If
emad:
End Sub

Private Sub TextBox65_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
On Error Resume Next
TextBox65.Text = ""
TextBox77.Text = ""
TextBox72.Text = ""

End Sub

Private Sub TextBox67_Change()

End Sub

Private Sub TextBox67_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo emad
Call rightToLeftListView
'=================================================
With Me.ListView1
'Setting  «⁄œ«œ«  «·„ŸÂ—
.FullRowSelect = True '· Ÿ·Ì· ﬂ«„· «·’›
.Gridlines = True
.View = lvwReport  '⁄—÷ ŒÿÊÿ «·‘»ﬂ…
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' ·Ê‰ «·Œÿ
.Font.Size = 11 'ÕÃ„ «·Œÿ
.Font.Bold = False  '€«„ﬁ
Font.Name = "Calibri"
ListView1.ListItems.Clear
With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer
last1 = Sheets(51).Range("A" & rowS.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(51).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(51).Cells(frw1, "b")
item1.SubItems(2) = Sheets(51).Cells(frw1, "c")
item1.SubItems(3) = Sheets(51).Cells(frw1, "d")
item1.SubItems(4) = Sheets(51).Cells(frw1, "e")
item1.SubItems(5) = Sheets(51).Cells(frw1, "f")
item1.SubItems(6) = Sheets(51).Cells(frw1, "g")
item1.SubItems(7) = Sheets(51).Cells(frw1, "h")
item1.SubItems(8) = Sheets(51).Cells(frw1, "i")
item1.SubItems(9) = Sheets(51).Cells(frw1, "j")
item1.SubItems(10) = Sheets(51).Cells(frw1, "k")
item1.SubItems(11) = Sheets(51).Cells(frw1, "l")
item1.SubItems(12) = Sheets(51).Cells(frw1, "m")
item1.SubItems(13) = Sheets(51).Cells(frw1, "n")
item1.SubItems(14) = Sheets(51).Cells(frw1, "o")
item1.SubItems(15) = Sheets(51).Cells(frw1, "p")
item1.SubItems(16) = Sheets(51).Cells(frw1, "q")
item1.SubItems(17) = Sheets(51).Cells(frw1, "r")
item1.SubItems(18) = Sheets(51).Cells(frw1, "s")
item1.SubItems(19) = Sheets(51).Cells(frw1, "t")
item1.SubItems(20) = Sheets(51).Cells(frw1, "u")
item1.SubItems(21) = Sheets(51).Cells(frw1, "v")
item1.SubItems(22) = Sheets(51).Cells(frw1, "w")
item1.SubItems(23) = Sheets(51).Cells(frw1, "x")
item1.SubItems(24) = Sheets(51).Cells(frw1, "y")
item1.SubItems(25) = Sheets(51).Cells(frw1, "z")
item1.SubItems(26) = Sheets(51).Cells(frw1, "aa")
item1.SubItems(27) = Sheets(51).Cells(frw1, "ab")
item1.SubItems(28) = Sheets(51).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(51).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(51).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(51).Cells(frw1, "af")
item1.SubItems(32) = Sheets(51).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(51).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(51).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(51).Cells(frw1, "aj")
item1.SubItems(36) = Sheets(51).Cells(frw1, "ak")
item1.SubItems(37) = Sheets(51).Cells(frw1, "al")
item1.SubItems(38) = Sheets(51).Cells(frw1, "am")
item1.SubItems(39) = Sheets(51).Cells(frw1, "an")
item1.SubItems(40) = Sheets(51).Cells(frw1, "ao")
item1.SubItems(41) = Sheets(51).Cells(frw1, "ap")
item1.SubItems(42) = Sheets(51).Cells(frw1, "aq")
item1.SubItems(43) = Sheets(51).Cells(frw1, "ar")
item1.SubItems(44) = Sheets(51).Cells(frw1, "as")
item1.SubItems(45) = Sheets(51).Cells(frw1, "at")
item1.SubItems(46) = Sheets(51).Cells(frw1, "au")
item1.SubItems(47) = Sheets(51).Cells(frw1, "av")
item1.SubItems(48) = Sheets(51).Cells(frw1, "aw")
item1.SubItems(49) = Sheets(51).Cells(frw1, "ax")
item1.SubItems(50) = Sheets(51).Cells(frw1, "ay")
item1.SubItems(51) = Sheets(51).Cells(frw1, "az")
item1.SubItems(52) = Sheets(51).Cells(frw1, "ba")
item1.SubItems(53) = Sheets(51).Cells(frw1, "bb")
item1.SubItems(54) = Sheets(51).Cells(frw1, "bc")
item1.SubItems(55) = Sheets(51).Cells(frw1, "bf")




Next frw1

UserForm1.TextBox62.value = Format(UserForm1.TextBox62.value, "#,## IQD")
UserForm1.TextBox63.value = Format(UserForm1.TextBox63.value, "#,## IQD")
UserForm1.TextBox64.value = Format(UserForm1.TextBox64.value, "#,## IQD")
UserForm1.TextBox80.value = Format(UserForm1.TextBox80.value, "#,## IQD")
UserForm1.TextBox66.value = Format(UserForm1.TextBox66.value, "#,## IQD")
End With


'==============================================
    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
   
   
   
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    
    
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
    
    
    
    
    ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
    
    ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
    
    












'=================================================




End With

Me.TextBox61.value = ListView1.ListItems.count
 CheckBox1.Locked = False
CommandButton5.Enabled = False

ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
emad:
End Sub

Private Sub TextBox7_Change()
arb

End Sub

Private Sub TextBox72_Change()
On Error Resume Next
Dim c As Range, ss As Long, k As Long, i As Integer

ListBox1.Clear

Sheets(51).Activate
For i = 1 To 63
    Controls("TextBox" & i).Text = ""
Next i

If TextBox72 = "" Then Exit Sub
Sheets(51).Activate

ss = Sheets(51).Cells(rowS.count, 5).End(xlUp).row
k = 0
     
' ====== «· ⁄œÌ· Â‰«:  „  €ÌÌ— «·»œ«Ì… „‰ E2 ≈·Ï E9 ======
For Each c In Range("E9:E" & ss)
    If c.value Like TextBox72.value & "*" Then
        ListBox1.AddItem
        ListBox1.List(k, 0) = Cells(c.row, 5).value
        ListBox1.List(k, 1) = c.row
        k = k + 1
    End If
Next c

'=============================== ›· —… «··” »Êﬂ”

Dim ws_Fltr As Worksheet
    Dim maxRow_Fltr As Long, maxCol_Fltr As Long
    Dim idxR_Fltr As Long, idxC_Fltr As Long
    Dim rowHeader_Fltr As Long, rowDataStart_Fltr As Long
    Dim matrixData_Fltr As Variant
    Dim searchPhrase_Fltr As String
    Dim rowEntry_Fltr As ListItem
    Dim cellValue_E_Fltr As String
    Dim firstColumnValue As String
    Dim recordCount_Fltr As Long ' „ €Ì— ·Õ”«» ⁄œœ «·”Ã·«  «·„ÿ«»ﬁ…
    
    ' „Õ«Ê·…  ⁄ÌÌ‰ «·‘Ì  Õ”» «·«”„ «·»—„ÃÌ √Ê «”„ «· »ÊÌ»
    On Error Resume Next
    Set ws_Fltr = ThisWorkbook.Sheets(51)
    If ws_Fltr Is Nothing Then Set ws_Fltr = Sheet51
    On Error GoTo 0
    
    ' ≈–« ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·‘Ì  Ì „ «·Œ—ÊÃ
    If ws_Fltr Is Nothing Then Exit Sub
    
    rowHeader_Fltr = 8
    rowDataStart_Fltr = 9
    recordCount_Fltr = 0 '  ’›Ì— «·⁄œ«œ ›Ì «·»œ«Ì…
    
    ' Ã·» «·‰’ «·„ﬂ Ê» ›Ì TextBox72 Ê ‰ŸÌ›Â „‰ «·„”«›« 
    searchPhrase_Fltr = Trim(TextBox72.Text)
    
    ' ≈Ìﬁ«›  ÕœÌÀ «·‘«‘… „ƒﬁ « · ”—Ì⁄ «·√œ«¡ Ê„‰⁄ «·Ê„Ì÷
    Application.ScreenUpdating = False
    
    '  ›—Ì€ ’›Ê› «·‹ ListView1  „«„«
    ListView1.ListItems.Clear
    
    '  ÕœÌœ √»⁄«œ «·ÃœÊ· «·›⁄·Ì ›Ì «·‘Ì 
    maxCol_Fltr = ws_Fltr.Cells(rowHeader_Fltr, ws_Fltr.Columns.count).End(xlToLeft).Column
    maxRow_Fltr = ws_Fltr.Cells(ws_Fltr.rowS.count, 1).End(xlUp).row
    
    ' «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  Ê«·Ê’Ê· ··⁄„Êœ E («·⁄„Êœ 5) ⁄·Ï «·√ﬁ·
    If maxRow_Fltr >= rowDataStart_Fltr And maxCol_Fltr >= 5 Then
        
        ' ”Õ» ÃœÊ· «·»Ì«‰«  »«·ﬂ«„· ≈·Ï „’›Ê›… «·–«ﬂ—… ·”—⁄… ›«∆ﬁ…
        matrixData_Fltr = ws_Fltr.Range(ws_Fltr.Cells(rowDataStart_Fltr, 1), ws_Fltr.Cells(maxRow_Fltr, maxCol_Fltr)).value
        
        ' „⁄«·Ã… Õ«·… «·”ÿ— «·Ê«Õœ · ›«œÌ √Œÿ«¡ «·„’›Ê›« 
        If Not IsArray(matrixData_Fltr) Then
            ReDim matrixData_Fltr(1 To 1, 1 To maxCol_Fltr)
            For idxC_Fltr = 1 To maxCol_Fltr
                matrixData_Fltr(1, idxC_Fltr) = ws_Fltr.Cells(rowDataStart_Fltr, idxC_Fltr).value
            Next idxC_Fltr
        End If
        
        ' «·„—Ê— ⁄·Ï „’›Ê›… «·»Ì«‰«  ›Ì «·–«ﬂ—… ·›Õ’Â« Ê›· — Â«
        For idxR_Fltr = LBound(matrixData_Fltr, 1) To UBound(matrixData_Fltr, 1)
            
            ' Ã·» ﬁÌ„… «·⁄„Êœ «·Œ«„” (E) «·„Œ’’ ··«”„
            cellValue_E_Fltr = Trim(CStr(matrixData_Fltr(idxR_Fltr, 5)))
            
            ' ‘—ÿ «·›· —…: «· Õﬁﬁ „‰ „ÿ«»ﬁ… «·‰’ «·„ﬂ Ê» „⁄ «·⁄„Êœ E
            If searchPhrase_Fltr = "" Or InStr(1, cellValue_E_Fltr, searchPhrase_Fltr, vbTextCompare) > 0 Then
                
                ' ‘—ÿ ≈÷«›Ì: «· √ﬂœ „‰ √‰ «·⁄„Êœ B («·⁄„Êœ —ﬁ„ 2 ›Ì «·„’›Ê›…) ·Ì” ›«—€« ·«Õ ”«»Â ﬂ”Ã· ÕﬁÌﬁÌ
                If Not IsEmpty(matrixData_Fltr(idxR_Fltr, 2)) And Trim(CStr(matrixData_Fltr(idxR_Fltr, 2))) <> "" Then
                    
                    ' “Ì«œ… «·⁄œ«œ »„ﬁœ«— 1 ·ﬂÊ‰Â ”Ã· „ÿ«»ﬁ ÊÌÕ ÊÌ ⁄·Ï „⁄—› ÊŸÌ›Ì
                    recordCount_Fltr = recordCount_Fltr + 1
                    
                    '  √„Ì‰ «·⁄„Êœ «·√Ê· («·⁄„Êœ A) ·÷„«‰ ŸÂÊ— «·”ÿ—
                    If Not IsEmpty(matrixData_Fltr(idxR_Fltr, 1)) Then
                        firstColumnValue = Trim(CStr(matrixData_Fltr(idxR_Fltr, 1)))
                    Else
                        firstColumnValue = ""
                    End If
                    
                    ' ≈÷«›… «·”ÿ— «·√”«”Ì ›Ì «·‹ ListView1
                    Set rowEntry_Fltr = ListView1.ListItems.Add(, , firstColumnValue)
                    
                    ' ≈÷«›… »ﬁÌ… «·»Ì«‰«  ··√⁄„œ… «·›—⁄Ì…
                    For idxC_Fltr = 2 To maxCol_Fltr
                        If Not IsEmpty(matrixData_Fltr(idxR_Fltr, idxC_Fltr)) Then
                            rowEntry_Fltr.ListSubItems.Add , , Trim(CStr(matrixData_Fltr(idxR_Fltr, idxC_Fltr)))
                        Else
                            rowEntry_Fltr.ListSubItems.Add , , ""
                        End If
                    Next idxC_Fltr
                    
                End If
                
            End If
        Next idxR_Fltr
    End If
    
    ' ⁄—÷ «·⁄œœ «·Õ«·Ì «·„›· — ··”Ã·«  ›Ì TextBox82
    TextBox82.Text = Format(recordCount_Fltr, "#,##0")
    
    ' ≈⁄«œ…  ›⁄Ì·  ÕœÌÀ «·‘«‘…
    Application.ScreenUpdating = True


End Sub

Private Sub TextBox72_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
On Error Resume Next
Dim i As Long, lastRow As Long, k As Long
    
    ListBox1.Clear ' „”Õ «·ﬁ«∆„… «·ﬁœÌ„… √Ê·«
    Sheets(51).Activate
    TextBox72.Text = ""
    
    ' Õ”«» ¬Œ— ”ÿ— ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ›Ì «·⁄„Êœ 5
    lastRow = Sheets(51).Cells(rowS.count, 5).End(xlUp).row
    
    ' «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  »œ¡« „‰ «·’› 9 ›„« ›Êﬁ
    If lastRow < 9 Then Exit Sub
    
    k = 0 ' „ €Ì— · ÕœÌœ  — Ì» «·”ÿÊ— œ«Œ· «·‹ ListBox (Ì»œ√ „‰ 0)
    
    ' ====== «· ⁄œÌ· Â‰«: «·Õ·ﬁ…  »œ√ „‰ «·’› 9 ======
    For i = 9 To lastRow
        ListBox1.AddItem
        ListBox1.List(k, 0) = Cells(i, 5).value ' ⁄—÷ «·‰’ „‰ «·⁄„Êœ E
        ListBox1.List(k, 1) = i                 '  Œ“Ì‰ —ﬁ„ «·’› ›Ì «·⁄„Êœ «·À«‰Ì ··‹ ListBox
        k = k + 1                               ' «·«‰ ﬁ«· ··”ÿ— «· «·Ì ›Ì «·‹ ListBox
    Next i
    
    Dim ws_Count As Worksheet
    Dim maxRow_Count As Long
    Dim idxR_Count As Long
    Dim rowDataStart_Count As Long
    Dim columnB_Data As Variant
    Dim totalRecords_Count As Long
    
    ' „Õ«Ê·…  ⁄ÌÌ‰ «·‘Ì  Õ”» «·«”„ «·»—„ÃÌ √Ê «”„ «· »ÊÌ» · ›«œÌ «·√Œÿ«¡
    On Error Resume Next
    Set ws_Count = ThisWorkbook.Sheets(51)
    If ws_Count Is Nothing Then Set ws_Count = Sheet51
    On Error GoTo 0
    
    ' ≈–« ·„ Ì „ «·⁄ÀÊ— ⁄·Ï «·‘Ì  Ì „ «·Œ—ÊÃ ›Ê—«
    If ws_Count Is Nothing Then Exit Sub
    
    rowDataStart_Count = 9
    totalRecords_Count = 0 '  ’›Ì— «·⁄œ«œ
    
    ' ≈ÌÃ«œ ¬Œ— ’› ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ›Ì «·⁄„Êœ B
    maxRow_Count = ws_Count.Cells(ws_Count.rowS.count, 2).End(xlUp).row
    
    ' «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  «» œ«¡ „‰ «·’› 9 ›„« ›Êﬁ
    If maxRow_Count >= rowDataStart_Count Then
        
        ' ”Õ» »Ì«‰«  «·⁄„Êœ B ›ﬁÿ ≈·Ï „’›Ê›… «·–«ﬂ—… ·”—⁄… ›«∆ﬁ…
        columnB_Data = ws_Count.Range(ws_Count.Cells(rowDataStart_Count, 2), ws_Count.Cells(maxRow_Count, 2)).value
        
        ' „⁄«·Ã… Õ«·… «·”Ã· «·Ê«Õœ (≈–« ﬂ«‰ «·ÃœÊ· ÌÕ ÊÌ ⁄·Ï ’› Ê«Õœ ›ﬁÿ)
        If Not IsArray(columnB_Data) Then
            If Trim(CStr(columnB_Count)) <> "" Then totalRecords_Count = 1
        Else
            ' «·„—Ê— ⁄·Ï ⁄‰«’— «·„’›Ê›… ·⁄œ «·”Ã·«  «·ÕﬁÌﬁÌ… (€Ì— «·›«—€…)
            For idxR_Count = LBound(columnB_Data, 1) To UBound(columnB_Data, 1)
                If Not IsEmpty(columnB_Data(idxR_Count, 1)) And Trim(CStr(columnB_Data(idxR_Count, 1))) <> "" Then
                    totalRecords_Count = totalRecords_Count + 1
                End If
            Next idxR_Count
        End If
    End If
    
    ' ⁄—÷ «·⁄œœ «·≈Ã„«·Ì «·ﬂ·Ì ›Ì TextBox82 » ‰”Ìﬁ —ﬁ„Ì Ê«÷Õ
    TextBox82.Text = Format(totalRecords_Count, "#,##0")
    
    ' ≈·€«¡ «·≈Ã—«¡ «·«› —«÷Ì ··œ»· ﬂ·Ìﬂ œ«Œ· «· Ìﬂ”  »Êﬂ” („À·  ÕœÌœ «·‰’ »«·ﬂ«„·)
    Cancel = True
    
End Sub

Private Sub TextBox8_Change()


End Sub

Private Sub TextBox81_Change()

End Sub

Private Sub TextBox9_Change()


End Sub
Private Sub UserForm_Activate()

On Error GoTo emad

Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

ComboBox2.List = Sheets(5).Range("a1:a373").value
ComboBox3.List = Sheets(4).Range("g1:g13").value
Me.TextBox67.Text = ""
'========================================================== Õ„Ì· »‰ﬂ «·»Ì«‰« 





















emad:

End Sub





Private Sub UserForm_Initialize()
'===============================
On Error Resume Next

'=========================
Call rightToLeftListView


Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim headerRow As Long, dataStartRow As Long
    Dim dataArray As Variant
    Dim lstItem As ListItem
    
    '  ÕœÌœ «·‘Ì  Ê«·’›Ê› «·√”«”Ì…
    Set ws = ThisWorkbook.Sheets(51) '  √ﬂœ „‰ √‰ «·«”„ „ÿ«»ﬁ  „«„«
    headerRow = 8
    dataStartRow = 9
    
    ' ≈⁄œ«œ«  «·‹ ListView Ê ›—Ì€Â«  „«„«
    With ListView1
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
    End With
    
    ' 1. ≈ÌÃ«œ ¬Œ— ⁄„Êœ ÌÕ ÊÌ ⁄·Ï »Ì«‰«  ›Ì ’› —ƒÊ” «·√⁄„œ… («·’› 8)
    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub ' Œ—ÊÃ ›Ì Õ«· ﬂ«‰ «·’› ›«—€«
    
    ' ≈÷«›… —ƒÊ” «·√⁄„œ… œÌ‰«„ÌﬂÌ« „‰ «·’› 8
    For j = 1 To lastCol
        ListView1.ColumnHeaders.Add , , ws.Cells(headerRow, j).value, 100
    Next j
    
    ' 2. ≈ÌÃ«œ ¬Œ— ’› ÌÕ ÊÌ ⁄·Ï »Ì«‰«  «» œ«¡ „‰ «·’› 9
    lastRow = ws.Cells(ws.rowS.count, 1).End(xlUp).row
    
    ' «· Õﬁﬁ „‰ ÊÃÊœ »Ì«‰«  „‰ «·’› 9 ›„« ›Êﬁ
    If lastRow >= dataStartRow Then
        
        ' ”Õ» «·»Ì«‰«  »«·ﬂ«„· ≈·Ï „’›Ê›… · ”—Ì⁄ «·„⁄«·Ã… ›Ì «·–«ﬂ—…
        dataArray = ws.Range(ws.Cells(dataStartRow, 1), ws.Cells(lastRow, lastCol)).value
        
        ' ›Ì Õ«· ﬂ«‰  «·»Ì«‰«  ’›« Ê«Õœ« ›ﬁÿ°  ÕÊÌ·Â« ·„’›Ê›… À‰«∆Ì… «·√»⁄«œ · Ã‰» «·√Œÿ«¡
        If Not IsArray(dataArray) Then
            ReDim dataArray(1 To 1, 1 To lastCol)
            For j = 1 To lastCol
                dataArray(1, j) = ws.Cells(dataStartRow, j).value
            Next j
        End If
        
        '  ⁄»∆… «·‹ ListView »«·»Ì«‰«  „‰ „’›Ê›… «·–«ﬂ—… »”—⁄…
        For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            ' ≈÷«›… «·”ÿ— «·√Ê· («·⁄„Êœ «·—∆Ì”Ì)
            Set lstItem = ListView1.ListItems.Add(, , CStr(dataArray(i, 1)))
            
            ' ≈÷«›… »ﬁÌ… «·√⁄„œ… «·›—⁄Ì… ··”ÿ—
            For j = 2 To lastCol
                If Not IsEmpty(dataArray(i, j)) Then
                    lstItem.ListSubItems.Add , , CStr(dataArray(i, j))
                Else
                    lstItem.ListSubItems.Add , , ""
                End If
            Next j
        Next i
    End If
End Sub

Private Function Cpu_Is_TextBox72_Available() As Boolean
    On Error Resume Next
    Dim t As String
    t = TextBox72.Text
    Cpu_Is_TextBox72_Available = (Err.Number = 0)
    On Error GoTo 0
End Function

