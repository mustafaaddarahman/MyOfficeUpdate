VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm73 
   Caption         =   "‰Ÿ«„ «·—Ê« » «·„ÿÊ— - ⁄„«œ „Õ„œ «·‰⁄Ì„Ì"
   ClientHeight    =   13440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   24780
   OleObjectBlob   =   "UserForm73.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm73"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Public NextTop As Double
Public GroupID As Long
Public DelButtons As New Collection

Private Sub CommandButton1_Click()
Unload Me
End Sub

Private Sub CommandButton2_Click()
  Dim i As Long
    ' 1.  Ã„Ìœ  ÕœÌÀ «·‘«‘… ·„‰⁄ «· Â‰ÌÃ  „«„« √À‰«¡ «·Õ–›
    Application.ScreenUpdating = False

    ' 2. Õ–› «·„ﬂÊ‰«  «·œÌ‰«„ÌﬂÌ… ›ﬁÿ («· Ì  Õ„· Tag)
    ' ‰»œ√ „‰ «·√ŒÌ— ··√Ê· ·÷„«‰ „”Õ «·–«ﬂ—… »·„Õ… »’—
    For i = Me.Controls.count - 1 To 0 Step -1
        ' ›Õ’: ‰Õ–› ›ﬁÿ „«  „ ≈‰‘«ƒÂ »—„ÃÌ« (txt, lbl, btnDel)
        If Me.Controls(i).Tag <> "" Or Left(Me.Controls(i).Name, 7) = "btnDel_" Then
            Me.Controls.Remove Me.Controls(i).Name
        End If
    Next i

    ' 3. --- [ ÷»ÿ ‰ﬁÿ… «·»œ«Ì… «·’ÕÌÕ… ] ---
    '  √ﬂœ √‰ Â–« «·—ﬁ„ (120 „À·«) ÂÊ «·„”«›…  Õ  √“—«— «·»ÕÀ Ê«·⁄‰«ÊÌ‰ «·À«» …
    NextTop = 120
    
    '  ’›Ì— ⁄œ«œ «·„Ã„Ê⁄«  · »œ√ «·≈÷«›… „‰ —ﬁ„ 1
    GroupID = 0
    
    ' 4. ≈⁄«œ… ÷»ÿ „‰ÿﬁ… «· „—Ì— (Scroll) Ê«·⁄Êœ… ··√⁄·Ï
    Me.ScrollTop = 0
    Me.ScrollHeight = Me.InsideHeight

    ' 5.  ÕœÌÀ «·Ê«ÃÂ… (√„— Repaint »œÊ‰ ⁄·«„… =)
    Application.ScreenUpdating = True
    Me.Repaint

    ' 6. ﬂ «»… «·‰ ÌÃ… ›Ì ‘—Ìÿ Õ«·… «·≈ﬂ”· (»œÊ‰ „”Ã »Êﬂ”)
    
    Application.StatusBar = " „ ≈⁄«œ… ÷»ÿ «·»Ì«‰«  Ê ’›Ì—Â« »‰Ã«Õ"

Set DelButtons = New Collection ' Â–« Ì„”Õ ﬂ· «·„›« ÌÕ «·ﬁœÌ„… „‰ «·–«ﬂ—… ›Ê—«
    
    NextTop = 120
    GroupID = 0

End Sub

Private Sub UserForm_Initialize()
    NextTop = 60
    GroupID = 0
    Me.ScrollHeight = Me.Height
End Sub

Private Sub txtSearch_Change()
    Dim ws1 As Worksheet: Set ws1 = Sheets("Sheet1")
    ListBox1.Clear
    Dim i As Long
    For i = 9 To ws1.Cells(ws1.rowS.count, "E").End(xlUp).row
        If InStr(1, ws1.Cells(i, "E"), txtSearch.Text, 1) > 0 Or InStr(1, ws1.Cells(i, "B"), txtSearch.Text, 1) > 0 Then
            ListBox1.AddItem ws1.Cells(i, "B")
            ListBox1.List(ListBox1.ListCount - 1, 1) = ws1.Cells(i, "E")
        End If
    Next i
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim ws1 As Worksheet: Set ws1 = Sheets("Sheet1")
    Dim ws9 As Worksheet: Set ws9 = Sheets("Sheet9")
    Dim r As Long, sVal As Variant
    If ListBox1.ListIndex = -1 Then Exit Sub
    sVal = ListBox1.value
    On Error Resume Next
    r = Application.match(CLng(sVal), ws1.Range("B:B"), 0)
    If Err.Number <> 0 Then
        Err.Clear
        r = Application.match(CStr(sVal), ws1.Range("B:B"), 0)
    End If
    On Error GoTo 0
    If r = 0 Then Exit Sub
    ws9.Range("G18") = ws1.Cells(r, "H")
    ws9.Range("F18") = ws1.Cells(r, "I") + 1
    GroupID = GroupID + 1
    CreateRow ListBox1.List(ListBox1.ListIndex, 1), CStr(sVal), CDbl(ws9.Range("H18").value), r, GroupID
    
    
    
    
    
    
End Sub

Private Sub btnSimulate_Click()
    Dim ctrl As Control: Dim t As Double: t = 0
    For Each ctrl In Me.Controls
        If InStr(ctrl.Name, "diff_") > 0 Then t = t + val(ctrl.Text)
    Next ctrl
    Dim res As Double: res = val(txtBudget.Text) - t
    If res >= 0 Then
        MsgBox "«·„Ì“«‰Ì…  ﬂ›Ì. «·›«∆÷: " & res, vbInformation
    Else
        MsgBox "⁄Ã“ „«·Ì: " & Abs(res), vbCritical
    End If
End Sub

Private Sub btnPrint_Click()
   Dim ctrl As MSForms.Control
    Dim wbNew As Workbook
    Dim wsTemp As Worksheet
    Dim folderPath As String, fileName As String
    Dim r As Long, c As Long
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    ' 1.  ÕœÌœ „”«— «·„Ã·œ («·»«— ‘‰ D)
    folderPath = "D:\exportprint\"
    
    ' --- [ ÕÌ·… –ﬂÌ…: ≈‰‘«¡ «·„Ã·œ »—„ÃÌ« ≈–« ﬂ«‰ „›ﬁÊœ« ] ---
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        MkDir folderPath ' ≈‰‘«¡ «·„Ã·œ
        On Error GoTo 0
        Application.StatusBar = " „ ≈‰‘«¡ «·„Ã·œ «·ÃœÌœ ›Ì D:\exportprint"
    End If

    ' 2. Ã„⁄ √—ﬁ«„ «·‹ Tag ··„ÊŸ›Ì‰ «·„ÊÃÊœÌ‰ Õ«·Ì« ›ﬁÿ ⁄·Ï «·‘«‘…
    For Each ctrl In Me.Controls
        ' ‰ √ﬂœ √‰ «·ﬂ«∆‰ ·Â Tag Ê√‰Â ·Ì” ›«—€«
        If ctrl.Tag <> "" Then
            If Not dict.exists(CStr(ctrl.Tag)) Then
                dict.Add CStr(ctrl.Tag), ""
            End If
        End If
    Next ctrl

    ' ›Õ’ ≈–« ﬂ«‰ «·›Ê—„ ›«—€«
    If dict.count = 0 Then
        MsgBox "√” «– ⁄„«œ° ·«  ÊÃœ ﬁÌÊœ Õ«·Ì« · ’œÌ—Â«!", vbExclamation
        Exit Sub
    End If

    ' 3. ≈‰‘«¡ „·› ≈ﬂ”· ÃœÌœ ›Ì «·–«ﬂ—… («·–«ﬂ—… «·„ƒﬁ …)
    Application.ScreenUpdating = False
    Set wbNew = Workbooks.Add
    Set wsTemp = wbNew.Sheets(1)
    
    '  ‰”Ìﬁ «·‘Ì  (Ì„Ì‰ ≈·Ï Ì”«—)
    wsTemp.DisplayRightToLeft = True
    
    ' 4.  ›—Ì€ »Ì«‰«  «· Ìﬂ”  »Êﬂ”«  Ê«·√”„«¡ »‰«¡ ⁄·Ï «·‹ Tag
    r = 1 ' ‰»œ√ «·’› «·√Ê·
    Dim key As Variant
    For Each key In dict.Keys
        ' ·ﬂ· „ÊŸ› (Tag)° ‰»ÕÀ ⁄‰ „ﬂÊ‰« Â
        For Each ctrl In Me.Controls
            If CStr(ctrl.Tag) = CStr(key) Then
                ' ≈–« ﬂ«‰ «·ﬂ«∆‰ ·Ì»· «·«”„
                If Left(ctrl.Name, 6) = "lblNm_" Then
                    wsTemp.Cells(r, 1).value = ctrl.Caption
                ' ≈–« ﬂ«‰ ’‰œÊﬁ »Ì«‰«  (TextBox)
                ElseIf TypeName(ctrl) = "TextBox" Then
                    '  Ê“Ì⁄ «·√⁄„œ… »‰«¡ ⁄·Ï „Êﬁ⁄ «·’‰œÊﬁ „‰ «·Ì„Ì‰ ··Ì”«—
                    ' 920 ÂÊ «·⁄—÷ «·√ﬁ’Ï° ‰ﬁ”„Â ⁄·Ï 68 (⁄—÷ «·’‰œÊﬁ) · — Ì» «·√⁄„œ…
                    c = Int((945 - ctrl.Left) / 68) + 2
                    wsTemp.Cells(r, c).value = ctrl.value
                End If
            End If
        Next ctrl
        r = r + 1 ' «·«‰ ﬁ«· ··„ÊŸ› «· «·Ì (”ÿ— ÃœÌœ)
    Next key

    ' 5. Õ›Ÿ «·„·› Ê ”„Ì Â »«·Êﬁ  Ê«· «—ÌŒ ·⁄œ„ «· ﬂ—«—
    fileName = " ﬁ—Ì—_—Ê« »_" & Format(Now, "yyyy-mm-dd_HH-mm-ss") & ".xlsx"
    
    '  Ã„Ì· «·„·› ﬁ»· «·Õ›Ÿ
    wsTemp.Columns.AutoFit
    wsTemp.Range("A1").CurrentRegion.Borders.LineStyle = 1
    
    ' «·Õ›Ÿ «·‰Â«∆Ì Ê≈€·«ﬁ «·„·› „‰ «·–«ﬂ—…
    On Error Resume Next
    wbNew.SaveAs folderPath & fileName
    wbNew.Close SaveChanges:=False
    On Error GoTo 0

    ' 6.  ‰ŸÌ› «·–«ﬂ—… «·„ƒﬁ …
    Set dict = Nothing: Set wsTemp = Nothing: Set wbNew = Nothing
    Application.ScreenUpdating = True
    
    MsgBox " „  ’œÌ— " & r - 1 & " ﬁÌœ »‰Ã«Õ ≈·Ï:" & vbCrLf & folderPath & fileName, vbInformation, "‰Ã«Õ «·⁄„·Ì…"
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 
Dim objSearch As MSForms.Control ' «·„ €Ì— «·√”«”Ì ··»ÕÀ ⁄‰ «·“—
    Dim objCheck As MSForms.Control  ' „ €Ì— ÃœÌœ  „«„« ·›Õ’ «·›—«€
    Dim i As Long
    Dim targetTag As String
    Dim hasDynamicControls As Boolean
    
    Dim curX As Single: curX = X
    Dim curY As Single: curY = Y + Me.ScrollTop

    ' 1. «·Õ·ﬁ… «·—∆Ì”Ì… ··»ÕÀ ⁄‰ “— «·Õ–›
    For Each objSearch In Me.Controls
        If Left(objSearch.Name, 7) = "btnDel_" Then
            If curX >= objSearch.Left And curX <= (objSearch.Left + objSearch.Width) And _
               curY >= objSearch.Top And curY <= (objSearch.Top + objSearch.Height) Then
                
                targetTag = objSearch.Tag
                
                ' 2.  ‰›Ì– ⁄„·Ì… «·Õ–›
                For i = Me.Controls.count - 1 To 0 Step -1
                    If Me.Controls(i).Tag = targetTag Then
                        Me.Controls.Remove Me.Controls(i).Name
                    End If
                Next i
                
                ' --- [ ›Õ’ –ﬂÌ: Â· √’»Õ «·›Ê—„ ›«—€«ø ] ---
                hasDynamicControls = False
                ' ‰” Œœ„ „ €Ì— objCheck Â‰« »œ·« „‰ objSearch ·„‰⁄ «·Œÿ√
                For Each objCheck In Me.Controls
                    If objCheck.Tag <> "" Then
                        hasDynamicControls = True
                        Exit For
                    End If
                Next objCheck
                
                ' 3. «· —”Ì  «· ·ﬁ«∆Ì ≈–« ·„ Ì »ﬁ √Ì ﬁÌœ
                If Not hasDynamicControls Then
                    NextTop = 120 ' «·ﬁ„… «· Ì  »œ√ ⁄‰œÂ« «·ﬁÌÊœ
                    GroupID = 0
                    Me.ScrollTop = 0
                    Me.ScrollHeight = Me.InsideHeight
                    Application.StatusBar = " „ Õ–› «·ﬁÌœ «·√ŒÌ— Ê ’›Ì— „Êﬁ⁄ «·≈÷«›… ··ﬁ„…."
                Else
                    Application.StatusBar = " „ Õ–› «·ﬁÌœ..  »ﬁÏ ﬁÌÊœ √Œ—Ï ›Ì «·›Ê—„."
                End If
                
                Me.Repaint
                Exit Sub ' «·Œ—ÊÃ ›Ê—« »⁄œ «·Õ–› Ê«· —”Ì 
            End If
        End If
    Next objSearch
 
End Sub

Sub CreateRow(Nm As String, ID As String, Nom As Double, rIdx As Long, gID As Long)
 Dim ws1 As Worksheet: Set ws1 = Sheets("Sheet1")
    Dim curR As Double: curR = 920
    Dim lbl As MSForms.Label ' <--- «·Õ· Â‰«:  ⁄—Ì› «·„ €Ì— «·„›ﬁÊœ
    
    ' 1. ≈÷«›… CheckBox ›Ì ‰Â«Ì… «·”ÿ— (ﬂ„« « ›ﬁ‰«)
    ' ”‰ﬁÊ„ »≈÷«› Â ·«Õﬁ« »⁄œ Õ”«» «·’‰«œÌﬁ
    
    ' 2. ≈‰‘«¡ ·Ì»· ≈·€«¡ «·„ÊŸ› («·⁄·ÊÌ) Ê ⁄ÌÌ‰Â ··„ €Ì— lbl
    Set lbl = Me.Controls.Add("Forms.Label.1", "btnDel_" & gID)
    With lbl
        .Caption = " [ ≈·€«¡ ] "
        .Top = NextTop: .Left = 945: .Width = 45: .Tag = gID
        .BackColor = &H8080FF: .ForeColor = &HFFFFFF: .TextAlign = 2: .Font.Bold = True
        .Enabled = True '  √ﬂœ √‰Â« True ·ﬂÌ Ì” ÃÌ» ··‰ﬁ—
    End With

    ' --- «·—»ÿ »«·‹ Class Module (·Ì⁄„· «·Õ–› «·‘—ÿÌ) ---
    Dim objDel As New clsDeleteButton
    Set objDel.DelLabel = lbl
    Set objDel.ParentForm = Me
    DelButtons.Add objDel, "K" & gID
    ' --------------------------------------------------

    ' 3. ·Ì»· «”„ «·„ÊŸ›
    With Me.Controls.Add("Forms.Label.1", "lblNm_" & gID)
        .Caption = Nm: .Top = NextTop: .Left = 20: .Width = 920: .Tag = gID
        .Font.Name = "Times New Roman": .Font.Size = 14: .Font.Bold = True
        .ForeColor = 16711680: .TextAlign = 3
    End With

    NextTop = NextTop + 30
    
    ' --- ﬂÊœ »‰«¡ «·’‰«œÌﬁ (AddBox) ﬂ„« ÂÊ ·œÌﬂ ---
    Dim rC, fC, j As Integer: rC = Array("M", "O", "Q", "Z", "AB"): fC = Array("S", "T", "U", "Y", "BH", "BP", "BR")
    AddBox Nom, ws1.Range("J8").value, curR, gID
    curR = curR - 68: Dim s As Double: s = Nom
    For j = 0 To 4: Dim v: v = Nom * ws1.Cells(rIdx, rC(j)): s = s + v: AddBox Format(v, "0"), ws1.Cells(8, rC(j)).value, curR, gID: curR = curR - 68: Next j
    For j = 0 To 6: Dim fV: fV = ws1.Cells(rIdx, fC(j)): s = s + fV: AddBox fV, ws1.Cells(8, fC(j)).value, curR, gID: curR = curR - 68: Next j
    AddBox Format(s - ws1.Cells(rIdx, "AE"), "0"), "«·“Ì«œ…", curR, gID, True
    
    ' 4. ≈÷«›… «·‹ CheckBox »⁄œ ¬Œ— ’‰œÊﬁ («·“Ì«œ…)
    With Me.Controls.Add("Forms.CheckBox.1", "chkDel_" & gID)
        .Caption = "": .Top = NextTop: .Left = curR - 30: .Width = 25: .Height = 25: .Tag = gID
    End With

    NextTop = NextTop + 90
    If NextTop > Me.Height Then Me.ScrollHeight = NextTop + 50
End Sub

Sub AddBox(v, t, r, gID, Optional isRes As Boolean = False)
    Dim tNm As String: If isRes Then tNm = "diff_" & gID Else tNm = "box_" & gID
    With Me.Controls.Add("Forms.Label.1"): .Caption = t: .Top = NextTop: .Left = r: .Width = 65: .Tag = gID: .TextAlign = 2: .Font.Size = 8: .Font.Bold = True: End With
    With Me.Controls.Add("Forms.TextBox.1", tNm): .Text = v: .Top = NextTop + 18: .Left = r: .Width = 65: .Height = 28: .Tag = gID: .TextAlign = 2: .Font.Size = 12: .Font.Bold = True
    If isRes Then .BackColor = 9240494: .Width = 80:
    End With
     
    
    End Sub


