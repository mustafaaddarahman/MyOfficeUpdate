Attribute VB_Name = "SAAD_EMAD"

Option Explicit

Sub Build_Final_System_73()
    Dim vbaProject As Object: Set vbaProject = ThisWorkbook.VBProject
    Dim vbaForm As Object
    Dim chkExist As Object
    
    ' 1.  ‰ŸÌ› «·‰”Œ «·ﬁœÌ„…
    On Error Resume Next
    Set chkExist = vbaProject.VBComponents("UserForm73")
    If Not chkExist Is Nothing Then
        vbaProject.VBComponents.Remove chkExist
    End If
    DoEvents
    On Error GoTo 0
    
    ' 2. ≈‰‘«¡ «·ÌÊ“— ›Ê—„
    Set vbaForm = vbaProject.VBComponents.Add(3)
    With vbaForm
        .Properties("Name") = "UserForm73"
        .Properties("Caption") = "‰Ÿ«„ «·—Ê« » «·„ÿÊ— - ⁄„«œ „Õ„œ «·‰⁄Ì„Ì"
        .Properties("Width") = 1250
        .Properties("Height") = 700
        .Properties("ScrollBars") = 2
        .Properties("KeepScrollBarsVisible") = 2
    End With

    BuildBaseLayout vbaForm
    InjectFinalLogic vbaForm

    MsgBox " „ «·≈’·«Õ «·‰Â«∆Ì. «·»ÕÀ Ê«·„Õ«ﬂ«… Ê«·ÿ»«⁄…  ⁄„· «·¬‰ »ﬂ›«¡….", vbInformation, "⁄„«œ „Õ„œ «·‰⁄Ì„Ì"
End Sub

Private Sub BuildBaseLayout(Frm As Object)
    Dim ctrl As Object
    
    ' Õﬁ· «·„Ì“«‰Ì…
    Set ctrl = Frm.Designer.Controls.Add("Forms.TextBox.1", "txtBudget")
    With ctrl: .Left = 450: .Top = 10: .Width = 180: .Height = 30: .value = "0": .Font.Size = 14: .Font.Bold = True: .TextAlign = 2: .BackColor = 14745568: End With

    ' “— «·„Õ«ﬂ«…
    Set ctrl = Frm.Designer.Controls.Add("Forms.CommandButton.1", "btnSimulate")
    With ctrl: .Left = 640: .Top = 10: .Width = 140: .Height = 30: .Caption = "»œ¡ «·„Õ«ﬂ«…": .Font.Bold = True: .BackColor = 12648384: End With

    ' “— «·ÿ»«⁄…
    Set ctrl = Frm.Designer.Controls.Add("Forms.CommandButton.1", "btnPrint")
    With ctrl: .Left = 790: .Top = 10: .Width = 140: .Height = 30: .Caption = "ÿ»«⁄…  ﬁ—Ì—": .Font.Bold = True: .BackColor = 16777164: End With

    ' «·»ÕÀ Ê«·ﬁ«∆„…
    Set ctrl = Frm.Designer.Controls.Add("Forms.TextBox.1", "txtSearch")
    With ctrl: .Left = 1000: .Top = 10: .Width = 220: .Height = 30: .Font.Name = "Times New Roman": .Font.Size = 12: .Font.Bold = True: .TextAlign = 2: End With
    
    Set ctrl = Frm.Designer.Controls.Add("Forms.ListBox.1", "ListBox1")
    With ctrl: .Left = 1000: .Top = 50: .Width = 220: .Height = 600: .ColumnCount = 2: .ColumnWidths = "60;140": .Font.Name = "Times New Roman": .Font.Size = 14: .Font.Bold = True: .TextAlign = 2: End With
End Sub

Private Sub InjectFinalLogic(Frm As Object)
    Dim c As String
    
    c = "Dim NextTop As Double" & vbCrLf
    c = c & "Dim GroupID As Long" & vbCrLf & vbCrLf
    
    c = c & "Private Sub UserForm_Initialize()" & vbCrLf
    c = c & "    NextTop = 60" & vbCrLf
    c = c & "    GroupID = 0" & vbCrLf
    c = c & "    Me.ScrollHeight = Me.Height" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf
    
    ' ﬂÊœ «·»ÕÀ
    c = c & "Private Sub txtSearch_Change()" & vbCrLf
    c = c & "    Dim ws1 As Worksheet: Set ws1 = Sheets(""Sheet1"")" & vbCrLf
    c = c & "    ListBox1.Clear" & vbCrLf
    c = c & "    Dim i As Long" & vbCrLf
    c = c & "    For i = 9 To ws1.Cells(ws1.Rows.Count, ""E"").End(xlUp).Row" & vbCrLf
    c = c & "        If InStr(1, ws1.Cells(i, ""E""), txtSearch.Text, 1) > 0 Or InStr(1, ws1.Cells(i, ""B""), txtSearch.Text, 1) > 0 Then" & vbCrLf
    c = c & "            ListBox1.AddItem ws1.Cells(i, ""B"")" & vbCrLf
    c = c & "            ListBox1.List(ListBox1.ListCount - 1, 1) = ws1.Cells(i, ""E"")" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next i" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' ﬂÊœ «·‰ﬁ— «·„“œÊÃ
    c = c & "Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)" & vbCrLf
    c = c & "    Dim ws1 As Worksheet: Set ws1 = Sheets(""Sheet1"")" & vbCrLf
    c = c & "    Dim ws9 As Worksheet: Set ws9 = Sheets(""Sheet9"")" & vbCrLf
    c = c & "    Dim r As Long, sVal As Variant" & vbCrLf
    c = c & "    If ListBox1.ListIndex = -1 Then Exit Sub" & vbCrLf
    c = c & "    sVal = ListBox1.Value" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    r = Application.Match(CLng(sVal), ws1.Range(""B:B""), 0)" & vbCrLf
    c = c & "    If Err.Number <> 0 Then" & vbCrLf
    c = c & "        Err.Clear" & vbCrLf
    c = c & "        r = Application.Match(CStr(sVal), ws1.Range(""B:B""), 0)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    On Error GoTo 0" & vbCrLf
    c = c & "    If r = 0 Then Exit Sub" & vbCrLf
    c = c & "    ws9.Range(""G18"") = ws1.Cells(r, ""H"")" & vbCrLf
    c = c & "    ws9.Range(""F18"") = ws1.Cells(r, ""I"") + 1" & vbCrLf
    c = c & "    GroupID = GroupID + 1" & vbCrLf
    c = c & "    CreateRow ListBox1.List(ListBox1.ListIndex, 1), CStr(sVal), CDbl(ws9.Range(""H18"").Value), r, GroupID" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' ﬂÊœ «·„Õ«ﬂ«…
    c = c & "Private Sub btnSimulate_Click()" & vbCrLf
    c = c & "    Dim ctrl As Control: Dim t As Double: t = 0" & vbCrLf
    c = c & "    For Each ctrl In Me.Controls" & vbCrLf
    c = c & "        If InStr(ctrl.Name, ""diff_"") > 0 Then t = t + Val(ctrl.Text)" & vbCrLf
    c = c & "    Next ctrl" & vbCrLf
    c = c & "    Dim res As Double: res = Val(txtBudget.Text) - t" & vbCrLf
    c = c & "    If res >= 0 Then" & vbCrLf
    c = c & "        MsgBox ""«·„Ì“«‰Ì…  ﬂ›Ì. «·›«∆÷: "" & res, vbInformation" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        MsgBox ""⁄Ã“ „«·Ì: "" & Abs(res), vbCritical" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' ﬂÊœ «·ÿ»«⁄…
    c = c & "Private Sub btnPrint_Click()" & vbCrLf
    c = c & "    Dim wsR As Worksheet: Set wsR = Worksheets.Add" & vbCrLf
    c = c & "    wsR.Range(""A1:C1"").Value = Array(""«·«”„"", ""«·«”„Ì"", ""«·“Ì«œ…"")" & vbCrLf
    c = c & "    Dim ctrl As Control: Dim rIdx As Long: rIdx = 2" & vbCrLf
    c = c & "    For Each ctrl In Me.Controls" & vbCrLf
    c = c & "        If InStr(ctrl.Name, ""lblNm_"") > 0 Then" & vbCrLf
    c = c & "            Dim gT As String: gT = ctrl.Tag" & vbCrLf
    c = c & "            wsR.Cells(rIdx, 1) = ctrl.Caption" & vbCrLf
    c = c & "            wsR.Cells(rIdx, 2) = Me.Controls(""box_"" & gT).Text" & vbCrLf
    c = c & "            wsR.Cells(rIdx, 3) = Me.Controls(""diff_"" & gT).Text" & vbCrLf
    c = c & "            rIdx = rIdx + 1" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next ctrl" & vbCrLf
    c = c & "    wsR.Columns.AutoFit: wsR.PrintPreview" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' ﬂÊœ «·Õ–›
    c = c & "Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)" & vbCrLf
    c = c & "    Dim ctrl As Control: For Each ctrl In Me.Controls" & vbCrLf
    c = c & "        If InStr(ctrl.Name, ""btnDel_"") > 0 Then" & vbCrLf
    c = c & "            If X >= ctrl.Left And X <= (ctrl.Left + ctrl.Width) And Y >= ctrl.Top And Y <= (ctrl.Top + ctrl.Height) Then" & vbCrLf
    c = c & "                Dim tG As String: tG = ctrl.Tag: Dim i As Integer" & vbCrLf
    c = c & "                For i = Me.Controls.Count - 1 To 0 Step -1" & vbCrLf
    c = c & "                    If Me.Controls(i).Tag = tG Then Me.Controls.Remove Me.Controls(i).Name" & vbCrLf
    c = c & "                Next i: Exit Sub" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next ctrl" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' œ«·… »‰«¡ «·”ÿ—
    c = c & "Sub CreateRow(Nm As String, ID As String, Nom As Double, rIdx As Long, gID As Long)" & vbCrLf
    c = c & "    Dim ws1 As Worksheet: Set ws1 = Sheets(""Sheet1""): Dim curR As Double: curR = 920" & vbCrLf
    c = c & "    With Me.Controls.Add(""Forms.Label.1"", ""btnDel_"" & gID): .Caption = "" [ ≈·€«¡ ] "": .Top = NextTop: .Left = 945: .Width = 45: .Tag = gID: .BackColor = &H8080FF: .ForeColor = &HFFFFFF: .TextAlign = 2: .Font.Bold = True: End With" & vbCrLf
    c = c & "    With Me.Controls.Add(""Forms.Label.1"", ""lblNm_"" & gID): .Caption = Nm: .Top = NextTop: .Left = 20: .Width = 920: .Tag = gID: .Font.Name = ""Times New Roman"": .Font.Size = 14: .Font.Bold = True: .ForeColor = 16711680: .TextAlign = 3: End With" & vbCrLf
    c = c & "    NextTop = NextTop + 30" & vbCrLf
    c = c & "    Dim rC, fC, j As Integer: rC = Array(""M"",""O"",""Q"",""Z"",""AB""): fC = Array(""S"",""T"",""U"",""Y"",""BH"",""BP"",""BR"")" & vbCrLf
    c = c & "    AddBox Nom, ws1.Range(""J8"").Value, curR, gID" & vbCrLf
    c = c & "    curR = curR - 68: Dim s As Double: s = Nom" & vbCrLf
    c = c & "    For j = 0 To 4: Dim v: v = Nom * ws1.Cells(rIdx, rC(j)): s = s + v: AddBox Format(v, ""0""), ws1.Cells(8, rC(j)).Value, curR, gID: curR = curR - 68: Next j" & vbCrLf
    c = c & "    For j = 0 To 6: Dim fV: fV = ws1.Cells(rIdx, fC(j)): s = s + fV: AddBox fV, ws1.Cells(8, fC(j)).Value, curR, gID: curR = curR - 68: Next j" & vbCrLf
    c = c & "    AddBox Format(s - ws1.Cells(rIdx, ""AE""), ""0""), ""«·“Ì«œ…"", curR, gID, True" & vbCrLf
    c = c & "    NextTop = NextTop + 90: If NextTop > Me.Height Then Me.ScrollHeight = NextTop + 50" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' œ«·… ≈÷«›… «·„—»⁄« 
    c = c & "Sub AddBox(v, t, R, gID, Optional isRes As Boolean = False)" & vbCrLf
    c = c & "    Dim tNm As String: If isRes Then tNm = ""diff_"" & gID Else tNm = ""box_"" & gID" & vbCrLf
    c = c & "    With Me.Controls.Add(""Forms.Label.1""): .Caption = t: .Top = NextTop: .Left = R: .Width = 65: .Tag = gID: .TextAlign = 2: .Font.Size = 8: .Font.Bold = True: End With" & vbCrLf
    c = c & "    With Me.Controls.Add(""Forms.TextBox.1"", tNm): .Text = v: .Top = NextTop + 18: .Left = R: .Width = 65: .Height = 28: .Tag = gID: .TextAlign = 2: .Font.Size = 12: .Font.Bold = True" & vbCrLf
    c = c & "    If isRes Then .BackColor = 9240494: .Width = 80: End With: End Sub"

    Frm.CodeModule.AddFromString c
End Sub
