VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm78 
   Caption         =   "UserForm78"
   ClientHeight    =   10776
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11256
   OleObjectBlob   =   "UserForm78.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm78"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Explicit

Dim ControlCount As Integer






Dim AllData As Variant

Private Sub CommandButton1_Click()

Dim i As Integer, j As Integer
    Dim found As Boolean
    
    ' «·œÊ—«‰ ⁄·Ï «·√”ÿ— «·„‰‘√…
    For i = 0 To ControlCount - 1
        ' «· √ﬂœ „‰ ÊÃÊœ «·ÃÌﬂ »Êﬂ” ÊﬁÌ„ Â True
        If Me.Controls("chkAbs" & i).value = True Then
            
            ' 1. Õ–› «·⁄‰«’— «·√—»⁄… «·„ﬂÊ‰… ··”ÿ—
            Me.Controls.Remove "lblJobID" & i
            Me.Controls.Remove "lblName" & i
            Me.Controls.Remove "txtAbs" & i
            Me.Controls.Remove "chkAbs" & i
            found = True
            
            ' 2. ≈⁄«œ…  — Ì» «·⁄‰«’— «· Ì  ·Ì «·”ÿ— «·„Õ–Ê› ( —ÕÌ· ··√⁄·Ï)
            For j = i + 1 To ControlCount - 1
                '  —ÕÌ· ·Ì»· «·—ﬁ„
                With Me.Controls("lblJobID" & j)
                    .Top = .Top - 45
                    .Name = "lblJobID" & (j - 1)
                End With
                
                '  —ÕÌ· ·Ì»· «·«”„
                With Me.Controls("lblName" & j)
                    .Top = .Top - 45
                    .Name = "lblName" & (j - 1)
                End With
                
                '  —ÕÌ· «· Ìﬂ”  »Êﬂ”
                With Me.Controls("txtAbs" & j)
                    .Top = .Top - 45
                    .Name = "txtAbs" & (j - 1)
                End With
                
                '  —ÕÌ· «·ÃÌﬂ »Êﬂ”
                With Me.Controls("chkAbs" & j)
                    .Top = .Top - 45
                    .Name = "chkAbs" & (j - 1)
                End With
            Next j
            
            ' 3.  ÕœÌÀ «·⁄œ«œ «·⁄«„
            ControlCount = ControlCount - 1
            Exit For ' ‰Œ—Ã »⁄œ Õ–› √Ê· ”ÿ— „ƒ‘— ⁄·ÌÂ (·÷„«‰ ”·«„… «· — Ì»)
        End If
    Next i
    
    '  ‰»ÌÂ ≈–« ·„ Ì „ «Œ Ì«— √Ì ”ÿ—
    If Not found Then
        MsgBox "Ì—ÃÏ  √‘Ì— «·ÃÌﬂ »Êﬂ” ··”ÿ— «·„—«œ Õ–›Â", vbExclamation, " ‰»ÌÂ"
    Else
        '  ÕœÌÀ „‰ÿﬁ… «· „—Ì— »⁄œ «·Õ–›
        Me.ScrollHeight = 42 + (ControlCount * 45) + 50
    End If

















End Sub

Private Sub CommandButton2_Click()

Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Integer, r As Long, lastRow As Long
    Dim targetID As String, TargetName As String
    Dim absDays As Double
    Dim salaryVal As Double ' «·ﬁÌ„… «·„ÊÃÊœ… ›Ì «·⁄„Êœ AE
    
    ' 1.  ÕœÌœ ¬Œ— ’› ›Ì «·‘Ì  »‰«¡ ⁄·Ï «·⁄„Êœ B
    lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
    
    ' √„«‰: «· √ﬂœ „‰ √‰ «·ÃœÊ· ÌÕ ÊÌ »Ì«‰«  „‰ «·’› 9 ›„« ›Êﬁ
    If lastRow < 9 Then
        MsgBox "«·ÃœÊ· ›Ì «·‘Ì  ›«—€ (Ì»œ√ „‰ «·’› 9)", vbCritical
        Exit Sub
    End If
    
    ' 2. «·œÊ—«‰ ⁄·Ï Ã„Ì⁄ «·√”ÿ— «·„‰‘√… »—„ÃÌ« ›Ì «·ÌÊ“— ›Ê—„
    For i = 0 To ControlCount - 1
        ' Ã·» «·»Ì«‰«  „‰ «··Ì»·«  Ê«· Ìﬂ”  »Êﬂ” „»«‘—…
        targetID = Trim(Me.Controls("lblJobID" & i).Caption)
        TargetName = Trim(Me.Controls("lblName" & i).Caption)
        absDays = val(Me.Controls("txtAbs" & i).value)
        
        ' 3. «·Õ·ﬁ… «· ﬂ—«—Ì… ··»ÕÀ œ«Œ· «·‘Ì   »œ√ „‰ «·’› 9 Õ’—«
        For r = 9 To lastRow
            ' «·‘—ÿ «·„“œÊÃ ·· ⁄‘Ìﬁ «·ÃÌ‰Ì ( ÿ«»ﬁ «·—ﬁ„ ›Ì B Ê«·«”„ ›Ì E)
            If Trim(ws.Cells(r, "B").value) = targetID And _
               Trim(ws.Cells(r, "E").value) = TargetName Then
                
                ' 4.  —ÕÌ· √Ì«„ «·€Ì«» (K) Ê«·„ »ﬁÌ (L)
                ws.Cells(r, "K").value = absDays
                ws.Cells(r, "L").value = 30 - absDays
                 ws.Cells(r, "K").value = absDays
                   ' 5. «Õ ”«» «·ﬁÌ„… ›Ì AU ﬂ‹ Integer »‰«¡ ⁄·Ï «·⁄„Êœ AE
                ' ”Õ» «·ﬁÌ„… „‰ «·⁄„Êœ AE («·⁄„Êœ —ﬁ„ 31)
                salaryVal = val(ws.Cells(r, "AE").value)
                
                If salaryVal > 0 Then
                    ' «·„⁄«œ·…: (ﬁÌ„… AE  ﬁ”Ì„ 30) „÷—Ê»… ›Ì √Ì«„ «·€Ì«»
                    ' «” Œœ«„ Int ·ﬁÿ⁄ «·ﬂ”Ê— ‰Â«∆Ì«
                    ws.Cells(r, "AU").value = Int((salaryVal / 30) * absDays)
                       ws.Cells(r, "bu").value = "€Ì«»"
            
                Else
                    ws.Cells(r, "AU").value = 0
                End If
                
                Exit For ' «·«‰ ﬁ«· ··„ÊŸ› «· «·Ì ›Ê— ≈ÌÃ«œ «·„ÿ«»ﬁ…
            End If
        Next r
    Next i
    
    ' 6. —”«·… «·‰Ã«Õ «·‰Â«∆Ì…
    MsgBox " „  —ÕÌ· «·»Ì«‰«  Ê«Õ ”«» „»·€ «·€Ì«» »‰Ã«Õ.", vbInformation

Call count_all

End Sub

Private Sub CommandButton3_Click()
Unload Me
UserForm34.Show


End Sub

Private Sub UserForm_Initialize()
  Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim lastRow As Long, i As Long
    
    lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
    
    ' ÷»ÿ ≈⁄œ«œ«  «··”  »Êﬂ”
    With ListBox1
        .ColumnCount = 2          ' ⁄„ÊœÌ‰: ··—ﬁ„ Ê«·«”„
        .ColumnWidths = "0;150"   ' ≈Œ›«¡ «·⁄„Êœ «·√Ê· («·—ﬁ„) Ê≈ŸÂ«— «·À«‰Ì («·«”„)
        
        '  ⁄»∆… «··”  »Êﬂ” „‰ «·‘Ì  «» œ«¡ „‰ «·’› 9
        For i = 9 To lastRow
            If ws.Cells(i, "B").value <> "" Then ' «· √ﬂœ „‰ ÊÃÊœ —ﬁ„ ÊŸÌ›Ì
                .AddItem ws.Cells(i, "B").value & ""  ' ≈÷«›… «·—ﬁ„ ›Ì «·⁄„Êœ 0
                .List(.ListCount - 1, 1) = ws.Cells(i, "E").value & "" ' ≈÷«›… «·«”„ ›Ì «·⁄„Êœ 1
            End If
        Next i
    End With
    
    ControlCount = 0 '  ’›Ì— «·⁄œ«œ ⁄‰œ › Õ «·›Ê—„
End Sub

Private Sub TextBox1_Change()
    ' ›· —… ›Ê—Ì… ⁄‰œ «·ﬂ «»…
    Call RefreshListBox(TextBox1.Text)
End Sub

Sub RefreshListBox(searchTerm As String)
    Dim i As Long
    Dim strName As String
    
    ListBox1.Clear
    
    If IsEmpty(AllData) Then Exit Sub
    
    ' œÊ—«‰ ·›· —… «·√”„«¡
    For i = 1 To UBound(AllData, 1)
        strName = CStr(AllData(i, 1))
        
        ' «·»ÕÀ ⁄‰ «·‰’
        If InStr(1, strName, searchTerm, vbTextCompare) > 0 Or searchTerm = "" Then
            ' ≈÷«›… «·«”„ „»«‘—… »œÊ‰ „”«›«  ≈÷«›Ì… ·÷„«‰ À»«  «·”ﬂ—Ê· »«—
            ListBox1.AddItem strName
        End If
    Next i
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)


Dim lblID As MSForms.Label
    Dim lblName As MSForms.Label
    Dim txtAbs As MSForms.TextBox
    Dim chkBox As MSForms.CheckBox ' ≈÷«›… «·ÃÌﬂ »Êﬂ”
    Dim topPosition As Single
    
    ' 1. «· Õﬁﬁ „‰ ÊÃÊœ «Œ Ì«— ›Ì «·ﬁ«∆„…
    If ListBox1.ListIndex = -1 Then Exit Sub
    
    ' 2. Õ”«» „Êﬁ⁄ «·ŸÂÊ— («·»œ«Ì… „‰ 42 Ê«·„”«›… 45)
    topPosition = 42 + (ControlCount * 45)
    
    ' 3. ≈‰‘«¡ ·Ì»· «·—ﬁ„ «·ÊŸÌ›Ì (B)
    Set lblID = Me.Controls.Add("Forms.Label.1", "lblJobID" & ControlCount, True)
    With lblID
        .Caption = ListBox1.Column(0) & ""
        .Left = 10: .Top = topPosition: .Width = 50: .Height = 20
        .Font.Name = "Times New Roman": .Font.Size = 10
        .BackColor = &H80FFFF: .TextAlign = 2: .BorderStyle = 1
    End With
    
    ' 4. ≈‰‘«¡ ·Ì»· «”„ «·„ÊŸ› (E)
    Set lblName = Me.Controls.Add("Forms.Label.1", "lblName" & ControlCount, True)
    With lblName
        .Caption = ListBox1.Column(1) & ""
        .Left = 65: .Top = topPosition: .Width = 150: .Height = 20
        .Font.Name = "Times New Roman": .Font.Size = 11: .Font.Bold = True
    End With
    
    ' 5. ≈‰‘«¡ «· Ìﬂ”  »Êﬂ” (≈œŒ«· «·€Ì«»)
    Set txtAbs = Me.Controls.Add("Forms.TextBox.1", "txtAbs" & ControlCount, True)
    With txtAbs
        .Left = 220: .Top = topPosition - 2: .Width = 45: .Height = 22
        .Font.Name = "Times New Roman": .Font.Size = 12: .Font.Bold = True
        .value = 0: .TextAlign = 2
        .Tag = lblID.Caption & "|" & lblName.Caption ' «· ⁄‘Ìﬁ «·ÃÌ‰Ì
    End With
    
    ' 6. ≈‰‘«¡ «·ÃÌﬂ »Êﬂ” (CheckBox) ›Ì ‰Â«Ì… «·”ÿ—
    Set chkBox = Me.Controls.Add("Forms.CheckBox.1", "chkAbs" & ControlCount, True)
    With chkBox
        .Left = 275 ' „Êﬁ⁄ ›Ì ‰Â«Ì… «·”ÿ— »⁄œ «· Ìﬂ”  »Êﬂ”
        .Top = topPosition
        .Width = 20: .Height = 20
        .Caption = "" ' »œÊ‰ ‰’  Ê÷ÌÕÌ ·ÌŸ· ’€Ì—«
        .value = False ' «·Õ«·… «·«› —«÷Ì… €Ì— „Õœœ
    End With
    
    ' 7.  ÕœÌÀ «·⁄œ«œ Ê„‰ÿﬁ… «· „—Ì—
    ControlCount = ControlCount + 1
    Me.ScrollBars = 2 ' fmScrollBarsVertical
    Me.ScrollHeight = topPosition + 80
    
    ' «· —ﬂÌ“ ⁄·Ï «· Ìﬂ”  »Êﬂ” ›Ê—«
    txtAbs.SetFocus
   
End Sub
