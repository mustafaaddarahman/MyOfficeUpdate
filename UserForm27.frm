VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm27 
   Caption         =   "UserForm27"
   ClientHeight    =   9108.001
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12864
   OleObjectBlob   =   "UserForm27.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


















Private Sub CheckBox1_Click()
'If CheckBox1.Value = True Then
'CheckBox2.Value = False
'CommandButton1.Enabled = True
'ComboBox1.DropDown
'ComboBox1.Value = "Õ–›"
'TextBox7.Text = "Õ–›"
'End If
End Sub

Private Sub CheckBox2_Click()
'If CheckBox2.Value = True Then
'CheckBox1.Value = False
'CommandButton1.Enabled = True
'ComboBox1.DropDown
'ComboBox1.Value = "«÷«›…"
'TextBox7.Text = "«÷«›…"
'End If
End Sub

Private Sub CheckBox3_Click()
If CheckBox3.value = True Then
CommandButton4.Enabled = True
Else
CommandButton4.Enabled = False
End If
End Sub

Private Sub CheckBox4_Click()
    If Me.CheckBox4.value = True Then
        Me.CheckBox4.Caption = "≈ŸÂ«— Ã„Ì⁄ «·„ €Ì—« "
    
    Call RefreshListBox
    
    Else
     Me.CheckBox4.Caption = "≈ŸÂ«— «·»Ì«‰«  Õ”» «·«”„ "
   
    Dim ws As Worksheet
    Dim i3 As Long, lastRow As Long
    Dim isFound As Boolean
    
    Set ws = ThisWorkbook.Sheets(35)
    isFound = False
    
    With Me.ListBox1
        .ColumnCount = 10 ' “œ‰« «·√⁄„œ… ≈·Ï 10
        .ColumnWidths = "120;120;170;100;180;150;180;60;50;0" ' «·√ŒÌ— „Œ›Ì
        .Clear
        
        ' ≈÷«›… —ƒÊ” «·√⁄„œ…
        .AddItem
        .List(0, 8) = CenterText(ws.Cells(8, "A").value)
        .List(0, 7) = CenterText(ws.Cells(8, "B").value)
        .List(0, 6) = CenterText(ws.Cells(8, "E").value)
        .List(0, 5) = CenterText(ws.Cells(8, "AE").value)
        .List(0, 4) = CenterText(ws.Cells(8, "BZ").value)
        .List(0, 3) = CenterText(ws.Cells(8, "CC").value)
        .List(0, 2) = CenterText(ws.Cells(8, "CD").value)
        .List(0, 1) = CenterText(ws.Cells(8, "CE").value)
        .List(0, 0) = CenterText(ws.Cells(8, "CF").value)
    End With
    
    lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
    
    For i3 = 9 To lastRow
        If Trim(CStr(ws.Cells(i3, "B").value)) = Trim(UserForm1.TextBox2.value) And _
           Trim(CStr(ws.Cells(i3, "E").value)) = Trim(UserForm1.TextBox5.value) Then
            
            isFound = True
            With Me.ListBox1
                .AddItem
                .List(.ListCount - 1, 8) = ws.Cells(i3, "A").value
                .List(.ListCount - 1, 7) = ws.Cells(i3, "B").value
                .List(.ListCount - 1, 6) = ws.Cells(i3, "E").value
                .List(.ListCount - 1, 5) = ws.Cells(i3, "AE").value
                .List(.ListCount - 1, 4) = ws.Cells(i3, "BZ").value
                .List(.ListCount - 1, 3) = ws.Cells(i3, "CC").value
                .List(.ListCount - 1, 2) = ws.Cells(i3, "CD").value
                .List(.ListCount - 1, 1) = ws.Cells(i3, "CE").value
                .List(.ListCount - 1, 0) = ws.Cells(i3, "CF").value
                '  Œ“Ì‰ —ﬁ„ «·’› «·ÕﬁÌﬁÌ ›Ì «·⁄„Êœ «·„Œ›Ì (index 9)
                .List(.ListCount - 1, 9) = i3
            End With
        End If
    Next i3
     End If
End Sub

Private Sub ComboBox1_AfterUpdate()
TextBox7.Text = ComboBox1.value
TextBox1.Text = ""
If TextBox7.Text Like "«÷«›…" & "*" Then
ComboBox2.Enabled = True
ComboBox3.Enabled = False

Else
ComboBox2.Enabled = False
ComboBox3.Enabled = True
End If
End Sub

Private Sub ComboBox1_Change()
TextBox7.Text = ComboBox1.value
TextBox1.Text = ""
If TextBox7.Text Like "«÷«›…" & "*" Then
ComboBox3.Enabled = False
Else
ComboBox2.Enabled = False
ComboBox3.Enabled = True
End If
End Sub

Private Sub ComboBox2_AfterUpdate()
 
 ' ComboBox3.Value = ""
  ' ComboBox4.Value = ""
 TextBox1.Text = ""
' TextBox1.Text = ComboBox2.Value & " " & ComboBox3.Value & " " & "·‘Â—" & " " & ComboBox4.Value
End Sub

Private Sub ComboBox2_Change()
TextBox1.Text = ""
 TextBox1.Text = ComboBox2.value
  
TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & ComboBox4.value
End Sub

Private Sub ComboBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub ComboBox3_AfterUpdate()
 
' ComboBox2.Value = ""
'ComboBox4.Value = ""
TextBox1.Text = ""

'TextBox1.Text = ComboBox2.Value & " " & ComboBox3.Value & " " & "·‘Â—" & " " & " " & ComboBox4.Value
End Sub

Private Sub ComboBox3_Change()
TextBox1.Text = ""

TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & " " & ComboBox4.value
End Sub

Private Sub ComboBox4_AfterUpdate()

 TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & ComboBox4.value
End Sub

Private Sub ComboBox4_Change()
TextBox1.Text = ""
TextBox1.Text = ComboBox2.value & " " & ComboBox3.value & " " & "·‘Â—" & " " & ComboBox4.value
End Sub

Private Sub CommandButton1_Click()
'Call Remove_Duplicates1
'Call delete_change_month1
Call change_salary
MsgBox " „  ‰›Ì– »Ì«‰«  «·ﬁÌœ «·ﬁœÌ„ ·—« » «·‘Â— «·„«÷Ì", vbDefaultButton1, "—”«·…  ÊÃÌÂ"
CommandButton2.Enabled = True
CommandButton1.Enabled = False
'Dim LastRow As Long
       
  '  LastRow = Sheets(35).Cells(Rows.Count, 1).End(xlUp).Row

'TextBox9.Value = Sheets(35).Cells(LastRow, 1).Value
Dim lastr110, n As Integer
lastr110 = Sheets(35).Cells(rowS.count, "a").End(xlUp).row

For n = 9 To lastr110
TextBox9.Text = Sheets(35).Cells(n, 84).value

Next

End Sub

Private Sub CommandButton2_Click()

Call update_change_salary


Sheets(55).Range("a9:bz8000").ClearContents
Application.ScreenUpdating = False
Sheets(55).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select
Selection.Copy Sheets(55).Range("A1:bz8000")

Application.ScreenUpdating = True
[a8].Select
Sheets(55).Select
Sheets(55).Activate
Application.ScreenUpdating = True


CommandButton2.Enabled = False
















End Sub



Private Sub CommandButton3_Click()
'=VLOOKUP(AL8:AM8;A8:AI13;35;0)
'Ê—ﬁ…33.Range("am8").Formula = "= MAXIFS(AC9:AC13,A9:A13,AL8,B9:B13,AK8)"
'Ê—ﬁ…33.Range("an8").Formula2R1C1 = "= VLOOKUP(RC[-2]:RC[-1],RC[-39]:R[5]C[-5],35,0)"
'Ê—ﬁ…33.Cells(8, "ak") = UserForm1.TextBox5.Text         ' «·«”„
'Ê—ﬁ…33.Cells(8, "al") = UserForm1.TextBox2.Text  '«·—ﬁ„ «·ÊŸÌ›Ì
Unload Me
End Sub

Private Sub CommandButton4_Click()
Dim rowToDelete As Long
    
    If Me.ListBox1.ListIndex <= 0 Then
        MsgBox "Ì—ÃÏ «Œ Ì«— ”Ã· „‰ «·ﬁ«∆„… ··Õ–›", vbExclamation
        Exit Sub
    End If
    
    ' Ã·» —ﬁ„ «·’› «·„Œ“‰ ›Ì «·⁄„Êœ «·⁄«‘— (index 9)
    rowToDelete = Me.ListBox1.List(Me.ListBox1.ListIndex, 9)
    
    If MsgBox("Â· √‰  „ √ﬂœ „‰ Õ–› Â–« «·”Ã·ø", vbYesNo + vbQuestion) = vbYes Then
        ThisWorkbook.Sheets(35).rowS(rowToDelete).Delete
        Call arb '  ÕœÌÀ «·ﬁ«∆„… »⁄œ «·Õ–›
        MsgBox " „ «·Õ–› »‰Ã«Õ", vbInformation
    End If

End Sub

Private Sub CommandButton5_Click()
 Dim lastRow As Long
    Dim i As Long
    
    
    lastRow = Sheets(35).Cells(rowS.count, 6).End(xlUp).row
    
   
    For i = 9 To lastRow
     If Sheets(35).Cells(i, 2).value = "" And Sheets(35).Cells(i, 5).value = "" And Sheets(35).Cells(i, 6).value = UserForm1.ComboBox5.value And Sheets(35).Cells(i, 7).value = UserForm1.TextBox7.value Then
           Sheets(35).rowS(i).Delete
        End If
    Next i
    MsgBox "ﬁœ  „ Õ–› «·ﬁÌœ »‰Ã«Õ Õ«Ê· «‰ ·«  ‰”Ï «·«Ã—«¡ «·’ÕÌÕ ÂÊ ﬁ„ »«· ‰›Ì– À„ «· ÕœÌÀ ‘ﬂ—« ·ﬂ", vbDefaultButton1, "—”«·…  ‰»ÌÂ"
End Sub







Private Sub CommandButton6_Click()
Unload Me
UserForm71.Show
End Sub

Private Sub CommandButton7_Click()
Unload Me
UserForm80.Show
End Sub

Private Sub ListBox1_Click()
Dim valCC As Double
    Dim rowIndex As Long
    
    ' 1.  ÕœÌœ «·”Ã· «·„Œ «—
    rowIndex = Me.ListBox1.ListIndex
    
    ' 2. Õ„«Ì…: ≈–« ·„ Ì „ «Œ Ì«— ‘Ì¡ √Ê  „ «Œ Ì«— «·—√” («·’› 0)° «Œ—Ã „‰ «·ﬂÊœ
    If rowIndex <= 0 Then Exit Sub
    
    ' 3. „⁄«·Ã… «·⁄„Êœ CC (·· Ìﬂ”  »Êﬂ”)
    ' «·⁄„Êœ CC ÂÊ «·›Â—” —ﬁ„ 3 (Õ”»  — Ì»ﬂ «·”«»ﬁ)
    If IsNumeric(Me.ListBox1.List(rowIndex, 3)) Then
        valCC = CDbl(Me.ListBox1.List(rowIndex, 3))
        
        If valCC < 0 Then
            Me.TextBox5.value = valCC
            Me.TextBox6.value = ""
             Me.Label13.Caption = Me.ListBox1.List(rowIndex, 1)
              Me.TextBox3.value = Me.ListBox1.List(rowIndex, 5)
               Me.TextBox4.value = Me.ListBox1.List(rowIndex, 4)
        Else
            Me.TextBox6.value = valCC
            Me.TextBox5.value = ""
             Me.Label13.Caption = Me.ListBox1.List(rowIndex, 1)
              Me.TextBox3.value = Me.ListBox1.List(rowIndex, 5)
               Me.TextBox4.value = Me.ListBox1.List(rowIndex, 4)
        End If
    Else
        ' ≈–« ﬂ«‰  «·ﬁÌ„… ·Ì”  —ﬁ„« (‰’)° ›—€ «· Ìﬂ”  »Êﬂ”
        Me.TextBox5.value = ""
        Me.TextBox6.value = ""
    End If
    
    
End Sub

Private Sub TextBox1_Change()
Call arb

End Sub

Private Sub TextBox7_Change()
ComboBox2.value = ""
ComboBox3.value = ""
If TextBox7.Text Like "«÷«›…" & "*" Then
ComboBox2.Enabled = True
ComboBox3.Enabled = False
TextBox1.Text = ""
Else
ComboBox2.Enabled = False
ComboBox3.Enabled = True
TextBox1.Text = ""
End If
End Sub

Private Sub UserForm_Activate()

Dim ws As Worksheet
    Dim i3 As Long, lastRow As Long
    Dim isFound As Boolean
    
    Set ws = ThisWorkbook.Sheets(35)
    isFound = False
    
    With Me.ListBox1
        .ColumnCount = 10 ' “œ‰« «·√⁄„œ… ≈·Ï 10
        .ColumnWidths = "120;120;170;100;180;150;180;60;50;0" ' «·√ŒÌ— „Œ›Ì
        .Clear
        
        ' ≈÷«›… —ƒÊ” «·√⁄„œ…
        .AddItem
        .List(0, 8) = CenterText(ws.Cells(8, "A").value)
        .List(0, 7) = CenterText(ws.Cells(8, "B").value)
        .List(0, 6) = CenterText(ws.Cells(8, "E").value)
        .List(0, 5) = CenterText(ws.Cells(8, "AE").value)
        .List(0, 4) = CenterText(ws.Cells(8, "BZ").value)
        .List(0, 3) = CenterText(ws.Cells(8, "CC").value)
        .List(0, 2) = CenterText(ws.Cells(8, "CD").value)
        .List(0, 1) = CenterText(ws.Cells(8, "CE").value)
        .List(0, 0) = CenterText(ws.Cells(8, "CF").value)
    End With
    
    lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
    
    For i3 = 9 To lastRow
        If Trim(CStr(ws.Cells(i3, "B").value)) = Trim(UserForm1.TextBox2.value) And _
           Trim(CStr(ws.Cells(i3, "E").value)) = Trim(UserForm1.TextBox5.value) Then
            
            isFound = True
            With Me.ListBox1
                .AddItem
                .List(.ListCount - 1, 8) = ws.Cells(i3, "A").value
                .List(.ListCount - 1, 7) = ws.Cells(i3, "B").value
                .List(.ListCount - 1, 6) = ws.Cells(i3, "E").value
                .List(.ListCount - 1, 5) = ws.Cells(i3, "AE").value
                .List(.ListCount - 1, 4) = ws.Cells(i3, "BZ").value
                .List(.ListCount - 1, 3) = ws.Cells(i3, "CC").value
                .List(.ListCount - 1, 2) = ws.Cells(i3, "CD").value
                .List(.ListCount - 1, 1) = ws.Cells(i3, "CE").value
                .List(.ListCount - 1, 0) = ws.Cells(i3, "CF").value
                '  Œ“Ì‰ —ﬁ„ «·’› «·ÕﬁÌﬁÌ ›Ì «·⁄„Êœ «·„Œ›Ì (index 9)
                .List(.ListCount - 1, 9) = i3
            End With
        End If
    Next i3
End Sub

' œ«·… „”«⁄œ… · Ê”Ìÿ «·‰’
Function CenterText(val As Variant) As String
    Dim s As String
    s = CStr(val)
    ' Ì„ﬂ‰ﬂ  €ÌÌ— «·—ﬁ„ 10 ·“Ì«œ… √Ê  ﬁ·Ì· «·„”«›«  «·»«œ∆…
    CenterText = Space(5) & s
















'=========================================================

CheckBox3.value = False
CommandButton2.Enabled = False


'================================================
Dim lastr100, i1 As Integer
lastr100 = Sheets(35).Cells(rowS.count, "f").End(xlUp).row
    For i1 = 9 To lastr100
If Sheets(35).Cells(i1, 2).value = "" And Sheets(35).Cells(i1, 5).value = "" And Sheets(35).Cells(i1, 6).value = UserForm1.ComboBox5.value And Sheets(35).Cells(i1, 7).value = UserForm1.TextBox7.value Then
TextBox8.Text = "Â‰«ﬂ Œÿ√ ›Ì «· ‰›Ì– ﬁœ ﬁ„  » ÕœÌÀ «·»Ì«‰«   ﬁ»· «‰  ﬁÊ„ » ‰›Ì– »Ì«‰«  «·‘Â— «·„«÷Ì ﬁ„ »Õ–› «·ﬁÌœ Ê«⁄«œ…  —ÕÌ·Â „‰ ÃœÌœ"
CommandButton5.Visible = True
ElseIf Sheets(35).Cells(i1, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(i1, 5).value = UserForm1.TextBox5.Text Then

TextBox8.Text = "«·ﬁÌœ „ÕœÀ Ê „  —ÕÌ·Â »‰Ã«Õ"
CommandButton5.Visible = False
End If
Next











'===========================================================

ComboBox2.Enabled = False
ComboBox3.Enabled = False

ComboBox1.value = ""
ComboBox2.value = ""
ComboBox3.value = ""
ComboBox4.value = ""

ComboBox1.List = Array("«÷«›…", "Õ–›")
ComboBox2.List = Array("«÷«›… ⁄·«ÊÂ ÃœÌœÂ", "«÷«›… œ—Ã… Ê —›Ì⁄", " €ÌÌ— ‰”»… ‘Â«œ…", " €Ì— ‰”»… Â‰œ”Ì…", "«÷«›… „‰’»", "«÷«›… „Êﬁ⁄ Ã€—«›Ì", "«÷«›… „Œ’’«  «ÿ›«·", _
"«÷«›… „Œ’’«  “ÊÃÌÂ ", "«÷«›… „Œ’’«  “ÊÃÌÂ „⁄ «ÿ›«·", "«÷«›… „Œ’’«  „Â‰Ì…", "«÷«›… „Œ’’«  ŒÿÊ—Â ··”Ê«ﬁ", "«÷«›… „Œ’’«  ŒÿÊ—… Õ«”»« ")
ComboBox3.List = Array("Õ–› „Œ’’«  „‰’»", "Œ–› „Œ’’«  „Êﬁ⁄ Ã€—«›Ì", "Õ–› „Œ’’«  «ÿ›«·", "Õ–› „Œ’’«  “ÊÃÌ…", "Õ–› „Œ’’«  «ÿ›«· „⁄ “ÊÃÌ…", "ÕÃ» „Œ’’«  „Â‰ÌÂ", _
"ÕÃ» „Œ’’«  ŒÿÊ—Â ··”Ê«ﬁ", "ÕÃ» „Œ’’«  ŒÿÊ—Â Õ«”»«  ‰”»…")
Dim nammonth1 As Date
Dim i As Date



For i = 1 To 12
nammonth = monthName(i, abbreviate:=False)
ComboBox4.AddItem nammonth
Next
Label1.Caption = "Õœœ ”»» «· ⁄œÌ· «·–Ì „‰ «Ã·Â  „ «· €ÌÌ— »Ì«‰«  «·„ÊŸ›" & ":" & UserForm1.TextBox5.Text
'==========================================

On Error Resume Next

    
   
  '==========================
  
  '= VLOOKUP(AL8:AM8,$A$8:$AI$300000,35,0)
'Ê—ﬁ…33.Range("am7").Formula = "= MAXIFS($BN$9:$BN$300000,$B$9:$B$300000,AL7;$E$9:$E$300000,AK7)"
'Ê—ﬁ…33.Range("an7").Formula2R1C1 = "= VLOOKUP(RC[-2]:RC[-1],R[1]C[-39]:R[4000]C[43],82,1)"
'Ê—ﬁ…33.Cells(7, "ak") = UserForm1.TextBox5.Text         ' «·«”„
'Ê—ﬁ…33.Cells(7, "al") = UserForm1.TextBox2.Text  '«·—ﬁ„ «·ÊŸÌ›Ì
'======================================================
Dim lastr001 As Integer
Dim y2, y3, y4 As Integer
lastr001 = Sheets(35).Cells(rowS.count, 1).End(xlUp).row
    For y2 = 9 To lastr001

 If Sheets(35).Cells(y2, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(y2, 5).value = UserForm1.TextBox5.Text And Sheets(35).Cells(y2, "ce").value = "Õ–›" Then
TextBox2.Text = "Â–« «Œ—  «—ÌŒ  „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y2, "bn").value & " " & "Ê „  ÕœÌÀ »Ì«‰« Â" & "   " & Sheets(35).Cells(y2, 82) & "  " & "Ê«·„ €Ì—«  ›Ì Õ«·…" & " " & Sheets(35).Cells(y2, "ce")
 TextBox3.value = Sheets(35).Cells(y2, 31).value
  TextBox4.value = Sheets(35).Cells(y2, 78).value
 TextBox5.value = Sheets(35).Cells(y2, 81).value
TextBox6.value = ""
TextBox7.value = Sheets(35).Cells(y2, 83).value

End If
Next
'===================================================
 For y3 = 9 To lastr001
If Sheets(35).Cells(y3, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(y3, 5).value = UserForm1.TextBox5.Text And Sheets(35).Cells(y3, "ce").value = "«÷«›…" Then
TextBox2.Text = "Â–« «Œ—  «—ÌŒ  „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y3, "bn").value & "Ê „  ÕœÌÀ »Ì«‰« Â" & "  " & Sheets(35).Cells(y3, 82) & "Ê«·„ €Ì—«  ›Ì Õ«·…" & Sheets(35).Cells(y3, "ce")
TextBox3.value = Sheets(35).Cells(y3, 31).value
  TextBox4.value = Sheets(35).Cells(y3, 78).value
 TextBox5.value = ""
TextBox6.value = Sheets(35).Cells(y3, 81).value
TextBox7.value = Sheets(35).Cells(y3, 83).value

End If
Next
'===============================
For y4 = 9 To lastr2
If Sheets(35).Cells(y4, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(y4, 5).value = UserForm1.TextBox5.Text _
And Sheets(35).Cells(y4, "ce").value <> "«÷«›…" And Sheets(35).Cells(y4, "ce").value <> "Õ–›" Then
UserForm27.TextBox2.Text = "·«   Ê›— ·œÌÂ «Ì „ €Ì—«  :"
    End If
Next
  
  
  '=================================



If TextBox3.value = "" And TextBox4.value = "" Then

CommandButton4.Enabled = False
CommandButton1.Enabled = True
Else
CommandButton4.Enabled = True
CommandButton1.Enabled = False
End If

If TextBox3.value > 1 And TextBox4.value = "" Then

CommandButton4.Enabled = False
End If
'===================================

If TextBox7.Text = "«÷«›…" Then
CheckBox2.value = True
CheckBox1.value = False
CheckBox1.Enabled = False

If CheckBox2.value = True Then
CheckBox1.value = False
CommandButton1.Enabled = True
ComboBox1.DropDown
ComboBox1.value = "Õ–›"
TextBox7.Text = "Õ–›"
End If





ElseIf TextBox7.Text = "Õ–›" Then
CheckBox1.value = True
CheckBox2.value = False
CheckBox2.Enabled = False
If CheckBox1.value = True Then
CheckBox2.value = False
CommandButton1.Enabled = True
ComboBox1.DropDown
ComboBox1.value = "«÷«›…"
TextBox7.Text = "«÷«›…"
End If




ElseIf TextBox7.Text = "" Then
CheckBox1.value = False
CheckBox2.value = False
CheckBox1.Enabled = False
CheckBox2.Enabled = False
End If
End Function










Private Sub UserForm_Click()

End Sub

Sub RefreshListBox()
   Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(35)
    Dim i3 As Long, lastRow As Long
    Dim isFound As Boolean: isFound = False
    
    ' 1.  ÂÌ∆… «·‹ ListBox („”Õ «·ﬁœÌ„ Ê≈⁄œ«œ «·√⁄„œ…)
    With Me.ListBox1
        .ColumnCount = 10
        .ColumnWidths = "120;120;170;100;180;150;180;60;50;0"
        .Clear
        ' ≈÷«›… —ƒÊ” «·√⁄„œ… (Ì „ «” œ⁄«¡ «·œ«·… CenterText «·Œ«’… »ﬂ)
        .AddItem
        .List(0, 8) = CenterText(ws.Cells(8, "A").value)
        .List(0, 7) = CenterText(ws.Cells(8, "B").value)
        .List(0, 6) = CenterText(ws.Cells(8, "E").value)
        .List(0, 5) = CenterText(ws.Cells(8, "AE").value)
        .List(0, 4) = CenterText(ws.Cells(8, "BZ").value)
        .List(0, 3) = CenterText(ws.Cells(8, "CC").value)
        .List(0, 2) = CenterText(ws.Cells(8, "CD").value)
        .List(0, 1) = CenterText(ws.Cells(8, "CE").value)
        .List(0, 0) = CenterText(ws.Cells(8, "CF").value)
    End With
    
    lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
    
    ' 2. Õ·ﬁ… «·»ÕÀ «·–ﬂÌ…
    For i3 = 9 To lastRow
        ' Â‰« «·–ﬂ«¡: ≈–« ﬂ«‰ CheckBox4=True Ì⁄—÷ «·ﬂ·° ≈–« False Ì⁄—÷ «·„ÿ«»ﬁ ›ﬁÿ
        If Me.CheckBox4.value = True Or _
           (Trim(CStr(ws.Cells(i3, "B").value)) = Trim(Me.TextBox2.value) And _
            Trim(CStr(ws.Cells(i3, "E").value)) = Trim(Me.TextBox5.value)) Then
            
            isFound = True
            With Me.ListBox1
                .AddItem
                .List(.ListCount - 1, 8) = ws.Cells(i3, "A").value
                .List(.ListCount - 1, 7) = ws.Cells(i3, "B").value
                .List(.ListCount - 1, 6) = ws.Cells(i3, "E").value
                .List(.ListCount - 1, 5) = ws.Cells(i3, "AE").value
                .List(.ListCount - 1, 4) = ws.Cells(i3, "BZ").value
                .List(.ListCount - 1, 3) = ws.Cells(i3, "CC").value
                .List(.ListCount - 1, 2) = ws.Cells(i3, "CD").value
                .List(.ListCount - 1, 1) = ws.Cells(i3, "CE").value
                .List(.ListCount - 1, 0) = ws.Cells(i3, "CF").value
                .List(.ListCount - 1, 9) = i3 '  Œ“Ì‰ —ﬁ„ «·’›
            End With
        End If
    Next i3
    
    ' 3.  ‰»ÌÂ ›Ì Õ«· ⁄œ„ ÊÃÊœ ‰ «∆Ã (⁄‰œ„«  ﬂÊ‰ «·›· —… „›⁄·… ›ﬁÿ)
    If isFound = False And Me.CheckBox4.value = False Then
        With Me.ListBox1
            .AddItem
            .List(.ListCount - 1, 2) = "·«  ÊÃœ ‰ «∆Ã „ÿ«»ﬁ… ··«”„ «·„Õœœ"
        End With
    End If
End Sub

Private Sub UserForm_Initialize()
    ' ÷»ÿ «·⁄‰Ê«‰ ⁄‰œ «·› Õ
    If Me.CheckBox4.value = True Then
        Me.CheckBox4.Caption = "≈ŸÂ«— Ã„Ì⁄ «·„ €Ì—« "
    Else
        Me.CheckBox4.Caption = "≈ŸÂ«— »Ì«‰«  Õ”» «·«”„"
    End If
    
    '  ⁄»∆… «·ﬁ«∆„…
  '  Call RefreshListBox
End Sub
