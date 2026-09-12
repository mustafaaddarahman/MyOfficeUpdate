VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm29 
   Caption         =   "UserForm29"
   ClientHeight    =   13980
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12240
   OleObjectBlob   =   "UserForm29.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






















Private Sub ComboBox1_AfterUpdate()

End Sub

Private Sub ComboBox1_Change()
On Error Resume Next

If ComboBox1.value = "«·„Êﬁ› «·Õ«·Ì" Then

Label25.Visible = False
'=======================================================
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 2) = TextBox68.Text And Sheets(1).Cells(Y, 5) = TextBox60.Text Then
   
    
     ComboBox1.Text = Sheets(1).Cells(Y, 58)
     TextBox67.Text = Sheets(1).Cells(Y, 49)
    


    End If
    Next Y

Dim num As Long

num = val(TextBox67.value / 30) * val(TextBox3.value)
TextBox66.value = num

  TextBox4.value = val(TextBox65) + val(TextBox66)
ElseIf ComboBox1.value = "„Õ«· ⁄·Ï «· ﬁ«⁄œ" Then
Label25.Visible = True
TextBox4.value = val(TextBox65)
End If
Dim lastr1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For y1 = 9 To lastr1
    
    If Sheets(1).Cells(y1, 2) = TextBox68.Text And Sheets(1).Cells(y1, 5) = TextBox60.Text Then
        TextBox67.Text = Sheets(1).Cells(y1, 49)
    









End If
Next


End Sub





Private Sub ComboBox2_Change()
Dim i As Integer
    
    If Me.ComboBox2.value = "≈Õ«·… ⁄·Ï «· ﬁ«⁄œ" Then
        ' 1. ”Õ» «·»Ì«‰«  „‰ «·„Êﬁ› «·√»ÊÌ (UserForm1) ≈·Ï «·„’›Ê›«  «·⁄«„…
        With UserForm1
            ' --- «·«” Õﬁ«ﬁ«  (13 Õﬁ·) ---
            Pos1_Data(1) = val(.TextBox10.value)  ' «·«”„Ì
            Pos1_Data(2) = val(.TextBox14.value)  ' «·“ÊÃÌ…
            Pos1_Data(3) = val(.TextBox16.value)  ' «·‘Â«œ…
            Pos1_Data(4) = val(.TextBox18.value)  ' «·Õ—›…
            Pos1_Data(5) = val(.TextBox19.value)  ' „‰’»
            Pos1_Data(6) = val(.TextBox21.value)  ' Â‰œ”Ì…
            Pos1_Data(7) = val(.TextBox20.value)  ' „ﬁÿÊ⁄…
            Pos1_Data(8) = val(.TextBox25.value)  ' ŒÿÊ—…
            Pos1_Data(9) = val(.TextBox27.value)  '  —ﬁÌ…
            Pos1_Data(10) = val(.TextBox29.value) ' €·«¡
            Pos1_Data(11) = val(.TextBox30.value) ' „‰«ÿﬁ
            Pos1_Data(12) = val(.TextBox31.value) ' √Œ—Ï
            Pos1_Data(13) = val(.TextBox49.value) ' Õﬁ· 49

            ' --- «·«” ﬁÿ«⁄«  (13 Õﬁ·) ---
            Ded1_Data(1) = val(.TextBox24.value)  '  ﬁ«⁄œ 10%
            Ded1_Data(2) = val(.TextBox32.value)  ' ÷—Ì»…
            Ded1_Data(3) = val(.TextBox33.value)  '  ÊﬁÌ›« 
            Ded1_Data(4) = val(.TextBox34.value)  ' €Ì«»
            Ded1_Data(5) = val(.TextBox35.value)  ' ≈Ã«“…
            Ded1_Data(6) = val(.TextBox36.value)  ' ⁄ﬁÊ»…
            Ded1_Data(7) = val(.TextBox38.value)  ' ”·›… “Ê«Ã
            Ded1_Data(8) = val(.TextBox40.value)  ' ”·›… ⁄ﬁ«—
            Ded1_Data(9) = val(.TextBox42.value)  ' <--- «·ﬁœÕ «·„ÿ·Ê» ·‹ 111 Ê 112
            Ded1_Data(10) = val(.TextBox44.value) ' √„«‰« 
            Ded1_Data(11) = val(.TextBox46.value) ' √Œ—Ï 1
            Ded1_Data(12) = val(.TextBox47.value) ' √Œ—Ï 2
            Ded1_Data(13) = val(.TextBox50.value) ' Õﬁ· 50
        End With

        ' 2.  ’›Ì— «·„Êﬁ› «·À«‰Ì  „«„« ›Ì «·–«ﬂ—…
        For i = 1 To 13
            Pos2_Data(i) = 0
            Ded2_Data(i) = 0
        Next i

        ' 3. ÷Œ «·»Ì«‰«  ›Ê—« ›Ì Ê«ÃÂ… «·Õ”«»«  (UserForm29)
        Call Refresh_UserForm29_Display
        
        MsgBox " „  ÕœÌÀ «·»Ì«‰«  »«·ﬂ«„· ÊÕﬁ‰ Õﬁ· 42 ›Ì «·„Êﬁ› «·√Ê·", vbInformation
       Else
       Me.TextBox111.value = val(UserForm1.TextBox42.value)
Me.TextBox112.value = val(UserForm1.TextBox42.value)
        
    End If
End Sub

Private Sub CommandButton1_Click()


End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub



Private Sub CommandButton3_Click()
' «· Õﬁﬁ „‰ ÊÃÊœ «·Ê«ÃÂ… «·—∆Ì”Ì…
    If UserForm1 Is Nothing Then
        MsgBox "«·Ê«ÃÂ… «·—∆Ì”Ì… €Ì— ‰‘ÿ…", vbCritical: Exit Sub
    End If

    ' --- √Ê·«:  —ÕÌ· «·„»«·€ «·„Ã„⁄… («·«” Õﬁ«ﬁ« ) ---
    ' Ì „ ‰ﬁ· «·ﬁÌ„ „‰ Œ«‰«  «·Ã„⁄ (122-133) ≈·Ï Œ«‰«  «·ÌÊ“—›Ê—„ 1 «·√’·Ì…
    With UserForm1
        .TextBox10.value = Me.TextBox122.value  ' «·—« » «·«”„Ì «·Ã“∆Ì
        .TextBox14.value = Me.TextBox123.value
        .TextBox16.value = Me.TextBox124.value
        .TextBox18.value = Me.TextBox125.value
        .TextBox19.value = Me.TextBox126.value
        .TextBox21.value = Me.TextBox127.value
        .TextBox20.value = Me.TextBox128.value
        .TextBox25.value = Me.TextBox129.value
        .TextBox27.value = Me.TextBox130.value
        .TextBox29.value = Me.TextBox131.value
        .TextBox30.value = Me.TextBox132.value  ' ≈Ã„«·Ì «·„Œ’’«  «·Ã“∆Ì…
        .TextBox31.value = Me.TextBox133.value  ' ≈Ã„«·Ì «·«” Õﬁ«ﬁ «·Ã“∆Ì

        ' --- À«‰Ì«:  —ÕÌ· «·„»«·€ «·„Ã„⁄… («·«” ﬁÿ«⁄« ) ---
        ' Ì „ ‰ﬁ· «·ﬁÌ„ „‰ Œ«‰«  «·Ã„⁄ (135-146)
        .TextBox24.value = Me.TextBox135.value
        .TextBox32.value = Me.TextBox136.value
        .TextBox33.value = Me.TextBox137.value
        .TextBox34.value = Me.TextBox138.value
        .TextBox35.value = Me.TextBox139.value
        .TextBox36.value = Me.TextBox140.value
        .TextBox38.value = Me.TextBox141.value
        .TextBox40.value = Me.TextBox142.value
        .TextBox42.value = Me.TextBox143.value
        .TextBox44.value = Me.TextBox144.value
        .TextBox46.value = Me.TextBox145.value
        .TextBox47.value = Me.TextBox146.value

        ' --- À«·À«:  À»Ì  «·Õ«·… «·ÊŸÌ›Ì… (—« » Ã“∆Ì) ---
        .ComboBox3.value = "—« » Ã“∆Ì"      ' ÷»ÿ «·ﬁ«∆„… «·„‰”œ·…
        .TextBox100.value = "11"            ' —„“  ⁄—Ì› «·—« » «·Ã“∆Ì
        .TextBox58.value = "—« » Ã“∆Ì"      ' »Ì«‰ «·Õ«·… ›Ì Õﬁ· «·„·«ÕŸ« 

        ' --- —«»⁄«: «· Õﬂ„ ›Ì √„‰ «·»Ì«‰«  ( ⁄ÿÌ· «·√“—«—) ---
        .CommandButton16.Enabled = False    '  ⁄ÿÌ· “— «·„Êﬁ› «·À«‰Ì
        .CommandButton16.BackColor = RGB(220, 220, 220) '  €ÌÌ— ·Ê‰Â ··—„«œÌ
        
        '  ‰»ÌÂ «·„” Œœ„ »«·‰Ã«Õ
        MsgBox " „ „⁄«·Ã… «·Õ«·… ﬂ‹ (—« » Ã“∆Ì) »‰Ã«Õ." & vbCrLf & _
               "«·„Êﬁ› «·√Ê·: " & Me.TextBox120.value & " ÌÊ„." & vbCrLf & _
               "«·„Êﬁ› «·À«‰Ì: " & Me.TextBox121.value & " ÌÊ„.", vbInformation, "‰Ÿ«„ «·—Ê« » «·Ã“∆Ì…"
    End With

    ' ≈€·«ﬁ Ê«ÃÂ… «·„⁄«·Ã… «·Ã“∆Ì…
    
    
   '=========================================================== «·«Õ ”«» ·Ã“∆Ì… «·—« »
    UserForm1.TextBox30.value = val(UserForm1.TextBox14) + val(UserForm1.TextBox105) + val(UserForm1.TextBox16) + val(UserForm1.TextBox18) + val(UserForm1.TextBox19) + val(UserForm1.TextBox21) + val(UserForm1.TextBox20) _
+ val(UserForm1.TextBox25) + val(UserForm1.TextBox27) + val(UserForm1.TextBox29) + val(UserForm1.TextBox60) + val(UserForm1.TextBox84) + val(UserForm1.TextBox86)

'============================„Ã„Ê⁄ «·«” Õﬁ«ﬁ« 
UserForm1.TextBox31.value = val(UserForm1.TextBox10) + val(UserForm1.TextBox30)
'============================«· Êﬁ›«  «· ﬁ«⁄œÌÂ


'===============
 '  «·«” ﬁÿ«⁄« 


Dim sumastktaa As Double
sumastktaa = val(UserForm1.TextBox24) + val(UserForm1.TextBox32) + val(UserForm1.TextBox33) + val(UserForm1.TextBox34) + val(UserForm1.TextBox35) + val(UserForm1.TextBox36) _
+ val(UserForm1.TextBox38) + val(UserForm1.TextBox40) + val(UserForm1.TextBox42) + val(UserForm1.TextBox44) + val(UserForm1.TextBox46) _
 + val(UserForm1.TextBox50) + val(UserForm1.TextBox62) + val(UserForm1.TextBox83) + val(UserForm1.TextBox85)
 
 
 UserForm1.TextBox48.value = sumastktaa
UserForm1.TextBox49.value = val(UserForm1.TextBox31) - val(UserForm1.TextBox48)
UserForm1.TextBox48.value = val(UserForm1.TextBox47) + sumastktaa
'=======================«·—« » «·’«›Ì
UserForm1.TextBox49.value = val(UserForm1.TextBox31) - val(UserForm1.TextBox48)

'=======================================================
Call update_data_acount_22

    
    
  '========================================   = —ÕÌ· «·»Ì«‰« 
    Dim ws37 As Worksheet, ws46 As Worksheet
    Dim lastR As Long, iRow As Long, Y As Long
    
    ' ÷»ÿ «·‘Ì «  «·„” Âœ›…
    Set ws37 = ThisWorkbook.Sheets(37)
    Set ws46 = ThisWorkbook.Sheets(46)
    
    ' 1.  ‰ŸÌ› «·”Ã·«  «·ﬁœÌ„… ( Ã‰» «· ﬂ—«—)
    ' ‰⁄ „œ ⁄·Ï TextBox68 («·—ﬁ„ «·ÊŸÌ›Ì) Ê TextBox60 («·«”„) ›Ì UF29
    lastR = ws37.Cells(ws37.Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
        If ws37.Cells(Y, 1).Text = Me.TextBox68.Text And ws37.Cells(Y, 2).Text = Me.TextBox60.Text Then
            ws37.Rows(Y).Delete
            Exit For
        End If
    Next Y

    lastR = ws46.Cells(ws46.Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
        If ws46.Cells(Y, 1).Text = Me.TextBox68.Text And ws46.Cells(Y, 2).Text = Me.TextBox60.Text Then
            ws46.Rows(Y).Delete
            Exit For
        End If
    Next Y

    ' 2. «· —ÕÌ· ≈·Ï ‘Ì  37 («·”Ã· «·À«» )
    iRow = ws37.Cells(ws37.Rows.count, 1).End(xlUp).row + 1
    With ws37
        .Cells(iRow, 1).value = Me.TextBox68.value       ' «·—ﬁ„ «·ÊŸÌ›Ì
        .Cells(iRow, 2).value = Me.TextBox60.value       ' «”„ «·„ÊŸ›
        
        '  Ê“Ì⁄ «·„Œ—Ã«  «·ÃœÌœ… Õ”»  Ê÷ÌÕﬂ:
        .Cells(iRow, 3).value = Me.TextBox65.value       ' „Ã„Ê⁄ «·„Êﬁ› «·√Ê· «·√’·Ì
        .Cells(iRow, 6).value = Me.TextBox66.value       ' „Ã„Ê⁄ «·„Êﬁ› «·À«‰Ì «·√’·Ì
        .Cells(iRow, "i").value = Me.TextBox120.value
        ' «·‰ ÌÃ… «·‰Â«∆Ì… («·’«›Ì «·„Õ ”» ··ﬁ»÷)
        .Cells(iRow, 11).value = Me.TextBox67.value      ' «·’«›Ì (122:131 - 134:145)
        
        ' «· Ê«—ÌŒ (”Õ»Â« „‰ UserForm1 ﬂ„« ÂÌ „Œ“‰…)
        .Cells(iRow, 12).value = Format(UserForm1.TextBox64.value, "yyyy/mm/dd")
        .Cells(iRow, 13).value = Format(UserForm1.TextBox65.value, "yyyy/mm/dd")
        
        .Cells(iRow, 14).value = "«Õ ”«» Ã“∆Ì „Õ›ÊŸ («·ŒÊ«—“„Ì… «·ÃœÌœ…)"
    End With

    ' 3. «· —ÕÌ· ≈·Ï ‘Ì  46 (”Ã· «·Œ·«’« )
    iRow = ws46.Cells(ws46.Rows.count, 1).End(xlUp).row + 1
    With ws46
        .Cells(iRow, 1).value = Me.TextBox68.value
        .Cells(iRow, 2).value = Me.TextBox60.value
        .Cells(iRow, 6).value = Me.TextBox67.value      '  ”ÃÌ· «·’«›Ì «·‰Â«∆Ì ›Ì «·Œ·«’…
        .Cells(iRow, "e").value = Me.TextBox120.value
        '  ÊÀÌﬁ «·‘Â— „‰ «·Œ·Ì… M1
        If IsDate(.Cells(1, "M").value) Then
            .Cells(iRow, 8).value = monthName(Month(.Cells(1, "M").value))
        End If
    End With

    ' 4.  ÕœÌÀ Õ«·… UserForm1 ·Ì⁄ﬂ” «·‰ ÌÃ… «·‰Â«∆Ì…
    With UserForm1
        .TextBox31.value = Me.TextBox67.value  ' Ê÷⁄ «·’«›Ì ›Ì Œ«‰… «·«” Õﬁ«ﬁ »«·Ê«ÃÂ… «·—∆Ì”Ì…
        .TextBox49.Text = Me.TextBox67.value
        .TextBox58.Text = "—« » Ã“∆Ì („Õ›ÊŸ)"
    End With

    MsgBox " „  —ÕÌ· «·»Ì«‰«  »‰Ã«Õ Ê›ﬁ  Ê“Ì⁄ «·„Œ—Ã«  (65° 66° 67)", vbInformation, " √ﬂÌœ «·Õ›Ÿ"
 
    
    
    
    
    
    
    Unload Me
End Sub

Private Sub Label32_Click()

End Sub








Private Sub Label39_Click()

End Sub

Private Sub Label59_Click()

End Sub

Private Sub TextBox120_Change()
Dim Days1 As Double, Days2 As Double
    Dim Fact1 As Double, Fact2 As Double
    Dim CaseType As String
    
    On Error Resume Next ' Õ„«Ì… ÷œ «·„œŒ·«  €Ì— «·—ﬁ„Ì…
    
    ' 2. ﬁ—«¡… «·Õ«·… „‰ «·Ê«ÃÂ… «·—∆Ì”Ì…
    CaseType = UserForm1.ComboBox2.value
    
    ' 3.  ÕœÌœ „—ﬂ“  ﬁ”Ì„ «·√Ì«„ Ê«·„⁄«„·«  «·Õ”«»Ì…
    ' -------------------------------------------------------
    Days1 = val(Me.TextBox120.value)
    If Days1 > 30 Then
        Days1 = 30
        Me.TextBox120.value = 30
    End If
    
    ' „⁄«„· «·„Êﬁ› «·√Ê·
    Fact1 = Days1 / 30
    
    ' ›Õ’ «·”Ì‰«—ÌÊ: Â· ÂÊ „Êﬁ› √Õ«œÌ ( ﬁ«⁄œ) √„ „Êﬁ›Ì‰ (Ã“∆Ì)ø
    If CaseType = "≈Õ«·… ⁄·Ï «· ﬁ«⁄œ" Then
        ' --- ”Ì‰«—ÌÊ «·„Êﬁ› «·√Õ«œÌ ---
        Days2 = 0
        Me.TextBox121.value = "„‰›ﬂ/„ÊﬁÊ›"
        Fact2 = 0
    Else
        ' --- ”Ì‰«—ÌÊ «·„Êﬁ›Ì‰ («·—« » «·Ã“∆Ì) ---
        ' «” Œœ«„ Int ·÷„«‰ »ﬁ«¡ ⁄œœ «·√Ì«„ —ﬁ„« ’ÕÌÕ«
        Days2 = Int(30 - Days1)
        Me.TextBox121.value = Days2
        Fact2 = Days2 / 30
    End If

    ' 4. „⁄«·Ã… «·«” Õﬁ«ﬁ«  («” »œ«· Round »‹ Int ·» — «·ﬂ”Ê— ‰Â«∆Ì«)
    ' ----------------------------------------------------------------------
    Me.TextBox122.value = Int((val(Me.TextBox70.value) * Fact1) + (val(Me.TextBox71.value) * Fact2))
    Me.TextBox123.value = Int((val(Me.TextBox72.value) * Fact1) + (val(Me.TextBox73.value) * Fact2))
    Me.TextBox124.value = Int((val(Me.TextBox74.value) * Fact1) + (val(Me.TextBox75.value) * Fact2))
    Me.TextBox125.value = Int((val(Me.TextBox76.value) * Fact1) + (val(Me.TextBox77.value) * Fact2))
    Me.TextBox126.value = Int((val(Me.TextBox79.value) * Fact1) + (val(Me.TextBox80.value) * Fact2))
    Me.TextBox127.value = Int((val(Me.TextBox81.value) * Fact1) + (val(Me.TextBox82.value) * Fact2))
    Me.TextBox128.value = Int((val(Me.TextBox83.value) * Fact1) + (val(Me.TextBox84.value) * Fact2))
    Me.TextBox129.value = Int((val(Me.TextBox78.value) * Fact1) + (val(Me.TextBox85.value) * Fact2))
    Me.TextBox130.value = Int((val(Me.TextBox87.value) * Fact1) + (val(Me.TextBox88.value) * Fact2))
    Me.TextBox131.value = Int((val(Me.TextBox89.value) * Fact1) + (val(Me.TextBox90.value) * Fact2))
    Me.TextBox132.value = Int((val(Me.TextBox91.value) * Fact1) + (val(Me.TextBox92.value) * Fact2))
    Me.TextBox133.value = Int((val(Me.TextBox93.value) * Fact1) + (val(Me.TextBox94.value) * Fact2))

    ' 5. „⁄«·Ã… «·«” ﬁÿ«⁄«  («” »œ«· Round »‹ Int ·» — «·ﬂ”Ê— ‰Â«∆Ì«)
    ' ----------------------------------------------------------------------
    Me.TextBox135.value = Int((val(Me.TextBox95.value) * Fact1) + (val(Me.TextBox96.value) * Fact2))
    Me.TextBox136.value = Int((val(Me.TextBox97.value) * Fact1) + (val(Me.TextBox98.value) * Fact2))
    Me.TextBox137.value = Int((val(Me.TextBox99.value) * Fact1) + (val(Me.TextBox100.value) * Fact2))
    Me.TextBox138.value = Int((val(Me.TextBox86.value) * Fact1) + (val(Me.TextBox101.value) * Fact2))
    Me.TextBox139.value = Int((val(Me.TextBox103.value) * Fact1) + (val(Me.TextBox105.value) * Fact2))
    Me.TextBox140.value = Int((val(Me.TextBox104.value) * Fact1) + (val(Me.TextBox106.value) * Fact2))
    Me.TextBox141.value = Int((val(Me.TextBox107.value) * Fact1) + (val(Me.TextBox108.value) * Fact2))
    Me.TextBox142.value = Int((val(Me.TextBox102.value) * Fact1) + (val(Me.TextBox109.value) * Fact2))
    Me.TextBox143.value = Int((val(Me.TextBox111.value) * Fact1) + (val(Me.TextBox112.value) * Fact2))
    Me.TextBox144.value = Int((val(Me.TextBox113.value) * Fact1) + (val(Me.TextBox114.value) * Fact2))
    Me.TextBox145.value = Int((val(Me.TextBox115.value) * Fact1) + (val(Me.TextBox116.value) * Fact2))
    Me.TextBox146.value = Int((val(Me.TextBox117.value) * Fact1) + (val(Me.TextBox118.value) * Fact2))

    ' 6.  ÕœÌÀ «·Ê«ÃÂ… ·÷„«‰ ŸÂÊ— «·‰ «∆Ã ›Ê—«
    Me.Repaint
    
    '===================================================== «·ﬁÌ„ «· Ã„Ì⁄Ì… («·”„Ì‘‰) »«·‹ Int
    
    Dim Sum_65 As Double, Sum_66 As Double
    Dim Sum_Estihkak_Part As Double, Sum_Ded_Part As Double
    Dim i As Integer
    
    On Error Resume Next

    ' --- 1. Õ”«»  Ìﬂ”  »Êﬂ” 65 (» — «·ﬂ”— ··„Ã„Ê⁄) ---
    Sum_65 = val(Me.TextBox70.value) + val(Me.TextBox72.value) + val(Me.TextBox74.value) + _
             val(Me.TextBox76.value) + val(Me.TextBox79.value) + val(Me.TextBox81.value) + _
             val(Me.TextBox83.value) + val(Me.TextBox78.value) + val(Me.TextBox87.value) + _
             val(Me.TextBox89.value)
    Me.TextBox65.value = Format(Int(Sum_65), "0")

    ' --- 2. Õ”«»  Ìﬂ”  »Êﬂ” 66 (» — «·ﬂ”— ··„Ã„Ê⁄) ---
    Sum_66 = val(Me.TextBox71.value) + val(Me.TextBox73.value) + val(Me.TextBox75.value) + _
             val(Me.TextBox77.value) + val(Me.TextBox80.value) + val(Me.TextBox82.value) + _
             val(Me.TextBox84.value) + val(Me.TextBox85.value) + val(Me.TextBox88.value) + _
             val(Me.TextBox90.value)
    Me.TextBox66.value = Format(Int(Sum_66), "0")

    ' --- 3. Õ”«»  Ìﬂ”  »Êﬂ” 67 («·’«›Ì «·Ã“∆Ì «·‰Â«∆Ì) ---
    ' √- „Ã„Ê⁄ «·«” Õﬁ«ﬁ«  «·Ã“∆Ì… „⁄ » — ﬂ”— ﬂ· Œ«‰… √À‰«¡ «·Ã„⁄
    Sum_Estihkak_Part = 0
    For i = 122 To 131
        Sum_Estihkak_Part = Sum_Estihkak_Part + Int(val(Me.Controls("TextBox" & i).value))
    Next i
    
    ' »- „Ã„Ê⁄ «·«” ﬁÿ«⁄«  «·Ã“∆Ì… „⁄ » — ﬂ”— ﬂ· Œ«‰… √À‰«¡ «·Ã„⁄
    Sum_Ded_Part = 0
    For i = 134 To 145
        Sum_Ded_Part = Sum_Ded_Part + Int(val(Me.Controls("TextBox" & i).value))
    Next i
    
    ' Ã- «·‰ ÌÃ… «·‰Â«∆Ì… (”„Ì‘‰ «·”„Ì‘‰) »œÊ‰ √Ì ﬂ”Ê—
    Me.TextBox67.value = Format(Int(Sum_Estihkak_Part - Sum_Ded_Part), "0")

    On Error GoTo 0
    
    
    
    
    '===========================================
    
  
Dim NominalSalary As Double
    Dim ChildAllowanceUnit As Double
    Dim Factor6 As Double, Factor7 As Double, Factor8 As Double
    Dim Factor10 As Double, Factor11 As Double
    Dim AssumedChildren As Double ' ⁄œœ «·√ÿ›«· «·„› —÷
    
    On Error Resume Next
    
    ' 1. ≈⁄œ«œ«  √”«”Ì…
    NominalSalary = val(UserForm1.TextBox10.value)
    ChildAllowanceUnit = 10000 ' ﬁÌ„… „Œ’’ «·ÿ›· «·Ê«Õœ (Ì„ﬂ‰ﬂ  €ÌÌ—Â« Õ”» «·ﬁ«‰Ê‰)
    
    If NominalSalary > 0 Then
        
        ' --- √Ê·«: Õ”«» «·‰”» («·„⁄«„·« ) «·„Ã„⁄… ··„Œ’’«  «·√Œ—Ï ---
        Factor6 = val(Me.TextBox123.value) / NominalSalary
        Factor7 = val(Me.TextBox124.value) / NominalSalary
        Factor8 = val(Me.TextBox125.value) / NominalSalary
        Factor10 = val(Me.TextBox130.value) / NominalSalary
        Factor11 = val(Me.TextBox131.value) / NominalSalary
        
        ' --- À«‰Ì«: Õ”«» "⁄œœ «·√ÿ›«· «·„› —÷" „‰ «·„»·€ «· Ã„Ì⁄Ì (128) ---
        ' «·ﬁ«⁄œ…: ⁄œœ «·√ÿ›«· = «·„»·€ «·„Ã„⁄ / ﬁÌ„… „Œ’’ «·ÿ›· «·Ê«Õœ
        AssumedChildren = val(Me.TextBox128.value) / ChildAllowanceUnit
        
        ' --- À«·À«:  Ê“Ì⁄ «·‰ «∆Ã ⁄·Ï ÕﬁÊ· «·„⁄«„·«  Ê⁄œœ «·√ÿ›«· (147-152) ---
        Me.TextBox147.value = Format(Factor6, "0.000000")
        Me.TextBox148.value = Format(Factor7, "0.000000")
        Me.TextBox149.value = Format(Factor8, "0.000000")
        Me.TextBox150.value = Format(Factor10, "0.000000")
        Me.TextBox151.value = Format(Factor11, "0.000000")
        
        ' Ê÷⁄ ⁄œœ «·√ÿ›«· «·„› —÷ ›Ì 152
        Me.TextBox152.value = Format(AssumedChildren, "0.00")
        
        ' --- —«»⁄«: «· —ÕÌ· «·‰Â«∆Ì ··›Ê—„ «·—∆Ì”Ì ÊÊ«ÃÂ… «·»Ì«‰«  ---
        UserForm1.ComboBox6.value = Me.TextBox147.value
        UserForm1.ComboBox7.value = Me.TextBox148.value
        UserForm1.ComboBox8.value = Me.TextBox149.value
        UserForm1.ComboBox10.value = Me.TextBox150.value
        UserForm1.ComboBox11.value = Me.TextBox151.value
        
        '  —ÕÌ· ⁄œœ «·√ÿ›«· ≈·Ï TextBox22 ›Ì «·›Ê—„ «·—∆Ì”Ì
        UserForm1.TextBox22.value = Me.TextBox152.value
        
        ' --- Œ«„”«:  ÕœÌÀ „’›Ê›… «·–«ﬂ—… ·€—÷ «·Õ›Ÿ ---
        Ded1_Data(14) = Factor6
        Ded1_Data(15) = Factor7
        Ded1_Data(16) = Factor8
        Ded1_Data(17) = Factor10
        Ded1_Data(18) = Factor11
        Ded1_Data(19) = AssumedChildren '  Œ“Ì‰ ⁄œœ «·√ÿ›«· ›Ì «·„’›Ê›…
        
    End If
    
    On Error GoTo 0
    
    
    
    
End Sub

Private Sub TextBox2_AfterUpdate()
TextBox65.value = val(TextBox61.value) / 30 * val(TextBox2.value)
 TextBox3.value = 30 - val(TextBox2.value)

End Sub

Private Sub TextBox2_Change()
'On Error Resume Next
TextBox65.value = val(TextBox61.value) / 30 * val(TextBox2.value)
 TextBox3.value = 30 - val(TextBox2.value)

End Sub

Private Sub TextBox3_AfterUpdate()
On Error Resume Next


TextBox66.value = val(TextBox67.value) / 30 * val(TextBox3.value)
End Sub

Private Sub TextBox3_Change()
On Error Resume Next

TextBox66.value = val(TextBox67.value) / 30 * val(TextBox3.value)

End Sub

Private Sub TextBox60_Change()

End Sub

Private Sub TextBox62_AfterUpdate()
If TextBox62.Text = "" Then
TextBox60.Text = ""
End If
End Sub

Private Sub TextBox62_Change()

Dim lastR, r As Integer
lastR = Sheets(38).Cells(Rows.count, 1).End(xlUp).row
For r = 2 To lastR
If Sheets(38).Cells(r, 2).value Like "*" & TextBox62.Text & "*" Then
TextBox68.Text = Sheets(38).Cells(r, 1)
 TextBox60.Text = Sheets(38).Cells(r, 2)
 TextBox1.Text = Sheets(38).Cells(r, 3)
 TextBox61.Text = Sheets(38).Cells(r, 4)
  TextBox69.Text = Sheets(38).Cells(r, 5)
 
 
 
 
Exit For
End If
Next
End Sub

Private Sub TextBox66_AfterUpdate()
 TextBox4.value = Format(val(TextBox65) + val(TextBox66), "00000000")
End Sub

Private Sub TextBox66_Change()

 

End Sub

Private Sub TextBox67_Change()

End Sub

Private Sub TextBox68_Change()

End Sub

Private Sub UserForm_Activate()
' ---  ›—Ì€ «·„Êﬁ› «·√Ê· («·„Ã„Ê⁄… «·›—œÌ…/«·ﬁœÌ„…) ---
   Me.TextBox70.value = Pos1_Data(1)
Me.TextBox72.value = Pos1_Data(2)
Me.TextBox74.value = Pos1_Data(3)
Me.TextBox76.value = Pos1_Data(4)
Me.TextBox79.value = Pos1_Data(5)
Me.TextBox81.value = Pos1_Data(6)
Me.TextBox83.value = Pos1_Data(7)
Me.TextBox78.value = Pos1_Data(8)
Me.TextBox87.value = Pos1_Data(9)
Me.TextBox89.value = Pos1_Data(10)
Me.TextBox91.value = Pos1_Data(11) ' Ìﬁ«»· 30
Me.TextBox93.value = Pos1_Data(12) ' Ìﬁ«»· 31
Me.TextBox95.value = Pos1_Data(13) ' Ìﬁ«»· 49

' ---  ›—Ì€ «·„Êﬁ› «·À«‰Ì (Pos2) ---
Me.TextBox71.value = Pos2_Data(1)
Me.TextBox73.value = Pos2_Data(2)
Me.TextBox75.value = Pos2_Data(3)
Me.TextBox77.value = Pos2_Data(4)
Me.TextBox80.value = Pos2_Data(5)
Me.TextBox82.value = Pos2_Data(6)
Me.TextBox84.value = Pos2_Data(7)
Me.TextBox85.value = Pos2_Data(8)
Me.TextBox88.value = Pos2_Data(9)
Me.TextBox90.value = Pos2_Data(10)
Me.TextBox92.value = Pos2_Data(11) ' Ìﬁ«»· 30
Me.TextBox94.value = Pos2_Data(12) ' Ìﬁ«»· 31
Me.TextBox96.value = Pos2_Data(13) ' Ìﬁ«»· 49

' ---  ›—Ì€ «” ﬁÿ«⁄«  «·„Êﬁ› «·√Ê· (Ded1) ---
Me.TextBox95.value = Ded1_Data(1)
Me.TextBox97.value = Ded1_Data(2)
Me.TextBox99.value = Ded1_Data(3)
Me.TextBox86.value = Ded1_Data(4)
Me.TextBox103.value = Ded1_Data(5)
Me.TextBox104.value = Ded1_Data(6)
Me.TextBox107.value = Ded1_Data(7)
Me.TextBox102.value = Ded1_Data(8)
Me.TextBox111.value = Ded1_Data(9)   ' Ìﬁ«»· «· Ìﬂ”  »Êﬂ” 42 ›Ì «·„Êﬁ› «·√Ê· ﬂ„« ÿ·» 
Me.TextBox113.value = Ded1_Data(10)
Me.TextBox115.value = Ded1_Data(11)
Me.TextBox117.value = Ded1_Data(12)
Me.TextBox110.value = Ded1_Data(13)

' ---  ›—Ì€ «” ﬁÿ«⁄«  «·„Êﬁ› «·À«‰Ì (Ded2) ---
Me.TextBox96.value = Ded2_Data(1)
Me.TextBox98.value = Ded2_Data(2)
Me.TextBox100.value = Ded2_Data(3)
Me.TextBox101.value = Ded2_Data(4)
Me.TextBox105.value = Ded2_Data(5)
Me.TextBox106.value = Ded2_Data(6)
Me.TextBox108.value = Ded2_Data(7)
Me.TextBox109.value = Ded2_Data(8)
Me.TextBox112.value = Ded2_Data(9)   ' Ìﬁ«»· «· Ìﬂ”  »Êﬂ” 42 ›Ì «·„Êﬁ› «·À«‰Ì ﬂ„« ÿ·» 
Me.TextBox114.value = Ded2_Data(10)
Me.TextBox116.value = Ded2_Data(11)
Me.TextBox118.value = Ded2_Data(12)
Me.TextBox119.value = Ded2_Data(13)

Me.Repaint

End Sub


Private Sub UserForm_Initialize()
' 1.  ›—Ì€ «·ﬂÊ„»Ê »Êﬂ”  „«„« ·÷„«‰ ⁄œ„ ÊÃÊœ »Ì«‰«  ﬁœÌ„…
    Me.ComboBox2.Clear
    
    ' 2. ≈÷«›… «·ﬁÌ„ Ì‰ ( √ﬂœ „‰ ﬂ «» Â„« »Â–« «·‘ﬂ· ·ﬂÌ Ìﬁ—√Â„« ﬂÊœ «·Õ”«»)
    With Me.ComboBox2
        .AddItem "—« » Ã“∆Ì"            ' «·ﬁÌ„… «·√Ê·Ï (··„Êﬁ›Ì‰)
        .AddItem "≈Õ«·… ⁄·Ï «· ﬁ«⁄œ"    ' «·ﬁÌ„… «·À«‰Ì… (··„Êﬁ› «·Ê«Õœ)
    End With
    
    ' 3. ÷»ÿ Œ’«∆’ «·ﬁ«∆„… · ”ÂÌ· «·«Œ Ì«—
    Me.ComboBox2.ListRows = 2          ' ≈Ã»«— «·ﬁ«∆„… ⁄·Ï ≈ŸÂ«— ”ÿ—Ì‰
    Me.ComboBox2.Style = fmStyleDropDownList
    
    Me.TextBox68.value = UserForm1.TextBox2.value
    
    ' 2. «”„ «·„ÊŸ›: „‰ TextBox5 ›Ì UF1 ≈·Ï TextBox60 ›Ì UF29
    Me.TextBox60.value = UserForm1.TextBox5.value
    
    ' Ì„ﬂ‰ﬂ ≈÷«›… √Ì »Ì«‰«  √Œ—Ï  Õ «Ã Ã·»Â« Â‰« »‰›” «·ÿ—Ìﬁ…
    
End Sub


Sub Refresh_UserForm29_Display()
    ' On Error Resume Next ·÷„«‰ ⁄œ„  Êﬁ› «·»—‰«„Ã ≈–« ﬂ«‰ UserForm29 „€·ﬁ«
    On Error Resume Next
    
    With UserForm29
        ' --- [1]  Ê“Ì⁄ «” Õﬁ«ﬁ«  «·„Êﬁ› «·√Ê· («·„’œ—: Pos1_Data) ---
        .TextBox70.value = Pos1_Data(1):  .TextBox72.value = Pos1_Data(2)
        .TextBox74.value = Pos1_Data(3):  .TextBox76.value = Pos1_Data(4)
        .TextBox79.value = Pos1_Data(5):  .TextBox81.value = Pos1_Data(6)
        .TextBox83.value = Pos1_Data(7):  .TextBox78.value = Pos1_Data(8)
        .TextBox87.value = Pos1_Data(9):  .TextBox89.value = Pos1_Data(10)
        .TextBox91.value = Pos1_Data(11): .TextBox93.value = Pos1_Data(12)
        .TextBox95.value = Pos1_Data(13)

        ' --- [2]  Ê“Ì⁄ «” Õﬁ«ﬁ«  «·„Êﬁ› «·À«‰Ì («·„’œ—: Pos2_Data - √’›«— ›Ì «· ﬁ«⁄œ) ---
        .TextBox71.value = Pos2_Data(1):  .TextBox73.value = Pos2_Data(2)
        .TextBox75.value = Pos2_Data(3):  .TextBox77.value = Pos2_Data(4)
        .TextBox80.value = Pos2_Data(5):  .TextBox82.value = Pos2_Data(6)
        .TextBox84.value = Pos2_Data(7):  .TextBox85.value = Pos2_Data(8)
        .TextBox88.value = Pos2_Data(9):  .TextBox90.value = Pos2_Data(10)
        .TextBox92.value = Pos2_Data(11): .TextBox94.value = Pos2_Data(12)
        .TextBox96.value = Pos2_Data(13)

        ' --- [3]  Ê“Ì⁄ «” ﬁÿ«⁄«  «·„Êﬁ› «·√Ê· («·„’œ—: Ded1_Data) ---
        .TextBox95.value = Ded1_Data(1):  .TextBox97.value = Ded1_Data(2)
        .TextBox99.value = Ded1_Data(3):  .TextBox86.value = Ded1_Data(4)
        .TextBox103.value = Ded1_Data(5): .TextBox104.value = Ded1_Data(6)
        .TextBox107.value = Ded1_Data(7): .TextBox102.value = Ded1_Data(8)
        
        ' «·Õﬁ· «·ÃÊÂ—Ì «·–Ì ”√·  ⁄‰Â («·ﬁ«œ„ „‰ TextBox42)
        .TextBox111.value = Ded1_Data(9)
        
        .TextBox113.value = Ded1_Data(10): .TextBox115.value = Ded1_Data(11)
        .TextBox117.value = Ded1_Data(12): .TextBox110.value = Ded1_Data(13)

        ' --- [4]  Ê“Ì⁄ «” ﬁÿ«⁄«  «·„Êﬁ› «·À«‰Ì («·„’œ—: Ded2_Data - √’›«— ›Ì «· ﬁ«⁄œ) ---
        .TextBox96.value = Ded2_Data(1):  .TextBox98.value = Ded2_Data(2)
        .TextBox100.value = Ded2_Data(3): .TextBox101.value = Ded2_Data(4)
        .TextBox105.value = Ded2_Data(5): .TextBox106.value = Ded2_Data(6)
        .TextBox108.value = Ded2_Data(7): .TextBox109.value = Ded2_Data(8)
        
        ' «·Õﬁ· «·ÃÊÂ—Ì ··„Êﬁ› «·À«‰Ì
        .TextBox112.value = Ded2_Data(9)
        
        .TextBox114.value = Ded2_Data(10): .TextBox116.value = Ded2_Data(11)
        .TextBox118.value = Ded2_Data(12): .TextBox119.value = Ded2_Data(13)

        '  ÕœÌÀ «·Ê«ÃÂ… ›Ê—« ·Ì—«Â« «·„” Œœ„
        .Repaint
    End With
End Sub
