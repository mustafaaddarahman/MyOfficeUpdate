Attribute VB_Name = "all_tareaba"
Sub state_tax()

On Error Resume Next
'ÇÓÊËäÇÆÇÊ ÇáÓãÇÍÇÊ ÇáÖÑíÈíÉ


 Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim empID As String: empID = UserForm1.TextBox2.value
    Dim i As Long, lastRow As Long
    Dim IsExcluded As Boolean: IsExcluded = False

    ' 1. ÇáÊÍŞŞ ãä ÍÇáÉ ÇáÌíß ÈæßÓ İí ÇáíæÒÑ İæÑã ÃæáÇğ
    If UserForm1.CheckBox5.value = True Then
        
        ' 2. ÇáÈÍË İí ÇáÔíÊ ááÊÃßÏ ãä æÌæÏ True İí ÇáÚãæÏ CG áäİÓ ÇáãæÙİ
        lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
        
        For i = 9 To lastRow
            If Trim(ws.Cells(i, "B").value) = empID Then
                ' ÇáÊÍŞŞ ÅĞÇ ßÇä ÇáÚãæÏ CG íÍÊæí Úáì ÇáŞíãÉ True
                If ws.Cells(i, "CG").value = True Or ws.Cells(i, "CG").value = "True" Then
                    IsExcluded = True ' ÇáãæÙİ ãÓÊËäì ÑÓãíÇğ
                End If
                Exit For
            End If
        Next i
    End If

    ' 3. ÇÊÎÇĞ ÇáŞÑÇÑ ÈäÇÁğ Úáì ÇáİÍÕ
    If IsExcluded = True Then
        ' Êã ÇáÚËæÑ Úáì ÇáÌíß ÈæßÓ ãİÚá æßáãÉ True İí ÇáÔíÊ
        MsgBox "Êã ÊİÚíá äÙÇã ÇáÇÓÊËäÇÁ.. ÇáÎæÇÑÒãíÉ ãÚØáÉ áåĞÇ ÇáŞíÏ.", vbInformation
        
        ' åäÇ äÌÚá ŞíãÉ ÇáÖÑíÈÉ ÕİÑ Ãæ äÎÑÌ ãä ÇáÅÌÑÇÁ
        ' State_Tax_Value = 0
        Exit Sub
    Else

'==========================================================================


If UserForm1.TextBox23.Text = "ÇáÇÚÒÈ" Or UserForm1.TextBox23.Text = "ÇáãØáŞ æáíÓ áÏíå ÍÖÇäÉ ááÃæáÇÏ" Or _
   UserForm1.TextBox23.Text = "ÇáãæÙİå ÇáãİÕæá ÏÎáåÇ Úä ÏÎá ÒæÌåÇ" Or _
   UserForm1.TextBox23.Text = "ÇáãæÙİ ÇáãÊÒæÌ æ ÒæÌÊå ÑÈÉ ÈíÊ" Or _
   UserForm1.TextBox23.Text = "ÇáãæÙİå ÇáãÊÒæÌå æÒæÌåÇ ÚÇØá Úãä ÇáÚãá" Then

    Dim rang1 As Range, rang2 As Range
    Dim taxCode As Long
    Dim salaryVal As Long
    Dim childVal As Long
    
    Set rang1 = Sheets(1).Range("zwjea4")
    Set rang2 = Sheets(2).Range("astktaat")
    
    ' ÌáÈ ÇáÑãÒ
    taxCode = CLng(Application.WorksheetFunction.VLookup(UserForm1.TextBox23.Text, rang1, 2, 0))
    salaryVal = val(UserForm1.TextBox10.Text)
    childVal = val(UserForm1.TextBox22.Text)
    
    ' --- ÊŞÓíã ÇáãÓÇÑÇÊ ÍÓÈ ÇáÊÑãíÒ ---
    
    Select Case taxCode
        
        Case 5
            ' ÇáÊÑãíÒ 5: ÇáŞÑÇÁÉ ãä ÇáÚãæÏ 3 (D) ÍÕÑÇğ (ÈÛÖ ÇáäÙÑ Úä ÇáÃØİÇá)
            UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(salaryVal, rang2, 3, 1)
            
        Case 4
            ' ÇáÊÑãíÒ 4: ÇáŞÑÇÁÉ ãä ÇáÚãæÏ 4 (E) ãÚ ÅÖÇİÉ ÚÏÏ ÇáÃØİÇá
            UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(salaryVal, rang2, 4 + childVal, 1)
            
        Case 3
            ' ÇáÊÑãíÒ 3: ÇáŞÑÇÁÉ ãä ÇáÚãæÏ 3 (D) ãÚ ÊÕİíÑ TextBox22 æ TextBox20
            UserForm1.TextBox22.value = 0
            UserForm1.TextBox20.value = 0
            UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(salaryVal, rang2, 3, 1)
            
        Case Is < 3
            ' ÇáÍÇáÇÊ ÇáÃÎÑì (ÃŞá ãä 3): íãßäß æÖÚ ãäØŞ ÎÇÕ ÈåÇ åäÇ
            UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(salaryVal, rang2, taxCode + childVal, 1)
            
        Case Else
            ' ÇİÊÑÇÖí
            UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(salaryVal, rang2, taxCode, 1)
            
    End Select
End If









'======================================================
'  ÇáãæÙİÉ ãØáŞÉ  / ÇáãæÙİÉ ÇÑãáÉ ...... ÇáÍÇáÉ ÇáÇæáì ááÖÑíÈå
On Error Resume Next
If UserForm1.TextBox23.Text = "ÇáãæÙİÉ ãØáŞÉ" Or UserForm1.TextBox23.Text = "ÇáãæÙİÉ ÇÑãáÉ" Then
Dim textvlue3 As Long
Dim textvlue4 As Long
textvlue3 = UserForm1.TextBox22.Text
textvlue4 = UserForm1.TextBox10.Text
Dim rang3 As Range
Dim rang4 As Range
Set rang3 = Sheets("thariaba1").Range("zwjea1")
Set rang4 = Sheets("thariaba1").Range("astktat1")

If UserForm1.TextBox23.Text = rang3 Then
UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(textvlue4, rang4, textvlue3 + 3, 1)
End If
End If
'========================================================
'  ÇáãæÙİÉ ãÊÒæÌÉ æÒæÌåÇ ÚÇÌÒ Úä ÇáÚãá/ ......ÇáÍÇáÉ ÇáËÇäíÉ ááÖÑíÈå
On Error Resume Next
If UserForm1.TextBox23.Text = "ÇáãæÙİÉ ãÊÒæÌÉ æÒæÌåÇ ÚÇÌÒ Úä ÇáÚãá" Then
Dim textvlue5 As Long
Dim textvlue6 As Long
textvlue5 = UserForm1.TextBox22.Text
textvlue6 = UserForm1.TextBox10.Text
Dim rang5 As Range
Dim rang6 As Range
Set rang5 = Sheets("thariaba2").Range("zwjea2")
Set rang6 = Sheets("thariaba2").Range("astktaat2")

If UserForm1.TextBox23.Text = rang5 Then
UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(textvlue6, rang6, textvlue5 + 3, 1)
End If
End If

'=========================================================
' ÇáãæÙİ ÇáãØáŞ ÇĞÇ ßÇä ÍÇÖäÇ ááÃæáÇÏ/ ãÊæİÇÉ ÒæÌÊå / ÇáãæÙİ ãÊÒæÌ æÏÎáå ãİÕæá Úä ÏÎá ÒæÌÊå ......ÇáÍÇáÉ ÇáËÇáËÉ ááÖÑíÈå
On Error Resume Next
If UserForm1.TextBox23.Text = "ÇáãæÙİ ÇáãØáŞ ÇĞÇ ßÇä ÍÇÖäÇ ááÃæáÇÏ" Or UserForm1.TextBox23.Text = "ãÊæİÇÉ ÒæÌÊå" Or UserForm1.TextBox23.Text = "ÇáãæÙİ ãÊÒæÌ æÏÎáå ãİÕæá Úä ÏÎá ÒæÌÊå" Then
Dim textvlue7 As Long
Dim textvlue8 As Long
textvlue7 = UserForm1.TextBox22.Text
textvlue8 = UserForm1.TextBox10.Text
Dim rang7 As Range
Dim rang8 As Range
Set rang7 = Sheets("thariaba3").Range("zwjea3")
Set rang8 = Sheets("thariaba3").Range("astktaat3")

If UserForm1.TextBox23.Text = rang7 Then
UserForm1.TextBox24.value = Application.WorksheetFunction.VLookup(textvlue8, rang8, textvlue7 + 3, 1)
End If

End If

End If

End Sub

