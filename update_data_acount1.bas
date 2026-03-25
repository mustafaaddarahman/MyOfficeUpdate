Attribute VB_Name = "update_data_acount1"
Sub update_data_acount_22()
'On Error GoTo emad

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Dim lastR As Long ' ÇÓÊÎÏÇã Long ÃÓÑÚ ãä Integer æÃÖãä ááÃÚÏÇÏ ÇáßÈíÑÉ
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row

' ÇáÈÍË Úä ÇáÕİ ÇáãØáæÈ ÈäÇÁğ Úáì ÔÑæØß (ÇáÚãæÏ 5 æÇáÚãæÏ 2)
For Y = 9 To lastR
    If Sheets(1).Cells(Y, 5).Text = UserForm1.TextBox5.Text And Sheets(1).Cells(Y, 2).Text = UserForm1.TextBox2.Text Then
        Exit For
    End If
Next Y

' 2. ÊÑÍíá ÇáÈíÇäÇÊ ÏİÚÉ æÇÍÏÉ (ÇÓÊÎÏÇã With íÓÑÚ ÇáÚãáíÉ ÌÏÇğ)
With Sheets(1)
    .Cells(Y, 1) = UserForm1.TextBox1.Text
    .Cells(Y, 2) = UserForm1.TextBox2.Text
    .Cells(Y, 3) = UserForm1.TextBox3.Text
    .Cells(Y, 4) = UserForm1.ComboBox4.value
    .Cells(Y, 5) = UserForm1.TextBox5.Text
    .Cells(Y, 6) = UserForm1.ComboBox5.value
    .Cells(Y, 7) = UserForm1.TextBox7.Text
    .Cells(Y, 8) = UserForm1.TextBox8.Text
    .Cells(Y, 9) = UserForm1.TextBox9.Text
    
    ' ÇáÍİÇÙ Úáì ÔÑØ ßáãÉ "ÚŞÏ" ÈÇááÛÉ ÇáÚÑÈíÉ
    If UserForm1.ComboBox3.value = "ÚŞÏ" Or UserForm1.TextBox58.Text = "ÚŞÏ" Then
        .Cells(Y, 10) = UserForm1.TextBox101.Text
    Else
        .Cells(Y, 10) = UserForm1.TextBox10.Text
    End If
    
    .Cells(Y, 11) = UserForm1.TextBox11.Text
    .Cells(Y, 12) = UserForm1.TextBox12.Text
    .Cells(Y, 13) = UserForm1.ComboBox6.value
    .Cells(Y, 14) = UserForm1.TextBox14.Text
    .Cells(Y, 15) = UserForm1.ComboBox7.value
    .Cells(Y, 16) = UserForm1.TextBox16.Text
    .Cells(Y, 17) = UserForm1.ComboBox8.value
    .Cells(Y, 18) = UserForm1.TextBox18.Text
    .Cells(Y, 19) = UserForm1.TextBox19.Text
    .Cells(Y, 20) = UserForm1.TextBox20.Text
    .Cells(Y, 21) = UserForm1.TextBox21.Text
    .Cells(Y, 22) = UserForm1.TextBox22.Text
    .Cells(Y, 23) = UserForm1.TextBox23.Text
    .Cells(Y, 24) = UserForm1.TextBox24.Text
    .Cells(Y, 25) = UserForm1.TextBox25.Text
    .Cells(Y, 26) = UserForm1.ComboBox10.value
    .Cells(Y, 27) = UserForm1.TextBox27.Text
    .Cells(Y, 28) = UserForm1.ComboBox11.value
    .Cells(Y, 29) = UserForm1.TextBox29.Text
    .Cells(Y, 30) = UserForm1.TextBox30.Text
    .Cells(Y, 31) = UserForm1.TextBox31.Text
    .Cells(Y, 32) = UserForm1.TextBox32.Text
    .Cells(Y, 33) = UserForm1.TextBox33.Text
    .Cells(Y, 34) = UserForm1.TextBox34.Text
    .Cells(Y, 35) = UserForm1.TextBox35.Text
    .Cells(Y, 36) = UserForm1.TextBox36.Text
    .Cells(Y, 37) = UserForm1.TextBox37.Text
    .Cells(Y, 38) = UserForm1.TextBox38.Text
    .Cells(Y, 39) = UserForm1.TextBox39.Text
    .Cells(Y, 40) = UserForm1.TextBox40.Text
    .Cells(Y, 41) = UserForm1.TextBox41.Text
    .Cells(Y, 42) = UserForm1.TextBox42.Text
    .Cells(Y, 43) = UserForm1.TextBox43.Text
    .Cells(Y, 44) = UserForm1.TextBox44.Text
    .Cells(Y, 45) = UserForm1.TextBox45.Text
    .Cells(Y, 46) = UserForm1.TextBox46.Text
    .Cells(Y, 47) = UserForm1.TextBox47.Text
    .Cells(Y, 48) = UserForm1.TextBox48.Text
    .Cells(Y, 49) = UserForm1.TextBox49.Text
    .Cells(Y, 50) = UserForm1.TextBox50.Text
    .Cells(Y, 51) = UserForm1.ComboBox9.value
    .Cells(Y, 52) = UserForm1.TextBox52.Text
    .Cells(Y, 53) = UserForm1.TextBox53.Text
    .Cells(Y, 54) = UserForm1.TextBox54.Text
    .Cells(Y, 55) = UserForm1.TextBox55.Text
    .Cells(Y, 56) = UserForm1.TextBox56.Text
    .Cells(Y, 57) = UserForm1.TextBox57.Text
    .Cells(Y, 58) = UserForm1.TextBox58.Text
    .Cells(Y, 59) = UserForm1.TextBox59.Text
    .Cells(Y, 60) = UserForm1.TextBox60.Text
    .Cells(Y, 61) = UserForm1.TextBox61.Text
    .Cells(Y, 62) = UserForm1.TextBox62.Text
    .Cells(Y, 63) = UserForm1.TextBox63.Text
    .Cells(Y, 64) = UserForm1.TextBox64.Text
    
    ' ÅßãÇá ÇáÊÑÍíá ááÃÚãÏÉ ÇáÈÚíÏÉ
    .Cells(Y, "BN") = Format(UserForm1.TextBox82.Text, "yyyy/mm/dd")
    .Cells(Y, 67) = UserForm1.TextBox83.Text
    .Cells(Y, 68) = UserForm1.TextBox84.Text
    .Cells(Y, 69) = UserForm1.TextBox85.Text
    .Cells(Y, 70) = UserForm1.TextBox86.Text
    .Cells(Y, 71) = UserForm1.TextBox87.Text
    .Cells(Y, 72) = UserForm1.TextBox88.Text
    .Cells(Y, 73) = UserForm1.TextBox64.Text
    .Cells(Y, 74) = UserForm1.TextBox91.Text
    .Cells(Y, "BU") = UserForm1.TextBox99.Text
    .Cells(Y, "CE") = UserForm1.TextBox105.Text
    .Cells(Y, "CA") = UserForm1.TextBox100.Text

    ' 3. ÊãííÒ ÇáÊßÑÇÑ İí ÇáÚãæÏ 5 (E) ÈÇááæä ÇáÃÕİÑ İæÑÇğ
    If WorksheetFunction.CountIf(.Range("E9:E" & lastR + 1), UserForm1.TextBox5.Text) > 1 Then
        .Cells(Y, 5).Interior.Color = vbYellow
    Else
        .Cells(Y, 5).Interior.ColorIndex = xlNone
    End If
End With

' 4. ÅÚÇÏÉ ÊİÚíá ÇáÅÚÏÇÏÇÊ (ÖÑæÑí ÌÏÇğ áíÑÌÚ ÇáÅßÓá áØÈíÚÊå)
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True





'=======================================================









End Sub
