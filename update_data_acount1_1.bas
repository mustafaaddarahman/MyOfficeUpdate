Attribute VB_Name = "update_data_acount1_1"
Sub update_data_acount_1()
'On Error GoTo emad


Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

Dim lastR As Long ' ÇÓÊÎÏÇã Long ÃÓÑÚ ãä Integer æÃÖãä ááÃÚÏÇÏ ÇáßÈíÑÉ
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row

' ÇáÈÍË Úä ÇáÕİ ÇáãØáæÈ ÈäÇÁğ Úáì ÔÑæØß (ÇáÚãæÏ 5 æÇáÚãæÏ 2)
For Y = 9 To lastR
    If Sheets(1).Cells(Y, 5).Text = UserForm48.TextBox5.Text And Sheets(1).Cells(Y, 2).Text = UserForm48.TextBox2.Text Then
        Exit For
    End If
Next Y

' 2. ÊÑÍíá ÇáÈíÇäÇÊ ÏİÚÉ æÇÍÏÉ (ÇÓÊÎÏÇã With íÓÑÚ ÇáÚãáíÉ ÌÏÇğ)
With Sheets(1)
    .Cells(Y, 1) = UserForm48.TextBox1.Text
    .Cells(Y, 2) = UserForm48.TextBox2.Text
    .Cells(Y, 3) = UserForm48.TextBox3.Text
    .Cells(Y, 4) = UserForm48.ComboBox4.value
    .Cells(Y, 5) = UserForm48.TextBox5.Text
    .Cells(Y, 6) = UserForm48.ComboBox5.value
    .Cells(Y, 7) = UserForm48.TextBox7.Text
    .Cells(Y, 8) = UserForm48.TextBox8.Text
    .Cells(Y, 9) = UserForm48.TextBox9.Text
    
    ' ÇáÍİÇÙ Úáì ÔÑØ ßáãÉ "ÚŞÏ" ÈÇááÛÉ ÇáÚÑÈíÉ
    If UserForm48.ComboBox3.value = "ÚŞÏ" Or UserForm48.TextBox58.Text = "ÚŞÏ" Then
        .Cells(Y, 10) = UserForm48.TextBox101.Text
    Else
        .Cells(Y, 10) = UserForm48.TextBox10.Text
    End If
    
    .Cells(Y, 11) = UserForm48.TextBox11.Text
    .Cells(Y, 12) = UserForm48.TextBox12.Text
    .Cells(Y, 13) = UserForm48.ComboBox6.value
    .Cells(Y, 14) = UserForm48.TextBox14.Text
    .Cells(Y, 15) = UserForm48.ComboBox7.value
    .Cells(Y, 16) = UserForm48.TextBox16.Text
    .Cells(Y, 17) = UserForm48.ComboBox8.value
    .Cells(Y, 18) = UserForm48.TextBox18.Text
    .Cells(Y, 19) = UserForm48.TextBox19.Text
    .Cells(Y, 20) = UserForm48.TextBox20.Text
    .Cells(Y, 21) = UserForm48.TextBox21.Text
    .Cells(Y, 22) = UserForm48.TextBox22.Text
    .Cells(Y, 23) = UserForm48.TextBox23.Text
    .Cells(Y, 24) = UserForm48.TextBox24.Text
    .Cells(Y, 25) = UserForm48.TextBox25.Text
    .Cells(Y, 26) = UserForm48.ComboBox10.value
    .Cells(Y, 27) = UserForm48.TextBox27.Text
    .Cells(Y, 28) = UserForm48.ComboBox11.value
    .Cells(Y, 29) = UserForm48.TextBox29.Text
    .Cells(Y, 30) = UserForm48.TextBox30.Text
    .Cells(Y, 31) = UserForm48.TextBox31.Text
    .Cells(Y, 32) = UserForm48.TextBox32.Text
    .Cells(Y, 33) = UserForm48.TextBox33.Text
    .Cells(Y, 34) = UserForm48.TextBox34.Text
    .Cells(Y, 35) = UserForm48.TextBox35.Text
    .Cells(Y, 36) = UserForm48.TextBox36.Text
    .Cells(Y, 37) = UserForm48.TextBox37.Text
    .Cells(Y, 38) = UserForm48.TextBox38.Text
    .Cells(Y, 39) = UserForm48.TextBox39.Text
    .Cells(Y, 40) = UserForm48.TextBox40.Text
    .Cells(Y, 41) = UserForm48.TextBox41.Text
    .Cells(Y, 42) = UserForm48.TextBox42.Text
    .Cells(Y, 43) = UserForm48.TextBox43.Text
    .Cells(Y, 44) = UserForm48.TextBox44.Text
    .Cells(Y, 45) = UserForm48.TextBox45.Text
    .Cells(Y, 46) = UserForm48.TextBox46.Text
    .Cells(Y, 47) = UserForm48.TextBox47.Text
    .Cells(Y, 48) = UserForm48.TextBox48.Text
    .Cells(Y, 49) = UserForm48.TextBox49.Text
    .Cells(Y, 50) = UserForm48.TextBox50.Text
    .Cells(Y, 51) = UserForm48.ComboBox9.value
    .Cells(Y, 52) = UserForm48.TextBox52.Text
    .Cells(Y, 53) = UserForm48.TextBox53.Text
    .Cells(Y, 54) = UserForm48.TextBox54.Text
    .Cells(Y, 55) = UserForm48.TextBox55.Text
    .Cells(Y, 56) = UserForm48.TextBox56.Text
    .Cells(Y, 57) = UserForm48.TextBox57.Text
    .Cells(Y, 58) = UserForm48.TextBox58.Text
    .Cells(Y, 59) = UserForm48.TextBox59.Text
    .Cells(Y, 60) = UserForm48.TextBox60.Text
    .Cells(Y, 61) = UserForm48.TextBox61.Text
    .Cells(Y, 62) = UserForm48.TextBox62.Text
    .Cells(Y, 63) = UserForm48.TextBox63.Text
    .Cells(Y, 64) = UserForm48.TextBox64.Text
    
    ' ÅßãÇá ÇáÊÑÍíá ááÃÚãÏÉ ÇáÈÚíÏÉ
    .Cells(Y, "BN") = Format(UserForm48.TextBox82.Text, "yyyy/mm/dd")
    .Cells(Y, 67) = UserForm48.TextBox83.Text
    .Cells(Y, 68) = UserForm48.TextBox84.Text
    .Cells(Y, 69) = UserForm48.TextBox85.Text
    .Cells(Y, 70) = UserForm48.TextBox86.Text
    .Cells(Y, 71) = UserForm48.TextBox87.Text
    .Cells(Y, 72) = UserForm48.TextBox88.Text
    .Cells(Y, 73) = UserForm48.TextBox64.Text
    .Cells(Y, 74) = UserForm48.TextBox91.Text
    .Cells(Y, "BU") = UserForm48.TextBox99.Text
    .Cells(Y, "CE") = UserForm48.TextBox105.Text
    .Cells(Y, "CA") = UserForm48.TextBox100.Text

    ' 3. ÊãííÒ ÇáÊßÑÇÑ İí ÇáÚãæÏ 5 (E) ÈÇááæä ÇáÃÕİÑ İæÑÇğ
    If WorksheetFunction.CountIf(.Range("E9:E" & lastR + 1), UserForm48.TextBox5.Text) > 1 Then
        .Cells(Y, 5).Interior.Color = vbYellow
    Else
        .Cells(Y, 5).Interior.ColorIndex = xlNone
    End If
End With

' 4. ÅÚÇÏÉ ÊİÚíá ÇáÅÚÏÇÏÇÊ (ÖÑæÑí ÌÏÇğ áíÑÌÚ ÇáÅßÓá áØÈíÚÊå)
Application.Calculation = xlCalculationAutomatic
Application.EnableEvents = True
Application.ScreenUpdating = True




End Sub
