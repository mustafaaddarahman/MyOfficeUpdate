VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm68 
   Caption         =   "UserForm68"
   ClientHeight    =   5868
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   22236
   OleObjectBlob   =   "UserForm68.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm68"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' 1. «·Ê’› «·ÊŸÌ›Ì (CF) -> «·⁄œœ ›Ì TextBox1 Ê≈Ã„«·Ì «·‹ 50% „‰ «·—« » «·ﬂ·Ì (AW) ›Ì TextBox14
Private Sub ComboBox1_Change()
    Dim i As Long, lr As Long, count As Long: Dim sumAmt As Double
    lr = Sheets(1).Cells(Rows.count, "CF").End(xlUp).row
    count = 0: sumAmt = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "CF").value)) = Trim(ComboBox1.value) Then
            count = count + 1
            ' Õ”«» 50% „‰ «·—« » «·ﬂ·Ì (AW) ·ﬂ· „ÊŸ› „ÿ«»ﬁ
            sumAmt = sumAmt + (val(Sheets(1).Cells(i, "AW").value) * 0.5)
        End If
    Next i
    TextBox1.value = count: TextBox14.value = Format(sumAmt, "#,##0")
End Sub

' 2. «·‘Â«œ… (M) -> «·⁄œœ ›Ì TextBox2 Ê„Ã„Ê⁄ „Œ’’« Â« „‰ «·—« » «·ﬂ·Ì (AW) ›Ì TextBox10
Private Sub ComboBox2_Change()
    Dim i As Long, lr As Long, count As Long: Dim sumAmt As Double, targetVal As Double
    lr = Sheets(1).Cells(Rows.count, "CF").End(xlUp).row
    Select Case ComboBox2.value
        Case "„ Ê”ÿÂ": targetVal = 0.15: Case "«⁄œ«œÌ…": targetVal = 0.25: Case "œ»·Ê„": targetVal = 0.35
        Case "»ﬂ·Ê—ÌÊ”": targetVal = 0.45: Case "œ»·Ê„ ⁄«·Ì": targetVal = 0.6: Case "„«Ã” Ì—": targetVal = 0.75
        Case "œﬂ Ê—«Â": targetVal = 1#: Case "„«Ã” Ì— „⁄ „Œ’’«  Ã«„⁄Ì…": targetVal = 1.25: Case "œﬂ Ê—«Â „⁄ „Œ’’«  Ã«„⁄Ì…": targetVal = 1.5
    End Select
    count = 0: sumAmt = 0
    For i = 9 To lr
        If Abs(val(Sheets(1).Cells(i, "M").value) - targetVal) < 0.001 Then
            count = count + 1
            ' ÷—» ‰”»… «·‘Â«œ… ›Ì «·—« » «·ﬂ·Ì (AW)
            sumAmt = sumAmt + (val(Sheets(1).Cells(i, "AW").value) * targetVal)
        End If
    Next i
    TextBox2.value = count: TextBox10.value = Format(sumAmt, "#,##0")
End Sub

' 3. «·„Œ’’«  «·„Â‰Ì… (Z) -> «·⁄œœ ›Ì TextBox3 Ê„Ã„Ê⁄Â« „‰ «·—« » «·ﬂ·Ì ›Ì TextBox11
Private Sub ComboBox3_Change()
    Dim i As Long, lr As Long, count As Long: Dim sumAmt As Double
    lr = Sheets(1).Cells(Rows.count, "Z").End(xlUp).row
    count = 0: sumAmt = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "Z").value)) = Trim(ComboBox3.value) Then
            count = count + 1
            ' «” Œ—«Ã «·‰”»… „‰ Z Ê÷—»Â« ›Ì AW
            sumAmt = sumAmt + (val(Sheets(1).Cells(i, "AW").value) * val(Sheets(1).Cells(i, "Z").value))
        End If
    Next i
    TextBox3.value = count: TextBox11.value = Format(sumAmt, "#,##0")
End Sub

' 4. «·„Êﬁ› «·ÊŸÌ›Ì (BF) -> TextBox4
Private Sub ComboBox4_Change()
    Dim i As Long, lr As Long, count As Long
    lr = Sheets(1).Cells(Rows.count, "BF").End(xlUp).row
    count = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "BF").value)) = Trim(ComboBox4.value) Then count = count + 1
    Next i
    TextBox4.value = count
End Sub

' 5. ÿ—Ìﬁ… «·œ›⁄ (AY) -> TextBox5
Private Sub ComboBox5_Change()
    Dim i As Long, lr As Long, count As Long
    lr = Sheets(1).Cells(Rows.count, "AY").End(xlUp).row
    count = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "AY").value)) = Trim(ComboBox5.value) Then count = count + 1
    Next i
    TextBox5.value = count
End Sub

' 6. «·Õ«·… «·«Ã „«⁄Ì… (W) -> TextBox6
Private Sub ComboBox6_Change()
    Dim i As Long, lr As Long, count As Long
    lr = Sheets(1).Cells(Rows.count, "W").End(xlUp).row
    count = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "W").value)) = Trim(ComboBox6.value) Then count = count + 1
    Next i
    TextBox6.value = count
End Sub

' 7. «·„‰’» (Q) -> «·⁄œœ ›Ì TextBox7 Ê„Ã„Ê⁄ „»«·€ «·„‰’» „‰ «·—« » ›Ì TextBox12
Private Sub ComboBox7_Change()
    Dim i As Long, lr As Long, count As Long: Dim sumAmt As Double
    lr = Sheets(1).Cells(Rows.count, "Q").End(xlUp).row
    count = 0: sumAmt = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "Q").value)) = Trim(ComboBox7.value) Then
            count = count + 1
            sumAmt = sumAmt + (val(Sheets(1).Cells(i, "AW").value) * val(Sheets(1).Cells(i, "Q").value))
        End If
    Next i
    TextBox7.value = count: TextBox12.value = Format(sumAmt, "#,##0")
End Sub

' 8. „Œ’’«  Â‰œ”Ì… (O) -> «·⁄œœ ›Ì TextBox8 Ê«·„Ã„Ê⁄ „‰ «·—« » ›Ì TextBox13
Private Sub ComboBox8_Change()
    Dim i As Long, lr As Long, count As Long: Dim sumAmt As Double
    lr = Sheets(1).Cells(Rows.count, "O").End(xlUp).row
    count = 0: sumAmt = 0
    For i = 9 To lr
        If Trim(CStr(Sheets(1).Cells(i, "O").value)) = Trim(ComboBox8.value) Then
            count = count + 1
            sumAmt = sumAmt + (val(Sheets(1).Cells(i, "AW").value) * val(Sheets(1).Cells(i, "O").value))
        End If
    Next i
    TextBox8.value = count: TextBox13.value = Format(sumAmt, "#,##0")
End Sub

' 9. «·œ—Ã«  (H) -> TextBox9 (8° 9° 10 √Ê «·ﬂ·)
Private Sub ComboBox9_Change()
    Dim i As Long, lr As Long, count As Long
    Dim sumAW As Double, cellVal As Double
    
    '  ÕœÌœ ¬Œ— ’› ›Ì «·⁄„Êœ H
    lr = Sheets(1).Cells(Rows.count, "H").End(xlUp).row
    count = 0
    sumAW = 0
    
    ' «· √ﬂœ „‰ ÊÃÊœ «Œ Ì«— ›Ì «·ﬂÊ„»Ê »Êﬂ”
    If ComboBox9.value = "" Then Exit Sub
    
    For i = 9 To lr
        cellVal = val(Sheets(1).Cells(i, "H").value)
        
        ' «·Õ«·… «·√Ê·Ï: «Œ Ì«— «·ﬂ· (8 Ê 9 Ê 10)
        If ComboBox9.value = "«·ﬂ· (8+9+10)" Then
            If cellVal = 8 Or cellVal = 9 Or cellVal = 10 Then
                count = count + 1
                sumAW = sumAW + val(Sheets(1).Cells(i, "AW").value)
            End If
            
        ' «·Õ«·… «·À«‰Ì…: «Œ Ì«— œ—Ã… „Õœœ… Ê«Õœ… ›ﬁÿ
        Else
            If cellVal = val(ComboBox9.value) Then
                count = count + 1
                sumAW = sumAW + val(Sheets(1).Cells(i, "AW").value)
            End If
        End If
    Next i
    
    ' ⁄—÷ «·‰ «∆Ã
    TextBox9.value = count ' ⁄œœ «·„ÊŸ›Ì‰
    ' ·‰› —÷ √‰ TextBox15 ÂÊ «·„Œ’’ ·„Ã„Ê⁄ „»«·€ «·œ—Ã«  «·„Œ «—…
    TextBox14.value = Format(sumAW, "#,##0")
End Sub

Private Sub CommandButton1_Click()
 Call ExecuteStatistics(Me)
    ' ⁄—÷  ‰»ÌÂ ”—Ì⁄
    MsgBox " „  ÕœÌÀ «·≈Õ’«∆Ì«   »‰Ã«Õ", vbInformation, " ÕœÌÀ"
End Sub

Private Sub CommandButton2_Click()
Unload Me
UserForm9.Show
End Sub



Private Sub CommandButton3_Click()
    Dim wsSource As Worksheet: Set wsSource = Sheets(1)
    Dim wsReport As Worksheet
    Dim i As Long, lr As Long, NextRow As Long, col As Long
    Dim targetM As Double, cellM As Double, cellH As Double
    Dim isMatch As Boolean
    Dim rngTable As Range
    Dim selectedCols As Variant
    
    ' 1.  ÕœÌœ «·√⁄„œ… «·√”«”Ì… («·—ﬁ„ «·ÊŸÌ›Ì B° «·„«” — C° «·«”„ E)
    selectedCols = Array("B", "C", "E")
    
    ' ≈÷«›… √⁄„œ… «·‘—Êÿ «·„Œ «—… ›ﬁÿ ≈·Ï „’›Ê›… «· ﬁ—Ì—
    If Me.ComboBox1.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",CF", ",")
    If Me.ComboBox2.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",M", ",")
    If Me.ComboBox3.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",Z", ",")
    If Me.ComboBox4.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",BF", ",")
    If Me.ComboBox5.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",AY", ",")
    If Me.ComboBox6.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",W", ",")
    If Me.ComboBox7.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",Q", ",")
    If Me.ComboBox8.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",O", ",")
    If Me.ComboBox9.value <> "" Then selectedCols = Split(Join(selectedCols, ",") & ",H", ",")
    
    ' ≈÷«›… «·—« » «·’«›Ì (AW) œ«∆„« ›Ì ‰Â«Ì… «·ÃœÊ·
    selectedCols = Split(Join(selectedCols, ",") & ",AW", ",")

    ' 2. ≈⁄œ«œ ‘Ì  «· ﬁ—Ì— «·„ƒﬁ  (Ê÷⁄Â »⁄œ ‘Ì  BlackBox)
    On Error Resume Next
    Application.DisplayAlerts = False: Sheets("Temp_Report").Delete: Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' «· ⁄œÌ· Â‰«:  ÕœÌœ „ﬂ«‰ «·‘Ì  «·ÃœÌœ »⁄œ BlackBox
    Set wsReport = Sheets.Add(After:=Sheets("BlackBox"))
    wsReport.Name = "Temp_Report"
    wsReport.DisplayRightToLeft = True

    ' Ê÷⁄ «·⁄‰«ÊÌ‰ ›Ì ‘Ì  «· ﬁ—Ì— „‰ «·’› 8 ›Ì «·‘Ì  «·√’·Ì
    For col = 0 To UBound(selectedCols)
        wsReport.Cells(1, col + 1).value = wsSource.Cells(8, selectedCols(col)).value
    Next col

    lr = wsSource.Cells(wsSource.Rows.count, "CF").End(xlUp).row
    NextRow = 2
    
    ' 3.  ÕœÌœ ﬁÌ„… «·‘Â«œ… (M)
    targetM = 0
    Select Case Me.ComboBox2.value
        Case "„ Ê”ÿÂ": targetM = 0.15: Case "«⁄œ«œÌ…": targetM = 0.25: Case "œ»·Ê„": targetM = 0.35
        Case "»ﬂ·Ê—ÌÊ”": targetM = 0.45: Case "œ»·Ê„ ⁄«·Ì": targetM = 0.6: Case "„«Ã” Ì—": targetM = 0.75: Case "œﬂ Ê—«Â": targetM = 1#: Case "„«Ã” Ì— „⁄ „Œ’’«  Ã«„⁄Ì…": targetM = 1.25: Case "œﬂ Ê—«Â „⁄ „Œ’’«  Ã«„⁄Ì…": targetM = 1.5
    End Select

    ' 4. «·›· —… Ê‰ﬁ· «·»Ì«‰«  «·„Œ «—… ›ﬁÿ ﬁÌœ« »ﬁÌœ
    For i = 9 To lr
        isMatch = True
        cellM = val(wsSource.Cells(i, "M").value): cellH = val(wsSource.Cells(i, "H").value)

        ' «· Õﬁﬁ „‰ «·‘—Êÿ ( Ã«Â· «·ﬂÊ„»Ê »Êﬂ” «·›«—€)
        If Me.ComboBox1.value <> "" Then If Trim(CStr(wsSource.Cells(i, "CF").value)) <> Trim(Me.ComboBox1.value) Then isMatch = False
        If isMatch And Me.ComboBox2.value <> "" Then If Abs(cellM - targetM) > 0.001 Then isMatch = False
        If isMatch And Me.ComboBox3.value <> "" Then If Trim(CStr(wsSource.Cells(i, "Z").value)) <> Trim(Me.ComboBox3.value) Then isMatch = False
        If isMatch And Me.ComboBox4.value <> "" Then If Trim(CStr(wsSource.Cells(i, "BF").value)) <> Trim(Me.ComboBox4.value) Then isMatch = False
        If isMatch And Me.ComboBox5.value <> "" Then If Trim(CStr(wsSource.Cells(i, "AY").value)) <> Trim(Me.ComboBox5.value) Then isMatch = False
        If isMatch And Me.ComboBox6.value <> "" Then If Trim(CStr(wsSource.Cells(i, "W").value)) <> Trim(Me.ComboBox6.value) Then isMatch = False
        If isMatch And Me.ComboBox7.value <> "" Then If Trim(CStr(wsSource.Cells(i, "Q").value)) <> Trim(Me.ComboBox7.value) Then isMatch = False
        If isMatch And Me.ComboBox8.value <> "" Then If Trim(CStr(wsSource.Cells(i, "O").value)) <> Trim(Me.ComboBox8.value) Then isMatch = False
        If isMatch And Me.ComboBox9.value <> "" Then
            If Me.ComboBox9.value = "«·ﬂ· (8+9+10)" Then
                If cellH < 8 Or cellH > 10 Then isMatch = False
            Else: If cellH <> val(Me.ComboBox9.value) Then isMatch = False
            End If
        End If

        ' ‰ﬁ· «·»Ì«‰«  «·„›· —… ›ﬁÿ
        If isMatch Then
            For col = 0 To UBound(selectedCols)
                wsReport.Cells(NextRow, col + 1).value = wsSource.Cells(i, selectedCols(col)).value
            Next col
            NextRow = NextRow + 1
        End If
    Next i

    ' 5. «· ‰”Ìﬁ° ÷»ÿ «·’›Õ…° Ê«·„⁄«Ì‰…
    If NextRow > 2 Then
        Set rngTable = wsReport.Range("A1").Resize(NextRow - 1, UBound(selectedCols) + 1)
        
        ' ≈⁄œ«œ«  «·’›Õ… (√›ﬁÌ Landscape)
        With wsReport.PageSetup
            .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1
            .CenterHeader = "&""Arial,Bold""&16  ﬁ—Ì— «·„ÊŸ›Ì‰ «·„›· —"
            .CenterFooter = " «—ÌŒ «· ﬁ—Ì—: " & Date & " - ’›Õ… &P „‰ &N"
        End With

        '  ‰”Ìﬁ «·ÃœÊ·
        With rngTable
            .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
            .HorizontalAlignment = xlCenter: .Font.Name = "Arial": .Columns.AutoFit
        End With
        
        '  „ÌÌ“ —√” «·ÃœÊ·
        With wsReport.Rows(1)
            .Interior.Color = RGB(220, 220, 220): .Font.Bold = True: .RowHeight = 25
        End With
        
        ' › Õ »Ì∆… ≈ﬂ”Ì· Ê«·—Ì»Ê‰ ··„⁄«Ì‰…
        Me.Hide
        Application.Visible = True
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        
        wsReport.PrintPreview
        
        ' ≈€·«ﬁ »Ì∆… ≈ﬂ”Ì· Ê«·—Ì»Ê‰ Ê«·⁄Êœ… ··›Ê—„
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        Application.Visible = False
        Me.Show
    Else
        MsgBox "·«  ÊÃœ ‰ «∆Ã  ÿ«»ﬁ Â–Â «·‘—Êÿ.", vbExclamation, " ‰»ÌÂ"
    End If
    
    ' Õ–› ‘Ì  «· ﬁ—Ì— «·„ƒﬁ  »⁄œ «·«‰ Â«¡
    Application.DisplayAlerts = False: wsReport.Delete: Application.DisplayAlerts = True
End Sub

' --- ﬂÊœ  Õ„Ì· «·»Ì«‰«  Ê ⁄»∆… «·ﬁÊ«∆„ (Initialize) ---
Private Sub UserForm_Initialize()
    Dim i As Long, lr As Long
    Dim d1 As Object, d3 As Object, d4 As Object, d5 As Object, d6 As Object, d7 As Object, d8 As Object
    Set d1 = CreateObject("Scripting.Dictionary"): Set d3 = CreateObject("Scripting.Dictionary")
    Set d4 = CreateObject("Scripting.Dictionary"): Set d5 = CreateObject("Scripting.Dictionary")
    Set d6 = CreateObject("Scripting.Dictionary"): Set d7 = CreateObject("Scripting.Dictionary")
    Set d8 = CreateObject("Scripting.Dictionary")
    
    lr = Sheets(1).Cells(Rows.count, "CF").End(xlUp).row
    
    ComboBox2.List = Array("„ Ê”ÿÂ", "«⁄œ«œÌ…", "œ»·Ê„", "»ﬂ·Ê—ÌÊ”", "œ»·Ê„ ⁄«·Ì", "„«Ã” Ì—", "œﬂ Ê—«Â", "„«Ã” Ì— „⁄ „Œ’’«  Ã«„⁄Ì…", "œﬂ Ê—«Â „⁄ „Œ’’«  Ã«„⁄Ì…")
    ComboBox9.List = Array("8", "9", "10", "«·ﬂ· (8+9+10)")
    
    For i = 9 To lr
        If Trim(Sheets(1).Cells(i, "CF").value) <> "" And Not d1.exists(Trim(Sheets(1).Cells(i, "CF").value)) Then d1.Add Trim(Sheets(1).Cells(i, "CF").value), "": ComboBox1.AddItem Trim(Sheets(1).Cells(i, "CF").value)
        If Trim(Sheets(1).Cells(i, "Z").value) <> "" And Not d3.exists(Trim(Sheets(1).Cells(i, "Z").value)) Then d3.Add Trim(Sheets(1).Cells(i, "Z").value), "": ComboBox3.AddItem Trim(Sheets(1).Cells(i, "Z").value)
        If Trim(Sheets(1).Cells(i, "BF").value) <> "" And Not d4.exists(Trim(Sheets(1).Cells(i, "BF").value)) Then d4.Add Trim(Sheets(1).Cells(i, "BF").value), "": ComboBox4.AddItem Trim(Sheets(1).Cells(i, "BF").value)
        If Trim(Sheets(1).Cells(i, "AY").value) <> "" And Not d5.exists(Trim(Sheets(1).Cells(i, "AY").value)) Then d5.Add Trim(Sheets(1).Cells(i, "AY").value), "": ComboBox5.AddItem Trim(Sheets(1).Cells(i, "AY").value)
        If Trim(Sheets(1).Cells(i, "W").value) <> "" And Not d6.exists(Trim(Sheets(1).Cells(i, "W").value)) Then d6.Add Trim(Sheets(1).Cells(i, "W").value), "": ComboBox6.AddItem Trim(Sheets(1).Cells(i, "W").value)
        If Trim(Sheets(1).Cells(i, "Q").value) <> "" And Not d7.exists(Trim(Sheets(1).Cells(i, "Q").value)) Then d7.Add Trim(Sheets(1).Cells(i, "Q").value), "": ComboBox7.AddItem Trim(Sheets(1).Cells(i, "Q").value)
        If Trim(Sheets(1).Cells(i, "O").value) <> "" And Not d8.exists(Trim(Sheets(1).Cells(i, "O").value)) Then d8.Add Trim(Sheets(1).Cells(i, "O").value), "": ComboBox8.AddItem Trim(Sheets(1).Cells(i, "O").value)
    Next i
End Sub








