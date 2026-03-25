Attribute VB_Name = "sumetion_all_datain_butom1"
Option Explicit

Sub ExecuteStatistics(frm As Object)
    Dim i As Long, lr As Long
    Dim sumGrandAW As Double, currentAW As Double
    Dim targetM As Double, cellM As Double, cellH As Double
    
    '  ÕœÌœ ¬Œ— ’› ›Ì «·‘Ì  —ﬁ„ 1
    lr = Sheets(1).Cells(Rows.count, "CF").End(xlUp).row
    If lr < 9 Then Exit Sub

    '  ÕœÌœ ‰”»… «·‘Â«œ… ·‹ ComboBox2
    targetM = 0
    Select Case frm.ComboBox2.value
        Case "„ Ê”ÿÂ": targetM = 0.15: Case "«⁄œ«œÌ…": targetM = 0.25: Case "œ»·Ê„": targetM = 0.35
        Case "»ﬂ·Ê—ÌÊ”": targetM = 0.45: Case "œ»·Ê„ ⁄«·Ì": targetM = 0.6: Case "„«Ã” Ì—": targetM = 0.75
        Case "œﬂ Ê—«Â": targetM = 1#: Case "„«Ã” Ì— „⁄ „Œ’’«  Ã«„⁄Ì…": targetM = 1.25: Case "œﬂ Ê—«Â „⁄ „Œ’’«  Ã«„⁄Ì…": targetM = 1.5
    End Select

    sumGrandAW = 0

    ' œÊ—…  ﬂ—«— Ê«Õœ…  „— ⁄·Ï ﬂ«›… «·ﬁÌÊœ
    For i = 9 To lr
        currentAW = val(Sheets(1).Cells(i, "AW").value)
        cellM = val(Sheets(1).Cells(i, "M").value)
        cellH = val(Sheets(1).Cells(i, "H").value)

        ' --- Ã„⁄ „»«·€ ﬂ· «Œ Ì«— »‘ﬂ· „” ﬁ· (√Ê/Or Logic) ---

        ' 1. ≈–« ﬂ«‰ «·’› Ìÿ«»ﬁ «·ﬂÊ„»Ê 1 («·Ê’›)
        If frm.ComboBox1.value <> "" Then
            If Trim(CStr(Sheets(1).Cells(i, "CF").value)) = Trim(frm.ComboBox1.value) Then
                sumGrandAW = sumGrandAW + currentAW
            End If
        End If

        ' 2. √÷› „»«·€ «·ﬂÊ„»Ê 2 («·‘Â«œ…)
        If frm.ComboBox2.value <> "" Then
            If Abs(cellM - targetM) < 0.001 Then
                sumGrandAW = sumGrandAW + currentAW
            End If
        End If

        ' 3. √÷› „»«·€ «·ﬂÊ„»Ê 3 («·„Œ’’«  Z)
        If frm.ComboBox3.value <> "" Then
            If Trim(CStr(Sheets(1).Cells(i, "Z").value)) = Trim(frm.ComboBox3.value) Then
                sumGrandAW = sumGrandAW + currentAW
            End If
        End If

        ' 7. √÷› „»«·€ «·ﬂÊ„»Ê 7 («·„‰’» Q)
        If frm.ComboBox7.value <> "" Then
            If Trim(CStr(Sheets(1).Cells(i, "Q").value)) = Trim(frm.ComboBox7.value) Then
                sumGrandAW = sumGrandAW + currentAW
            End If
        End If

        ' 8. √÷› „»«·€ «·ﬂÊ„»Ê 8 («·Â‰œ”Ì… O)
        If frm.ComboBox8.value <> "" Then
            If Trim(CStr(Sheets(1).Cells(i, "O").value)) = Trim(frm.ComboBox8.value) Then
                sumGrandAW = sumGrandAW + currentAW
            End If
        End If

        ' 9. √÷› „»«·€ «·ﬂÊ„»Ê 9 («·œ—Ã«  H)
        If frm.ComboBox9.value <> "" Then
            If frm.ComboBox9.value = "«·ﬂ· (8+9+10)" Then
                If cellH >= 8 And cellH <= 10 Then sumGrandAW = sumGrandAW + currentAW
            Else
                If cellH = val(frm.ComboBox9.value) Then sumGrandAW = sumGrandAW + currentAW
            End If
        End If
    Next i

    ' ⁄—÷ «·‰« Ã «·‰Â«∆Ì «·„‰”ﬁ ›Ì TextBox15
    frm.TextBox15.value = Format(sumGrandAW, "#,##0")
End Sub
