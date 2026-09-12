Attribute VB_Name = "true_salary"
Sub true_salary_1()


    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    '  ⁄—Ì› „ €Ì—«  «·Õ”«»«  «·√”«”Ì…
    Dim totalEntitly As Double      ' «·«” Õﬁ«ﬁ «·ﬂ·Ì (UserForm1.TextBox31)
    Dim absenceDays As Double       ' √Ì«„ «·€Ì«» (UserForm1.TextBox11)
    Dim dailyRate As Double         ' ﬁÌ„… €Ì«» «·ÌÊ„ «·Ê«Õœ
    Dim calculatedAbsence As Double ' „»·€ «·€Ì«» «·„Õ”Ê» „»œ∆Ì«
    Dim netSalaryBox As Double      ' ’«›Ì «·—« » «·Ÿ«Â— ›Ì «·›Ê—„ (UserForm1.TextBox49)
    
    '  ÕÊÌ· «·ﬁÌ„ „‰ «·‹ UserForm „⁄ —»ÿÂ« »«”„ «·›Ê—„ «·’—ÌÕ
    absenceDays = val(UserForm1.TextBox11.value)
    totalEntitly = val(UserForm1.TextBox31.value)
    netSalaryBox = val(UserForm1.TextBox49.value)
    
    ' √Ê·«:  ÕœÌÀ  ·ﬁ«∆Ì ·‹ TextBox12 (30 - √Ì«„ «·€Ì«»)
    If UserForm1.TextBox11.value <> "" Then
        UserForm1.TextBox12.value = 30 - absenceDays
    Else
        UserForm1.TextBox12.value = ""
    End If
    
    ' À«‰Ì«: «Õ ”«» „»·€ «·€Ì«» «·«› —«÷Ì: («·«” Õﬁ«ﬁ / 30) * ⁄œœ «·√Ì«„
    If absenceDays > 0 And totalEntitly > 0 Then
        dailyRate = totalEntitly / 30
        calculatedAbsence = dailyRate * absenceDays
    Else
        calculatedAbsence = 0
    End If
    
    ' À«·À«: Œ—Ìÿ… «·ŒÊ«—“„Ì… Ê«·„Ê«“‰… (Balance) ⁄‰œ  Ã«Ê“ «·’«›Ì
    If calculatedAbsence > netSalaryBox And calculatedAbsence > 0 Then
        
        '  €ÌÌ— «·„”«—: «·»ÕÀ ›Ì Sheet1 »‰«¡ ⁄·Ï «·—ﬁ„ «·ÊŸÌ›Ì Ê«·«”„
        Dim searchId As String
        Dim searchName As String
        searchId = Trim(UserForm1.TextBox2.value)   ' «·—ﬁ„ «·ÊŸÌ›Ì («·⁄„Êœ B)
        searchName = Trim(UserForm1.TextBox5.value) ' «”„ «·„ÊŸ› («·⁄„Êœ E)
        
        Dim lastRow As Long
        Dim i As Long
        Dim foundMatch As Boolean
        Dim sheetNetSalary As Double
        
        lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
        foundMatch = False
        
        ' Õ·ﬁ…  »œ√ „‰ «·’› 9 ·„ÿ«»ﬁ… «·»Ì«‰« 
        For i = 9 To lastRow
            If Trim(ws.Cells(i, "B").value) = searchId And Trim(ws.Cells(i, "E").value) = searchName Then
                sheetNetSalary = val(ws.Cells(i, "AW").value) ' Ã·» «·’«›Ì „‰ «·⁄„Êœ AW
                foundMatch = True
                Exit For
            End If
        Next i
        
        '  ÿ»Ìﬁ ‘—Êÿ «·„Ê«“‰… („⁄  ÕÊÌ· «·‰« Ã ≈·Ï ⁄œœ ’ÕÌÕ »œÊ‰ ﬂ”Ê— "0")
        If foundMatch Then
            If sheetNetSalary < calculatedAbsence Then
                ' «·„Ê«“‰…: Ã⁄· «·€Ì«» Ì”«ÊÌ «·’«›Ì ﬂ⁄œœ ’ÕÌÕ
                UserForm1.TextBox47.value = Format(sheetNetSalary, "0")
            Else
                ' ≈–« ﬂ«‰ «·’«›Ì √ﬂ»—° ‰ Õﬁﬁ „‰ «·Õœ «·√œ‰Ï 5000
                If sheetNetSalary >= 5000 Then
                    UserForm1.TextBox47.value = Format(calculatedAbsence, "0")
                Else
                    UserForm1.TextBox47.value = Format(sheetNetSalary, "0")
                End If
            End If
        Else
            ' √„«‰: ≈–« ·„ Ì ÿ«»ﬁ «·«”„ Ê«·—ﬁ„ ›Ì «·‘Ì 
            UserForm1.TextBox47.value = Format(netSalaryBox, "0")
        End If
        
    Else
        ' «·„”«— «·ÿ»Ì⁄Ì ﬂ⁄œœ ’ÕÌÕ »œÊ‰ ﬂ”Ê—
        UserForm1.TextBox47.value = Format(calculatedAbsence, "0")
    End If
End Sub


