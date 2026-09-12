Attribute VB_Name = "zero_athar_rag3e"
Sub ProcessMonthlyAdjustments()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.rowS.count, "B").End(xlUp).row
    
    For i = 9 To lastRow
        ' «б Ќёё: еб нжћѕ  «—нќ «” Ќё«ёњ жеб г÷м Џбне √яЋ— гд ‘е—њ
        If IsDate(ws.Cells(i, "CJ").value) Then
            If DateDiff("m", CDate(ws.Cells(i, "CJ").value), Date) >= 1 Then
                
                ' 1. «бЎ—Ќ ж«б ’Ён—: ед« дёжг »г”Ќ «б»н«д«  «б н «” хебя 
                ' е–« «б≈ћ—«Ѕ нгдЏ «бгЏ«ѕб«  гд «Ќ ”«» «б√Ћ— «б—ћЏн гћѕѕ«р
                ws.Cells(i, "BV").value = 0
                ws.Cells(i, "CI").value = 0
                ws.Cells(i, "CJ").value = ""
                
                ' гб«Ќў…: «б—« » «б«”гн (J) н»ём яг« еж б√де √’»Ќ «б—« » «бћѕнѕ ббгжўЁ.
            End If
        End If
    Next i
End Sub
