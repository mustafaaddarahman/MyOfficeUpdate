Attribute VB_Name = "export_all_data54"

Sub export_all_data_2()
On Error Resume Next
Sheets(54).Range("bw8:ci10000").EntireColumn.Hidden = True
Dim PASS As Variant
Dim lastR, i As Integer
lastR = Sheets(45).Cells(rowS.count, 2).End(xlUp).row
For i = 1 To lastR
PASS = Sheets(45).Cells(i, 2).value

Next i


    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(54).Range("a2:ci7").ClearContents
    Sheets(54).Range("a2:ci7").Merge
    Range("a2:ci7").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(54).Range("a2:ci7").Font.Name = "Arial"
   Sheets(54).Range("a2").Font.Bold = True
       
     Sheets(54).Range("a2:ci7").ClearContents
    Sheets(54).Range("a2:ci7").HorizontalAlignment = xlCenter
     Sheets(54).Range("a2:ci7").VerticalAlignment = xlCenter
       
    Sheets(54).Range("a2:ci7").value = UserForm36.ComboBox1.value & " " & "«·»Ì«‰«  «·ﬂ·Ì… ·’—›Ì«  «·—« » " & "    " & "·‘Â—" & "    " & monthName(Month(Date)) & "    " & Sheets _
    (1).Cells(9, "be").value
    
    'Cells(2, "a").EntireColumn.AutoFit
    
    
   
   
    
    
    
    
    
    
    
    
    
    
    Set ws = Sheets(54)
   
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Application.Visible = True
            
            ' =================================================================
            ' «·ﬂÊœ «·„÷«›: »‰«¡  –ÌÌ· «· ﬁ—Ì— Ê Ê”ÌÿÂ »ÕÃ„ Œÿ 20 Bold
            ' =================================================================
            Dim wsTarget As Worksheet
            Dim lastRowData As Long
            Dim lastColData As Long
            Dim startRowFooter As Long
            Dim endRowFooter As Long
            Dim midColFooter As Long
            Dim footerText As String
            
            '  ⁄ÌÌ‰ «·‘Ì  «·√Ê· ›Ì «·„’‰› «·ÃœÌœ «·„› ÊÕ Õ«·Ì« »⁄œ «·‹ Copy
            Set wsTarget = ActiveWorkbook.Sheets(1)
            
            ' Õ”«» ¬Œ— ’› ÌÕ ÊÌ ⁄·Ï »Ì«‰«  «·„ÊŸ›Ì‰ (»œ¡« „‰ «·’› 8 · ŒÿÌ «· —ÊÌ”…)
            lastRowData = wsTarget.Cells(wsTarget.rowS.count, "B").End(xlUp).row
            If lastRowData < 8 Then lastRowData = wsTarget.UsedRange.rowS.count
            
            '  ÕœÌœ ¬Œ— ⁄„Êœ Ÿ«Â— ··»Ì«‰«  («·⁄„Êœ 74 ÊÂÊ BV ·√‰ BW „Œ›Ì) ·ÌﬂÊ‰ «· –ÌÌ· „ ‰«”ﬁ«
            lastColData = 74
            
            ' ≈⁄œ«œ«  ’›Ê› «· –ÌÌ·:  —ﬂ 4 ’›Ê› ›«—€… À„ ÕÃ“ 4 ’›Ê› „œ„Ã… ··›Ê —
            startRowFooter = lastRowData + 5
            endRowFooter = startRowFooter + 3
            
            '  ﬁ”Ì„ «·√⁄„œ… «·Ÿ«Â—… ≈·Ï ‰’›Ì‰ „ ”«ÊÌÌ‰  „«„«
            midColFooter = lastColData \ 2
            
            ' ⁄»«—… «· –ÌÌ· «·„ÿ·Ê»… »«· Ê«ﬁÌ⁄ «·À·«À…
            footerText = "„”ƒÊ· ‘⁄»… „«·Ì…            „”ƒÊ· ‘⁄»… —ﬁ«»…               „œÌ— «·„’‰⁄            —∆Ì” ﬁ”„ «·„«·Ì                       —∆Ì” ﬁ”„ «·—ﬁ«»… «·œ«Œ·Ì…                      „œÌ— «·⁄«„"
            
            ' --- «· –ÌÌ· «·√Ê· («·‰’› «·√Ì„‰: „‰ «·⁄„Êœ 1 ≈·Ï «·„‰ ’›) ---
            wsTarget.Range(wsTarget.Cells(startRowFooter, 1), wsTarget.Cells(endRowFooter, midColFooter)).Merge
            With wsTarget.Cells(startRowFooter, 1)
                .value = footerText
                .Font.Name = "Arial"
                .Font.Size = 20 '  ⁄œÌ· ÕÃ„ «·Œÿ ≈·Ï 20 Õ”» ÿ·»ﬂ
                .Font.Bold = True ' ≈»ﬁ«¡ «·Œÿ »Ê·œ
                .Alignment.Horizontal = xlCenter ' «· Ê”Ìÿ «·√›ﬁÌ
                .Alignment.Vertical = xlCenter ' «· Ê”Ìÿ «·⁄„ÊœÌ («·«— ›«⁄ ›Ì «·„‰ ’›)
                .Alignment.WrapText = True
            End With
            ' ≈÷«›… ≈ÿ«— Œ›Ì› ÕÊ· «· –Ì· «·√Ê·
            With wsTarget.Range(wsTarget.Cells(startRowFooter, 1), wsTarget.Cells(endRowFooter, midColFooter)).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(150, 150, 150)
            End With
            
            ' --- «· –ÌÌ· «·À«‰Ì («·‰’› «·√Ì”—: „‰ «·„‰ ’› + 1 ≈·Ï ‰Â«Ì… «·√⁄„œ… «·Ÿ«Â—…) ---
            wsTarget.Range(wsTarget.Cells(startRowFooter, midColFooter + 1), wsTarget.Cells(endRowFooter, lastColData)).Merge
            With wsTarget.Cells(startRowFooter, midColFooter + 1)
                .value = footerText
                .Font.Name = "Arial"
                .Font.Size = 20 '  ⁄œÌ· ÕÃ„ «·Œÿ ≈·Ï 20 Õ”» ÿ·»ﬂ
                .Font.Bold = True ' ≈»ﬁ«¡ «·Œÿ »Ê·œ
                .Alignment.Horizontal = xlCenter ' «· Ê”Ìÿ «·√›ﬁÌ
                .Alignment.Vertical = xlCenter ' «· Ê”Ìÿ «·⁄„ÊœÌ («·«— ›«⁄ ›Ì «·„‰ ’›)
                .Alignment.WrapText = True
            End With
            ' ≈÷«›… ≈ÿ«— Œ›Ì› ÕÊ· «· –Ì· «·À«‰Ì »‰›” «·‰”ﬁ Ê«· ‰”Ìﬁ  „«„«
            With wsTarget.Range(wsTarget.Cells(startRowFooter, midColFooter + 1), wsTarget.Cells(endRowFooter, lastColData)).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
                .Color = RGB(150, 150, 150)
            End With
            ' =================================================================
            
             Application.ActiveWorkbook.SaveAs fileName:="D:\export\" & " " & " ›Ê—„… «·»Ì«‰«  «·ﬂ·Ì…" & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
             Application.Workbooks(2).Close True
            
        End With
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False

End Sub















