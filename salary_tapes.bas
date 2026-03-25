Attribute VB_Name = "salary_tapes"
Sub UpdateSalaryTapesWithDynamicFormulas()
    Dim wsTapes As Worksheet
    Dim wsData As Worksheet
    
    '  ÕœÌœ «·‘Ì « 
    Set wsTapes = ThisWorkbook.Sheets("Salary_tapes")
    Set wsData = ThisWorkbook.Sheets("sheet1")
    
    ' ÷„«‰  ‰”Ìﬁ «·Œ·«Ì« ﬂ‹ "⁄«„" ·÷„«‰ ⁄„· «·œÊ«·
    wsTapes.Range("A1:L7").NumberFormat = "General"
    
    With wsTapes
        ' 1. «·Œ·Ì… G1  ⁄—÷ «”„ «·‘Â— «·Õ«·Ì Ê«·”‰…
        .Range("G1").Formula = "=TEXT(TODAY(),""[$-ar-IQ]mmmm yyyy"")"

        ' 2. —»ÿ B1 »«·Œ·Ì… E9 ›Ì ‘Ì  1 »‘ﬂ· À«» 
        .Range("B1").Formula = "=sheet1!E9"
        
        ' 3. “—⁄ «·œÊ«· »«·‰„ÿ «·„ÿ«ÿÌ («·⁄„Êœ «·ﬂ«„·)
        .Range("E1").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$G:$G,,0)"
        .Range("J1").Formula = "=title_factory!$A$2"
        .Range("K1").Formula = "=XLOOKUP(B1,sheet30!$E:$E,sheet30!$AZ:$AZ,,0)"
        
        ' «·”ÿ— «·À«‰Ì
        .Range("E2").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$B:$B,,0)"
        .Range("H2").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$F:$F,,0)"
        .Range("J2").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$h:$h,,0)"
        .Range("L2").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$I:$I,,0)"
        
        ' «·”ÿ— «·À«·À
        .Range("B3").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$J:$J,,0)"
        .Range("D3").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$R:$R,,0)"
        .Range("F3").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$P:$P,,0)"
        .Range("H3").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$N:$N,,0)"
        .Range("J3").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AC:$AC,,0)"
        .Range("L3").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$Z:$Z,,0)"
        
        ' «·”ÿ— «·—«»⁄
        .Range("B4").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AA:$AA,,0)"
        .Range("D4").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$T:$T,,0)"
        .Range("F4").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$U:$U,,0)"
        .Range("H4").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$S:$S,,0)"
        .Range("J4").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AD:$AD,,0)"
        .Range("L4").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$X:$X,,0)"
        
        ' «·”ÿ— «·Œ«„”
        .Range("B5").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AJ:$AJ,,0)"
        .Range("D5").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$K:$K,,0)"
        .Range("F5").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AU:$AU,,0)"
        .Range("H5").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AI:$AI,,0)"
        .Range("J5").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AN:$AN,,0)"
        .Range("L5").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AO:$AO,,0)"
        
        ' «·”ÿ— «·”«œ”
        .Range("B6").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AR:$AR,,0)"
        .Range("D6").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AT:$AT,,0)"
        .Range("F6").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AF:$AF,,0)+ XLOOKUP(B1,sheet1!$E:$E,sheet1!$Ag:$Ag,,0)"
        .Range("H6").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AP:$AP,,0)"
        .Range("J6").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AH:$AH,,0)"
        .Range("L6").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$BH:$BH,,0)"
        .Range("E6").Formula = "=IF(XLOOKUP(B1,sheet1!$E:$E,sheet1!$BF:$BF,"""")=""Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« "",""«· Êﬁ›«  «· ﬁ«⁄œÌ… 25 %"",""«· Êﬁ›«  «· ﬁ«⁄œÌ… 10 %"")"
        
        ' «·”ÿ— «·”«»⁄
        .Range("B7").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AL:$AL,,0)"
        .Range("D7").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$BJ:$BJ,,0)"
        .Range("F7").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AV:$AV,,0)"
        .Range("H7").Formula = "=XLOOKUP(B1,sheet1!$E:$E,sheet1!$AW:$AW,,0)"
    End With
    
    Application.Calculate
End Sub
