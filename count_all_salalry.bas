Attribute VB_Name = "count_all_salalry"
Sub count_all()
On Error Resume Next
Application.Calculation = xlCalculationAutomatic


Dim wsMain As Worksheet: Set wsMain = ThisWorkbook.Sheets(1)
    Dim lastR_Main As Long: lastR_Main = wsMain.Cells(wsMain.rowS.count, 2).End(xlUp).row
    If lastR_Main < 9 Then lastR_Main = 9

    ' --- 1. «· ‰ŸÌ› Ê«·Õ”«»«  «· „ÂÌœÌ… («·„ €Ì—«  „ÕœÊœ… «·‰ÿ«ﬁ Â‰«) ---
    Dim rngCleanup As Range
    Dim cellTarget As Range
    Set rngCleanup = wsMain.Range("AU9:AU" & lastR_Main)
    
    
    lastR_Main = wsMain.Cells(wsMain.rowS.count, 2).End(xlUp).row
'«· √ﬂœ „‰ √‰ ¬Œ— ’› ·Ì” √ﬁ· „‰ 9 · Ã‰» «·Œÿ√
  
    
    
    For Each cellTarget In rngCleanup
        If IsEmpty(cellTarget) Then cellTarget.value = 0
    Next cellTarget


    wsMain.Range("A9").Formula = "=IF(B9<>"""", ROW()-8, """")"
    wsMain.Range("A9:A" & lastR_Main).FillDown


    wsMain.Cells(9, "N").Formula = "=INT(J9*M9)"
    wsMain.Range("N9:N" & lastR_Main).FillDown

    wsMain.Cells(9, "P").Formula = "=INT(J9*O9)"
    wsMain.Range("P9:P" & lastR_Main).FillDown

    wsMain.Cells(9, "R").Formula = "=INT(J9*Q9)"
    wsMain.Range("R9:R" & lastR_Main).FillDown

    wsMain.Cells(9, "T").Formula = "=INT(IF(BF9=""Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« "", V9*0, IF(V9>4, 4*10000, V9*10000)))"
    wsMain.Range("T9:T" & lastR_Main).FillDown

    ' --- 2. «·Ã“¡ «·Œ«’ »›Õ’ «·⁄„Êœ CG (√”„«¡ „ €Ì—«  ›—Ìœ…) ---
    Dim arrCG_Data As Variant
    Dim idx_CG As Long
    Dim countTrue_CG As Long, countFalse_CG As Long
    
    arrCG_Data = wsMain.Range("CG9:CG" & lastR_Main).value
    For idx_CG = 1 To UBound(arrCG_Data, 1)
        If LCase(CStr(arrCG_Data(idx_CG, 1))) = "true" Then
            countTrue_CG = countTrue_CG + 1
        Else
            countFalse_CG = countFalse_CG + 1
        End If
    Next idx_CG

    ' «· Õﬁﬁ „‰ «·Õ«·…
    If countFalse_CG = 0 Then
        ' «·Õ«·… 1: ﬂ· «·ﬁÌ„ True
        Dim arrCH_Data As Variant, arrOutput_Data As Variant
        arrCH_Data = wsMain.Range("CH9:CH" & lastR_Main).value
        ReDim arrOutput_Data(1 To UBound(arrCH_Data, 1), 1 To 2)
        For idx_CG = 1 To UBound(arrCH_Data, 1)
            arrOutput_Data(idx_CG, 1) = arrCH_Data(idx_CG, 1)
            arrOutput_Data(idx_CG, 2) = "true"
        Next idx_CG
        wsMain.Range("X9:X" & lastR_Main).value = Application.Index(arrOutput_Data, 0, 1)
        wsMain.Range("CG9:CG" & lastR_Main).value = Application.Index(arrOutput_Data, 0, 2)

    ElseIf countTrue_CG > 0 And countFalse_CG > 0 Then
        ' «·Õ«·… 2: Œ·Ìÿ - «·ﬁ›“
        GoTo SkipTaxLogic

    ElseIf countTrue_CG = 0 Then
        ' «·Õ«·… 3: ﬂ· «·ﬁÌ„ False
        wsMain.Range("X9").Formula2R1C1 = "=IFERROR(IF(OR(RC[-1]=thariaba1!R1C20:R2C20),VLOOKUP(RC[-14],astktat1,RC[-2]+3,1)," & _
                  "IF(OR(RC[-1]=thariaba2!R1C19),VLOOKUP(RC[-14],astktaat2,RC[-2]+3,1)," & _
                  "IF(OR(RC[-1]=thariaba3!R1C19:R3C19),VLOOKUP(RC[-14],astktaat3,RC[-2]+3,1)," & _
                  "IF(VLOOKUP(RC[-1],sheet1!R1C77:R5C78,2,0)=5," & _
                  "INDEX(Sheet2!R9C4:R8000C4, MATCH(1, (Sheet2!R9C2:R8000C2<=RC[-14])*(Sheet2!R9C3:R8000C3>=RC[-14]), 0))," & _
                  "IF(AND(VLOOKUP(RC[-1],sheet1!R1C77:R5C78,2,0)<4,RC[-2]>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(RC[-14],astktaat,VLOOKUP(RC[-1],sheet1!R1C77:R5C78,2,0)+RC[-2],1)))))),0)"
        wsMain.Range("X9:X" & lastR_Main).FillDown
        wsMain.Range("CG9").value = "false"
        wsMain.Range("CG9:CG" & lastR_Main).FillDown
    End If

SkipTaxLogic:
    ' --- 3. «·Õ”«»«  «·Œ «„Ì… ---
    wsMain.Cells(9, "AA").Formula = "=INT(J9*Z9)"
    wsMain.Range("AA9:AA" & lastR_Main).FillDown

    wsMain.Cells(9, "AC").Formula = "=INT(J9*AB9)"
    wsMain.Range("AC9:AC" & lastR_Main).FillDown

    wsMain.Range("AD9").Formula = "=INT(N9+P9+R9+S9+T9+U9+Y9+AA9+AC9+BH9+BP9+BR9)"
    wsMain.Range("AD9:AD" & lastR_Main).FillDown

    wsMain.Range("AE9").Formula = "=INT(AD9+J9+CE9)"
    wsMain.Range("AE9:AE" & lastR_Main).FillDown

    wsMain.Range("AT9").Formula = "=IF(BF9=""⁄ﬁœ"", 0, INT(J9*0.0025))"
    wsMain.Range("AT9:AT" & lastR_Main).FillDown

    wsMain.Cells(9, "AF").Formula = "=INT(IF(BF9=""⁄ﬁœ"", J9*$DH$8, J9*0.1))"
    wsMain.Range("AF9:AF" & lastR_Main).FillDown

    wsMain.Cells(9, "AP").Formula = "=INT(IF(H9=0,0,IF(H9<=6,1500,1000)))"
    wsMain.Range("AP9:AP" & lastR_Main).FillDown

    wsMain.Range("AV9").Formula = "=INT(X9+AF9+AG9+AH9+AI9+AJ9+AL9+AN9+AP9+AR9+AT9+AU9+BJ9+BO9+BQ9+AX9+CK9)"
    wsMain.Range("AV9:AV" & lastR_Main).FillDown

    wsMain.Range("AW9").Formula = "=INT(AE9-AV9)"
    wsMain.Range("AW9:AW" & lastR_Main).FillDown

    DoEvents


End Sub




' €·«¡ „⁄Ì‘… Formula: =IF(J9>500000,30000,20000)
