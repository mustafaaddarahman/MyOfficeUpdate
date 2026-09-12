VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm81 
   Caption         =   "UserForm81"
   ClientHeight    =   11292
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8040
   OleObjectBlob   =   "UserForm81.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public ColDict As Object


Private Sub ComboBox3_Change()
TextBox5.value = GetFormulaByTarget(ComboBox3.value)
    
End Sub

Private Sub ComboBox4_Change()

End Sub

Private Sub CommandButton3_Click()
Dim ColLetter As String
    ' ÇİÊÑÖäÇ Ãäß ÇÓÊÎÑÌÊ ÍÑİ ÇáÚãæÏ ÇáÌÏíÏ ãä ÇáßæãÈæÈæßÓ Ãæ ÇáÊíßÓÊ ÈæßÓ
    ColLetter = "CK" ' (ãËÇá)
    
    ' ÌáÈ ÇáãÚÇÏáÉ ÇáÍÇáíÉ
    Dim currentForm As String
    currentForm = GetFormulaByTarget(ComboBox3.value)
    
    ' ÚÑÖ ÇáãÚÇÏáÉ "ÈÚÏ ÇáÍÔÑ" İí ÇááíÈá ááÇÎÊÈÇÑ
    TextBox5.Text = Mid(SimulateNewFormula(currentForm, ColLetter), 2)
    
    MsgBox "ÊãÊ ÇáãÍÇßÇÉ! åá ÊÈÏæ ÇáãÚÇÏáÉ ÕÍíÍÉ İí TextBox5¿", vbQuestion
End Sub

Private Sub CommandButton4_Click()

Dim VBProj As Object
    Dim VBComp As Object
    Dim CodeMod As Object
    Dim NewFormula As String
    Dim oldFormula As String
    Dim i As Long
    Dim rawText As String
    
    ' 1. ÇáÊÍŞŞ ãä æÌæÏ ãÚÇÏáÉ
    If TextBox5.Text = "" Then
        MsgBox "áÇ ÊæÌÏ ãÚÇÏáÉ İí TextBox5!", vbCritical
        Exit Sub
    End If
    
    ' --- ÈÏÇíÉ ÇáÊÕÍíÍ ÇáĞßí (ÇáÔÑØ áãäÚ ÊßÑÇÑ ÇáíÓÇæí) ---
    rawText = TextBox5.Text
    
    ' ÊäÙíİ Ãí ÚáÇãÇÊ íÓÇæí ŞÏ Êßæä ãæÌæÏÉ ãÓÈŞÇğ İí ÇáäÕ
    Do While Left(rawText, 1) = "="
        rawText = Mid(rawText, 2)
    Loop
    
    ' ÅÖÇİÉ ÚáÇãÉ íÓÇæí æÇÍÏÉ İŞØ ááäÕ ÇáäÙíİ
    NewFormula = "=" & rawText
    ' --- äåÇíÉ ÇáÊÕÍíÍ ÇáĞßí ---
    
    oldFormula = GetFormulaByTarget(ComboBox3.value)
    
    ' 2. ÇáæÕæá ááãæÏíæá
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents("count_all_salalry")
    Set CodeMod = VBComp.CodeModule
    
    ' 3. ÇáÈÍË æÇáÇÓÊÈÏÇá ÏÇÎá ÇáãæÏíæá
    Dim found As Boolean
    found = False
    
    For i = 1 To CodeMod.CountOfLines
        ' ÇáÈÍË Úä ÇáãÚÇÏáÉ ÇáŞÏíãÉ ÏÇÎá Ãí ÓØÑ İí ÇáãæÏíæá
        If InStr(CodeMod.Lines(i, 1), oldFormula) > 0 Then
            ' ÇÓÊÈÏÇá ÇáãÚÇÏáÉ (ÇáãÊÛíÑ newFormula ÇáÂä íÍÊæí Úáì íÓÇæí æÇÍÏÉ İŞØ)
            CodeMod.ReplaceLine i, Replace(CodeMod.Lines(i, 1), oldFormula, NewFormula)
            found = True
            Exit For
        End If
    Next i
    
    ' 4. ÇáäÊíÌÉ
    If found Then
        MsgBox "Êã ÊÍÏíË ÇáãÚÇÏáÉ İí ÇáãæÏíæá ÈäÌÇÍ!", vbInformation
        ' ÊÍÏíË ÇáÍÓÇÈÇÊ İæÑÇğ
        Call count_all
    Else
        MsgBox "áã íÊã ÇáÚËæÑ Úáì ÇáãÚÇÏáÉ ÇáŞÏíãÉ İí ÇáãæÏíæá!", vbExclamation
    End If


End Sub

Private Sub CommandButton5_Click()
Unload Me
UserForm82.Show

End Sub

Private Sub TextBox5_Change()
If Len(Trim(TextBox5.value)) = 0 Then
        ComboBox4.Clear
        Exit Sub
    End If
    
    ' ÅÖÇİÉ ÊÃÎíÑ ÈÓíØ ÌÏÇğ (ÇÎÊíÇÑí) ÅĞÇ ßÇä ÇáäÕ ØæíáÇğ ÌÏÇğ
    ' áÊÌäÈ ÊäİíĞ ÇáßæÏ ÃËäÇÁ ÇáßÊÇÈÉ ÇáãÓÊãÑÉ
    
    Call AnalyzeFormulaFromTextBox5
End Sub

Private Sub AnalyzeFormulaFromTextBox5()
    Dim ws As Worksheet
    Dim formulaStr As String
    Dim formulaParts As Variant
    Dim i As Long, j As Long
    Dim ColLetter As String
    Dim HeaderName As String
    Dim part As String
    
    ' ÖÈØ æÑŞÉ ÇáÚãá (ÊÃßÏ ãä ãØÇÈŞÉ ÇáÇÓã "Sheet1")
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    formulaStr = TextBox5.value
    
    ' 1. ÇáÊäÙíİ ãä ÇáÑãæÒ ÇáÑíÇÖíÉ
    formulaStr = Replace(formulaStr, "=INT(", "")
    formulaStr = Replace(formulaStr, "(", "")
    formulaStr = Replace(formulaStr, ")", "")
    formulaStr = Replace(formulaStr, "+", "|")
    formulaStr = Replace(formulaStr, "-", "|")
    formulaStr = Replace(formulaStr, " ", "")
    
    formulaParts = Split(formulaStr, "|")
    
    ComboBox4.Clear
    
    ' 2. ÍáŞÉ ÇáãÚÇáÌÉ ãÚ ÇáÍãÇíÉ ãä ÇáÃÎØÇÁ
    For i = LBound(formulaParts) To UBound(formulaParts)
        part = Trim(formulaParts(i))
        
        ' ÇÓÊÎÑÇÌ ÇáÍÑæİ İŞØ
        ColLetter = ""
        For j = 1 To Len(part)
            If IsNumeric(Mid(part, j, 1)) Then Exit For
            ColLetter = ColLetter & Mid(part, j, 1)
        Next j
        
        ' 3. ÇáÊÍŞŞ ãä Ãä colLetter íãËá ÚãæÏÇğ ÕÍíÍÇğ
        ' (Øæá ÇáÚãæÏ ÚÇÏÉ áÇ íÒíÏ Úä 3 ÃÍÑİ ãËá AAA)
        If ColLetter <> "" And Len(ColLetter) <= 3 Then
            
            ' ÇÓÊÎÏÇã ÇáÎØÃ ÇáãßÊæã áÊÌÇæÒ Ãí ÚãæÏ ÛíÑ ãæÌæÏ Ãæ ÎØÃ İí ÇáäØÇŞ
            On Error Resume Next
            HeaderName = ws.Range(ColLetter & "8").value
            
            If Err.Number = 0 And HeaderName <> "" Then
                ' ãäÚ ÇáÊßÑÇÑ İí ÇáßæãÈæ ÈæßÓ
                Dim exists As Boolean: exists = False
                Dim k As Long
                For k = 0 To ComboBox4.ListCount - 1
                    If ComboBox4.List(k) = HeaderName Then exists = True: Exit For
                Next k
                
                If Not exists Then ComboBox4.AddItem HeaderName
            End If
            On Error GoTo 0
            
        End If
    Next i
End Sub



Private Sub UserForm_Activate()
End Sub

' --- 1. ÊåíÆÉ ÇáİæÑã æÇáŞÇãæÓ ---
Private Sub UserForm_Initialize()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Set ColDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    
    TextBox2.Visible = False: TextBox3.Visible = False: TextBox4.Visible = False
    
    ' ãáÁ ÇáŞÇãæÓ (İÍÕ ÇáÕİ 8 ãä ÇáÚãæÏ 1 Åáì 89 - Ãí ÍÊì CJ)
    For i = 1 To 89
        If Len(Trim(ws.Cells(8, i).value)) > 0 Then
            ColDict(ws.Cells(8, i).value) = Split(ws.Cells(1, i).Address, "$")(1)
            ComboBox1.AddItem ws.Cells(8, i).value
        End If
    Next i
    
    ComboBox2.List = Array("ÌãÚ ÚÇÏí", "ØÑÍ", "ÔÑØ: (IF)", "ÊÌãíÚ ÔÑØí: (SUMIF)", "äÓÈÉ: (%)", "ÍÏ ÃŞÕì: (MIN)")
    ComboBox3.List = Array("ãÌãæÚ ÇáãÎÕÕÇÊ", "ãÌãæÚ ÇáÇÓÊÍŞÇŞÇÊ", "ãÌãæÚ ÇáÇÓÊŞØÇÚÇÊ", "ÇáÑÇÊÈ Çáßáí")
        
    Label6.Caption = "ÇáÅÑÔÇÏ: ÇÎÊÑ ÏÇáÉ ááÈÏÁ..."





End Sub

' --- 2. ÇáÊÍßã İí ÅÙåÇÑ ÇáÍŞæá ---
Private Sub ComboBox2_Change()
    TextBox2.Visible = True: TextBox3.Visible = False: TextBox4.Visible = False
    Select Case ComboBox2.value
        Case "ÌãÚ ÚÇÏí", "ØÑÍ": TextBox2.Visible = False: Label6.Caption = "ÇáÅÑÔÇÏ: ÇÖÛØ ÅÖÇİÉ áÏãÌ ÇáÚãæÏ."
        Case "ÔÑØ: (IF)": TextBox3.Visible = True: TextBox4.Visible = True: Label6.Caption = "IF: ÇáÔÑØ | ŞíãÉ ÊÍŞŞ | ŞíãÉ ÚÏã ÊÍŞŞ"
        Case "ÊÌãíÚ ÔÑØí: (SUMIF)": TextBox3.Visible = True: Label6.Caption = "SUMIF: ÇáÔÑØ (ãËá >1000) | ÚãæÏ ÇáÌãÚ"
        Case "äÓÈÉ: (%)": Label6.Caption = "ÇáäÓÈÉ: ÃÏÎá ÇáäÓÈÉ (ãËáÇğ 10 ÊÚäí 10%)."
        Case "ÍÏ ÃŞÕì: (MIN)": Label6.Caption = "MIN: ÃÏÎá ÇáŞíãÉ ÇáŞÕæì."
    End Select
End Sub

' --- 3. ÅÖÇİÉ ÇáÎØæÇÊ ááÜ ListBox ---
Private Sub CommandButton1_Click()
Dim stepStr As String, colL As String
    If ComboBox1.value = "" Then Exit Sub
    colL = ColDict(ComboBox1.value)
    
    Select Case ComboBox2.value
        Case "ÌãÚ ÚÇÏí": stepStr = "=+" & colL & "9"
        Case "ØÑÍ": stepStr = "=-" & colL & "9"
        Case "äÓÈÉ: (%)": stepStr = "=*" & colL & "9*" & (val(TextBox2.value) / 100)
        Case "ÔÑØ: (IF)": stepStr = "=IF(" & colL & "9" & TextBox2.value & ";" & TextBox3.value & ";" & TextBox4.value & ")"
        Case "ÊÌãíÚ ÔÑØí: (SUMIF)": stepStr = "=SUMIF(" & colL & "9:" & colL & "500;" & """" & TextBox2.value & """;" & ColDict(TextBox3.value) & "9:" & ColDict(TextBox3.value) & "500)"
        Case "ÍÏ ÃŞÕì: (MIN)": stepStr = "=MIN(" & colL & "9;" & TextBox2.value & ")"
    End Select
    
    ListBox1.AddItem stepStr
    ' ... ÈÇŞí ßæÏ ÇáÊÍÏíË ...
End Sub

' --- 4. ÒÑ ÇáÊäİíĞ (ÇáÈÏÁ ÏÇÆãÇğ ãä ÇáÚãæÏ 90 / CK) ---


' --- 5. ÏÇáÉ ÇáÊäÙíİ ---




' ÇáÌÒÁ ÇáËÇäí: ÇáÊæÌíå (ÅÖÇİÉ ÇáãÑÌÚ ááãÚÇÏáÉ ÇáÃã)
Sub UpdateRouterFormula(ws As Worksheet, TargetCol As Long, lastR As Long, routerCellName As String)
    Dim TargetCell As Range
    Dim ColLetter As String
    
    ' ÊÍÏíÏ ÇáÎáíÉ ÇáãæÌåÉ (AD, AE, etc)
    Set TargetCell = ws.Range(routerCellName)
    
    ' ÇÓÊÎÑÇÌ ÇáÍÑİ ÇáÎÇÕ ÈÇáÚãæÏ ÇáÌÏíÏ (ãËáÇğ CK)
    ColLetter = Split(ws.Cells(1, TargetCol).Address, "$")(1)
    
    ' ÊÍÏíË ÇáãÚÇÏáÉ ÈÇáãÑÌÚ (ÈÏæä ÊŞØíÚ äÕí ãÚŞÏ)
    ' äÚÊãÏ Úáì ÇÓÊÈÏÇá ÇáŞæÓ ÇáäåÇÆí (ÊÖãä ÇÓÊŞÑÇÑ ÇáãÚÇÏáÉ)
    Dim currentFormula As String
    currentFormula = TargetCell.Formula
    
    If InStr(currentFormula, ")") > 0 Then
        TargetCell.Formula = Left(currentFormula, Len(currentFormula) - 1) & "+" & ColLetter & "9)"
    Else
        TargetCell.Formula = "=INT(" & ColLetter & "9)"
    End If
    
    ' ÊÚãíã ÇáÊæÌíå
    TargetCell.AutoFill Destination:=ws.Range(TargetCell, ws.Cells(lastR, TargetCell.Column))
End Sub

Sub ProcessNewColumn(ws As Worksheet, TargetCol As Long, lastR As Long)
    ' 1. ßÊÇÈÉ ÇÓã ÇáÚãæÏ (ÇáåíÏÑ) İí ÇáÕİ 8
    ws.Cells(8, TargetCol).value = UserForm1.TextBox1.value
    
    ' 2. ÊäİíĞ ÇáãÚÇÏáÉ İí ÇáÕİ 9 ÈäÇÁğ Úáì ÇÎÊíÇÑÇÊ ÇáãÓÊÎÏã
    ' äÓÊÎÏã ÏÇáÉ JoinList ÇáÊí ŞãÊ ÈÈÑãÌÊåÇ ãÓÈŞÇğ
    ws.Cells(9, TargetCol).Formula = "=INT(" & JoinList(UserForm1.ListBox1) & ")"
    
    ' 3. ÊÚãíã ÇáãÚÇÏáÉ áÃÓİá ÍÊì ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ (FillDown)
    ws.Range(ws.Cells(9, TargetCol), ws.Cells(lastR, TargetCol)).FillDown
End Sub


Private Sub CommandButton2_Click()

Dim ws As Worksheet
    Dim TargetCol As Long
    Dim lastR As Long
    
    ' 1. ÊÍÏíÏ æÑŞÉ ÇáÚãá
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 2. ÇáÈÍË Úä Ãæá ÚãæÏ İÇÑÛ ÈÏÁÇğ ãä ÇáÚãæÏ 89
    TargetCol = 89
    Do While Len(Trim(ws.Cells(8, TargetCol).value)) > 0
        TargetCol = TargetCol + 1
    Loop
    
    ' 3. ßÊÇÈÉ ÇÓã ÇáÚãæÏ ÇáÌÏíÏ İí ÇáÕİ ÇáËÇãä (Header)
    ws.Cells(8, TargetCol).value = TextBox1.value
    
    ' 4. æÖÚ ÇáãÚÇÏáÉ ÇáÃÕáíÉ (ÇáãæÌæÏÉ İí ÇááíÓÊ ÈæßÓ) İí ÇáÕİ ÇáÊÇÓÚ
    ' ÊÃßÏ Ãä ListBox1 íÍÊæí İÚáÇğ Úáì ÇáãÚÇÏáÉ İí ÇáÚäÕÑ ÇáÃæá
    ws.Cells(9, TargetCol).FormulaLocal = ListBox1.List(0)
    
    ' 5. ÊÍÏíÏ ÂÎÑ Õİ ÈäÇÁğ Úáì ÇáÚãæÏ ÇáËÇäí (B)
    lastR = ws.Cells(ws.rowS.count, 2).End(xlUp).row
    
    ' 6. ÊÚãíã ÇáãÚÇÏáÉ ááÚãæÏ ÈÇáßÇãá
    If lastR >= 9 Then
        ws.Range(ws.Cells(9, TargetCol), ws.Cells(lastR, TargetCol)).FillDown
    End If
    
    ' 7. ÑÓÇáÉ ÊÃßíÏ ÈÓíØÉ
    MsgBox "Êã ÅäÔÇÁ ÇáÚãæÏ '" & TextBox1.value & "' æÊÚÈÆÉ ãÚÇÏáÇÊå ÈäÌÇÍ!", vbInformation
    
    ' 8. ÅÛáÇŞ ÇáäãæĞÌ
    




End Sub

' ÊÃßÏ ãä æÌæÏ ÏÇáÉ JoinList İí äİÓ ÇáãæÏíæá
Function JoinList(lb As ListBox) As String
 Dim i As Integer
    Dim res As String
    
    ' ÈãÇ Ãä ßá ÚäÕÑ İíå =¡ äÏãÌ İŞØ ÇáÚäÇÕÑ ßãÇ åí
    res = lb.List(0)
    JoinList = res
   End Function


Public Function GetFormulaByTarget(ByVal LabelName As String) As String
    Dim TargetCol As String
    
    ' 1. ÎÑíØÉ ÇáÑÈØ ÇáÌÏíÏÉ
    Select Case Trim(LabelName)
        Case "ãÌãæÚ ÇáãÎÕÕÇÊ":     TargetCol = "AD"
        Case "ãÌãæÚ ÇáÇÓÊÍŞÇŞÇÊ":  TargetCol = "AE"
        Case "ãÌãæÚ ÇáÇÓÊŞØÇÚÇÊ":  TargetCol = "AV"
        Case "ÇáÑÇÊÈ Çáßáí":       TargetCol = "AW"
        Case Else:                 GetFormulaByTarget = "ÛíÑ ãÚÑİ": Exit Function
    End Select
    
    ' 2. ÎÑíØÉ ÇáãÚÇÏáÇÊ (ãÑÈæØÉ ÈÇáÍÑæİ ÇáãÍÏÏÉ ÃÚáÇå)
    Select Case TargetCol
        Case "AD": GetFormulaByTarget = Sheets("Sheet1").Range("AD9").Formula
        Case "AE": GetFormulaByTarget = Sheets("Sheet1").Range("Ae9").Formula
        Case "AV": GetFormulaByTarget = Sheets("Sheet1").Range("Av9").Formula
        Case "AW": GetFormulaByTarget = Sheets("Sheet1").Range("Aw9").Formula
    End Select
End Function

Public Function GetColLetterFromName(ByVal Name As String) As String
    Select Case Trim(Name)
        Case "ãÌãæÚ ÇáãÎÕÕÇÊ":    GetColLetterFromName = "AD"
        Case "ãÌãæÚ ÇáÇÓÊÍŞÇŞÇÊ":  GetColLetterFromName = "AE"
        Case "ãÌãæÚ ÇáÇÓÊŞØÇÚÇÊ":  GetColLetterFromName = "AV"
        Case "ÇáÑÇÊÈ Çáßáí":       GetColLetterFromName = "AW"
        Case Else:                 GetColLetterFromName = ""
    End Select
End Function

Public Function SimulateNewFormula(ByVal OriginalFormula As String, ByVal NewCol As String) As String
    ' äÊÍŞŞ ÃæáÇğ ÅĞÇ ßÇä ÇáÚãæÏ ãæÌæÏÇğ áÊÌäÈ ÇáÊßÑÇÑ
    If InStr(1, OriginalFormula, NewCol & "9", vbTextCompare) > 0 Then
        SimulateNewFormula = OriginalFormula ' áÇ ÍÇÌÉ ááÊÛííÑ
    Else
        ' äÍÔÑ ÇáÚãæÏ ÇáÌÏíÏ İí ÇáãÚÇÏáÉ
        SimulateNewFormula = Left(OriginalFormula, Len(OriginalFormula) - 1) & "+" & NewCol & "9)"
    End If
End Function








' ÏÇáÉ ãÓÇÚÏÉ ÊÌáÈ ÇÓã ÇáÚãæÏ ãä ÇáÕİ 8 ÈäÇÁğ Úáì ÇáÍÑİ
Function FindHeaderByLetter(ByVal ColLetter As String, ws As Worksheet) As String
    Dim ColIndex As Long
    ColIndex = ws.Columns(ColLetter).Column
    FindHeaderByLetter = ws.Cells(8, ColIndex).value
End Function

Dim formulaStr As String
    Dim colParts As Variant
    Dim i As Long
    Dim ColLetter As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    UserForm1.ComboBox4.Clear
    
    ' 1. ÌáÈ ÇáãÚÇÏáÉ
    formulaStr = GetFormulaByTarget(LabelName)
    
    ' 2. İÍÕ åá ÌáÈ ãÚÇÏáÉ Ãã ÎØÃ¿
    If formulaStr = "ÛíÑ ãÚÑİ" Or InStr(formulaStr, "ÎØÃ") > 0 Then
        MsgBox "ÇáÏÇáÉ áã ÊÌÏ ãÚÇÏáÉ İí ÇáãæÏíæá áÜ: " & LabelName, vbCritical
        Exit Sub
    End If
    
    ' 3. ÊäÙíİ ÇáãÚÇÏáÉ (äİÓ ÇáãäØŞ ÇáÓÇÈŞ)
    formulaStr = Replace(Replace(formulaStr, "=INT(", ""), ")", "")
    formulaStr = Replace(formulaStr, " ", "")
    colParts = Split(formulaStr, "+")
    
    ' 4. ÊÚÈÆÉ ÇáßæãÈæ ÈæßÓ
    Dim foundCount As Integer: foundCount = 0
    For i = LBound(colParts) To UBound(colParts)
        ' ÇÓÊÎÑÇÌ ÇáÍÑİ (ÍĞİ ÂÎÑ ÍÑİ æåæ ÇáÑŞã 9)
        ColLetter = Left(colParts(i), Len(colParts(i)) - 1)
        
        ' ÔÑØ ÇáÜ CK
        If ColLetter >= "CK" Then
            Dim HeaderName As String
            HeaderName = ws.Range(ColLetter & "8").value
            
            If HeaderName <> "" Then
                UserForm1.ComboBox4.AddItem HeaderName
                foundCount = foundCount + 1
            End If
        End If
    Next i
    
    ' İÍÕ äåÇÆí
    If foundCount = 0 Then
        MsgBox "Êã ÌáÈ ÇáãÚÇÏáÉ æáßä áã íÊã ÇáÚËæÑ Úáì ÃÚãÏÉ ÊÈÏÃ ãä CK İí ÇáÕİ 8!", vbExclamation
    End If
End Sub
Public Sub FillComboBox4FromFormula(LabelName As String)
    Dim formulaStr As String
    Dim colParts As Variant
    Dim i As Long
    Dim ColLetter As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    UserForm1.ComboBox4.Clear
    
    ' 1. ÌáÈ ÇáãÚÇÏáÉ
    formulaStr = GetFormulaByTarget(LabelName)
    
    ' 2. İÍÕ åá ÌáÈ ãÚÇÏáÉ Ãã ÎØÃ¿
    If formulaStr = "ÛíÑ ãÚÑİ" Or InStr(formulaStr, "ÎØÃ") > 0 Then
        MsgBox "ÇáÏÇáÉ áã ÊÌÏ ãÚÇÏáÉ İí ÇáãæÏíæá áÜ: " & LabelName, vbCritical
        Exit Sub
    End If
    
    ' 3. ÊäÙíİ ÇáãÚÇÏáÉ (äİÓ ÇáãäØŞ ÇáÓÇÈŞ)
    formulaStr = Replace(Replace(formulaStr, "=INT(", ""), ")", "")
    formulaStr = Replace(formulaStr, " ", "")
    colParts = Split(formulaStr, "+")
    
    ' 4. ÊÚÈÆÉ ÇáßæãÈæ ÈæßÓ
    Dim foundCount As Integer: foundCount = 0
    For i = LBound(colParts) To UBound(colParts)
        ' ÇÓÊÎÑÇÌ ÇáÍÑİ (ÍĞİ ÂÎÑ ÍÑİ æåæ ÇáÑŞã 9)
        ColLetter = Left(colParts(i), Len(colParts(i)) - 1)
        
        ' ÔÑØ ÇáÜ CK
        If ColLetter >= "CK" Then
            Dim HeaderName As String
            HeaderName = ws.Range(ColLetter & "8").value
            
            If HeaderName <> "" Then
                UserForm1.ComboBox4.AddItem HeaderName
                foundCount = foundCount + 1
            End If
        End If
    Next i
    
    ' İÍÕ äåÇÆí
    If foundCount = 0 Then
        MsgBox "Êã ÌáÈ ÇáãÚÇÏáÉ æáßä áã íÊã ÇáÚËæÑ Úáì ÃÚãÏÉ ÊÈÏÃ ãä CK İí ÇáÕİ 8!", vbExclamation
    End If
End Sub


Private Sub ComboBox4_AfterUpdate()
Dim selectedHeader As String, ColLetter As String
    Dim NewFormula As String, oldFormula As String
    Dim ws As Worksheet
    Dim response As VbMsgBoxResult
    Dim FoundCell As Range, lastR As Long
    
    selectedHeader = ComboBox4.value
    If selectedHeader = "" Then Exit Sub
    
    ' 1. ÓÄÇá ÇáãÓÊÎÏã ááÊÃßíÏ
    response = MsgBox("åá ÊÑíÏ ÍĞİ ÇáÚãæÏ '" & selectedHeader & "' ãä ÃÑßÇä åĞå ÇáÏÇáÉ¿", _
                      vbYesNo + vbQuestion, "ÊÃßíÏ ÇáÍĞİ")
    
    If response = vbNo Then
        ComboBox4.value = ""
        Exit Sub
    End If
    
    ' 2. ÍİÙ ÇáäÕ ÇáÍÇáí ááÈÍË Úäå
    oldFormula = TextBox5.value
    
    ' 3. ãÚÇáÌÉ ÇáäÕ: ÌáÈ ÇáÍÑİ æÊäÙíİ ÇáãÚÇÏáÉ
    ColLetter = GetColLetterByHeader(selectedHeader)
    NewFormula = oldFormula
    
    ' ÅÒÇáÉ ÚáÇãÉ ÇáíÓÇæí ãÄŞÊÇğ áÖãÇä ÚÏã æÌæÏåÇ ÃËäÇÁ ÇáÊäÙíİ
    NewFormula = Replace(NewFormula, "=", "")
    
    ' ÇáÊäÙíİ (ÍĞİ ÇáÚãæÏ æÇáÅÔÇÑÇÊ ÇáÒÇÆÏÉ)
    NewFormula = Replace(NewFormula, "+" & ColLetter & "9", "")
    NewFormula = Replace(NewFormula, "-" & ColLetter & "9", "")
    NewFormula = Replace(NewFormula, ColLetter & "9", "")
    
    ' ãÚÇáÌÉ ÇáÅÔÇÑÇÊ ÇáãÊÈŞíÉ
    NewFormula = Replace(NewFormula, "+-", "-")
    NewFormula = Replace(NewFormula, "-+", "-")
    NewFormula = Replace(NewFormula, "++", "+")
    
    ' 4. ÊÍÏíË TextBox5 (íÙåÑ ÈÏæä íÓÇæí)
    TextBox5.value = NewFormula
    
    ' 5. ÊÍÏíË ÇáÔíÊ (äÖÚ ÇáŞíãÉ ÈÏæä íÓÇæí áÃä count_all ÓíÖíİåÇ)
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set FoundCell = ws.rowS(9).Find(What:=oldFormula, LookIn:=xlFormulas, LookAt:=xlPart)
    
    If Not FoundCell Is Nothing Then
        ' äÓÊÎÏã .Value áäÑÓá ÇáäÕ ÇáÎÇã (ÈÏæä íÓÇæí)
        FoundCell.value = NewFormula
        
        ' ÊÚãíã ÇáãÚÇÏáÉ ááÃÓİá (FillDown)
        lastR = ws.Cells(ws.rowS.count, 2).End(xlUp).row
        If lastR >= 9 Then
            ws.Range(FoundCell, ws.Cells(lastR, FoundCell.Column)).FillDown
        End If
        
        MsgBox "Êã ÍĞİ ÇáÚãæÏ æÊÍÏíË ÇáÏÇáÉ ÈäÌÇÍ!", vbInformation
    Else
        MsgBox "áã íÊã ÇáÚËæÑ Úáì ãßÇä ÇáãÚÇÏáÉ İí ÇáÕİ 9.", vbExclamation
    End If
    
    ' 6. ÊÍÏíË ÇáŞÇÆãÉ
    Call AnalyzeFormulaFromTextBox5
    Call count_all
End Sub

' ÏÇáÉ ãÓÇÚÏÉ ááÍÕæá Úáì ÍÑİ ÇáÚãæÏ (íÌÈ Ãä Êßæä ãæÌæÏÉ İí ÇáÜ UserForm)
Private Function GetColLetterByHeader(HeaderName As String) As String
    Dim ws As Worksheet, i As Long
    Set ws = ThisWorkbook.Sheets("Sheet1")
    For i = 1 To 200
        If ws.Cells(8, i).value = HeaderName Then
            GetColLetterByHeader = Split(ws.Cells(1, i).Address, "$")(1)
            Exit Function
        End If
    Next i
End Function


Public Sub UpdateFormulaInModule(ByVal NewFormula As String)

    ' åäÇ íÌÈ Ãä ÊÖÚ ÇáßæÏ ÇáĞí íŞæã ÈÊÍÏíË ÇáãÚÇÏáÉ İí ÇáãßÇä ÇáĞí ÊõÎÒä İíå
    ' ÅĞÇ ßÇäÊ ãÚÇÏáÇÊß ãÎÒäÉ İí ãÕİæİÉ Ãæ İí ÎáíÉ ãÚíäÉ İí ÇáÔíÊ
    
    ' ãËÇá: ÅĞÇ ßäÊ ÊÍİÙ ÇáãÚÇÏáÉ İí ÎáíÉ ãÎİíÉ İí Sheet1 ÇÓãåÇ A1
    ThisWorkbook.Sheets("Sheet1").Range("A1").value = NewFormula
    
    ' Ãæ ÅĞÇ ßäÊ ÊÍÏË ãÊÛíÑÇğ ÚÇãÇğ (Global Variable)
    ' GlobalFormula = newFormula
    
    MsgBox "Êã ÊÍÏíË ÇáãÚÇÏáÉ İí ÇáãæÏíæá ÈäÌÇÍ."
End Sub


Private Function GetActualFormulaFromSheet(ColIndex As Long) As String
    ' ÊÌáÈ ÇáãÚÇÏáÉ ãä ÇáÎáíÉ ãÈÇÔÑÉ İí ÇáÕİ 9
    GetActualFormulaFromSheet = ThisWorkbook.Sheets("Sheet1").Cells(9, ColIndex).FormulaLocal
End Function
