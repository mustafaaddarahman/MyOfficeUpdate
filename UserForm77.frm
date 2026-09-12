VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm77 
   Caption         =   "UserForm77"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "UserForm77.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm77"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 Dim maxWidth As Double, totalEntries As Long, currentEntry As Long, pctDone As Single
    Dim Counter As Long, EndRowData As Long, TargetRow As Long
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim startTime As Double, secondsElapsed As Double, estRemaining As Double
    
    Set wsTarget = Sheets(15)
    Set wsSource = Sheets(30)

    With UserForm10.Label3
        .Width = 0
        .Caption = ""
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = vbWhite
        .TextAlign = 2
    End With
    
    maxWidth = UserForm10.Label2.Width
    EndRowData = wsSource.Cells(wsSource.rowS.count, 2).End(xlUp).row
    totalEntries = EndRowData - 8
    
    If totalEntries <= 0 Then Exit Sub

    ' --- ≈⁄œ«œ«  «·”—⁄… «·ﬁ’ÊÏ ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False ' ≈Ìﬁ«› «·√Õœ«À ·“Ì«œ… «·”—⁄…
    Application.DisplayStatusBar = False ' ≈Ìﬁ«› ‘—Ìÿ «·Õ«·… · Ê›Ì— «·„Ê«—œ
    Application.CutCopyMode = False
    
    Call UpdateSalaryTapesWithDynamicFormulas
    wsSource.Range("B8:BP30000").Sort Key1:=wsSource.Range("e9"), Order1:=xlAscending, header:=xlYes
    
    wsTarget.rowS("10:" & wsTarget.rowS.count).ClearContents
    wsTarget.rowS("10:" & wsTarget.rowS.count).ClearFormats

    startTime = Timer
    wsTarget.rowS("1:7").Copy

    TargetRow = 1
    For Counter = 9 To EndRowData
        currentEntry = currentEntry + 1
        
        '  ﬁ·Ì·  ÕœÌÀ «·Ê«ÃÂ… („—… ﬂ· 10 ﬁÌÊœ) ÂÊ «·”— «·ÕﬁÌﬁÌ ··”—⁄… ›Ì Ê÷⁄ «· Œ›Ì
        If currentEntry Mod 10 = 0 Or currentEntry = totalEntries Then
            pctDone = currentEntry / totalEntries
            secondsElapsed = Timer - startTime
            If currentEntry > 1 Then
                estRemaining = (secondsElapsed / currentEntry) * (totalEntries - currentEntry)
            End If
            
            UserForm10.Label3.Width = pctDone * maxWidth
            UserForm10.Label3.Caption = "Ã«—Ì «‰‘«¡ «‘—ÿ… «·—« ».. „ »ﬁÌ: " & Format(estRemaining / 86400, "nn:ss")
            DoEvents
        End If

        TargetRow = ((Counter - 9) * 9) + 1
        If Counter > 9 Then
            wsTarget.rowS(TargetRow & ":" & TargetRow + 6).PasteSpecial Paste:=xlPasteAll
        End If
        wsTarget.Cells(TargetRow, 2).value = wsSource.Range("E" & Counter).value
    Next Counter

    ' --- ≈⁄«œ… «·≈⁄œ«œ«  ---
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    
    UserForm10.Label3.Caption = "100%  „ «·«‰ Â«¡"
    MsgBox " „ ≈‰‘«¡ «·√‘—ÿ… »‰Ã«Õ ·„’‰⁄ «·ﬂ—«„….", vbInformation
    
    Unload Me
    
    
End Sub

Private Sub CommandButton2_Click()
 Dim maxWidth As Double, totalEntries As Long, currentEntry As Long, pctDone As Single
    Dim Counter As Long, EndRowData As Long, TargetRow As Long
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim startTime As Double, secondsElapsed As Double, estRemaining As Double
    
    Set wsTarget = Sheets(15)
    Set wsSource = Sheets(30)

    With UserForm10.Label3
        .Width = 0
        .Caption = ""
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = vbWhite
        .TextAlign = 2
    End With
    
    maxWidth = UserForm10.Label2.Width
    EndRowData = wsSource.Cells(wsSource.rowS.count, 2).End(xlUp).row
    totalEntries = EndRowData - 8
    
    If totalEntries <= 0 Then Exit Sub

    ' --- ≈⁄œ«œ«  «·”—⁄… «·ﬁ’ÊÏ ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False ' ≈Ìﬁ«› «·√Õœ«À ·“Ì«œ… «·”—⁄…
    Application.DisplayStatusBar = False ' ≈Ìﬁ«› ‘—Ìÿ «·Õ«·… · Ê›Ì— «·„Ê«—œ
    Application.CutCopyMode = False
    
    Call UpdateSalaryTapesWithDynamicFormulas
    wsSource.Range("B8:BP30000").Sort Key1:=wsSource.Range("G9"), Order1:=xlAscending, header:=xlYes
    
    wsTarget.rowS("10:" & wsTarget.rowS.count).ClearContents
    wsTarget.rowS("10:" & wsTarget.rowS.count).ClearFormats

    startTime = Timer
    wsTarget.rowS("1:7").Copy

    TargetRow = 1
    For Counter = 9 To EndRowData
        currentEntry = currentEntry + 1
        
        '  ﬁ·Ì·  ÕœÌÀ «·Ê«ÃÂ… („—… ﬂ· 10 ﬁÌÊœ) ÂÊ «·”— «·ÕﬁÌﬁÌ ··”—⁄… ›Ì Ê÷⁄ «· Œ›Ì
        If currentEntry Mod 10 = 0 Or currentEntry = totalEntries Then
            pctDone = currentEntry / totalEntries
            secondsElapsed = Timer - startTime
            If currentEntry > 1 Then
                estRemaining = (secondsElapsed / currentEntry) * (totalEntries - currentEntry)
            End If
            
            UserForm10.Label3.Width = pctDone * maxWidth
            UserForm10.Label3.Caption = "Ã«—Ì «‰‘«¡ «‘—ÿ… «·—« ».. „ »ﬁÌ: " & Format(estRemaining / 86400, "nn:ss")
            DoEvents
        End If

        TargetRow = ((Counter - 9) * 9) + 1
        If Counter > 9 Then
            wsTarget.rowS(TargetRow & ":" & TargetRow + 6).PasteSpecial Paste:=xlPasteAll
        End If
        wsTarget.Cells(TargetRow, 2).value = wsSource.Range("E" & Counter).value
    Next Counter

    ' --- ≈⁄«œ… «·≈⁄œ«œ«  ---
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayStatusBar = True
    Application.ScreenUpdating = True
    
    UserForm10.Label3.Caption = "100%  „ «·«‰ Â«¡"
    MsgBox " „ ≈‰‘«¡ «·√‘—ÿ… »‰Ã«Õ ·„’‰⁄ «·ﬂ—«„….", vbInformation
    
    Unload Me
   
End Sub
