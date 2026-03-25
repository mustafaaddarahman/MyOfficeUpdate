VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
On Error GoTo emad
    Application.Calculation = xlCalculationAutomatic
 Call MonthlyBackupManager_DriveD
Dim FirstRun As String
    Dim lockTime As Variant
    Dim diff As Long

    ' 1. ›Õ’ Â· «·‰Ÿ«„ „€·ﬁ »”»» „Õ«Ê·… Œ«ÿ∆… (ﬁ›· 5 œﬁ«∆ﬁ)
    lockTime = GetSetting("MySystem", "Security", "LockTime", "")
    If lockTime <> "" Then
        diff = DateDiff("n", CDate(lockTime), Now)
        If diff < 5 Then
            MsgBox "«·‰Ÿ«„ „€·ﬁ „ƒﬁ « »”»» ≈œŒ«· ﬂÊœ Œ«ÿ∆. Ì—ÃÏ «·«‰ Ÿ«— " & (5 - diff) & " œﬁ«∆ﬁ.", vbCritical
            ThisWorkbook.Close SaveChanges:=False
            Exit Sub
        End If
    End If

    ' 2. ›Õ’ Â· ÂÌ «·„—… «·√Ê·Ï √Ê «‰ Â  «·”‰…
    FirstRun = GetSetting("MySystem", "Security", "ActivationDate", "")
    
    If FirstRun = "" Then
        ' √Ê· „—… Ì› Õ ›ÌÂ« «·‰Ÿ«„
        UserForm50.Show
    Else
        ' ›Õ’ Â· «‰ Â  «·”‰… (365 ÌÊ„)
        If Date >= DateAdd("yyyy", 1, CDate(FirstRun)) Then
            MsgBox "·ﬁœ «‰ Â  ’·«ÕÌ… «·‰Ÿ«„. Ì—ÃÏ ≈⁄«œ… «· ›⁄Ì· ··«” „—«—.", vbExclamation
            UserForm50.Show
        Else
            ' «· ›⁄Ì· ”«—Ì° «·œŒÊ· ··ÌÊ“—›Ê—„ 20
            UserForm20.Show
        End If
    End If

'============================================

Application.Visible = False
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"


Call open_workbok_1
 '================================================== Â–« «·ﬂÊœ ÌﬁÊ„ »Ã·» ﬂ«›… «·„ﬂ »«  Ê«· «ﬂœ „‰Â«
  Dim SystemPath As String
    Dim LibPath As String
    
    '  ÕœÌœ „”«— „Ã·œ «·‰Ÿ«„ »‰«¡ ⁄·Ï ‰Ê⁄ «·ÊÌ‰œÊ“ (64 √Ê 32 » )
    If Dir("C:\Windows\SysWOW64", vbDirectory) <> "" Then
        SystemPath = "C:\Windows\SysWOW64\"
    Else
        SystemPath = "C:\Windows\System32\"
    End If
    
    LibPath = SystemPath & "MSCOMCTL.OCX"
    
    '  ”ÃÌ· «·„ﬂ »… ’«„ « (ÌﬁÊ„ »«· ›⁄Ì·  ·ﬁ«∆Ì«)
    On Error Resume Next
    Shell "regsvr32 /s """ & LibPath & """", vbHide
    On Error GoTo 0
emad:
End Sub

'=====================================  ÕœÌÀ ”—⁄… „“«„‰… «·‘Ì « 
Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Â–« «·”ÿ— ÌÃ»— ≈ﬂ”· ⁄·Ï ≈⁄«œ… —”„ «·‘«‘… ›Ê— «·«‰ ﬁ«· ··‘Ì 
    Application.ScreenUpdating = True
    
    ' Â–« «·”ÿ— ÌﬁÊ„ »⁄„· “Ê„ ”—Ì⁄ Ê⁄Êœ Â · ’ÕÌÕ ≈“«Õ… «·√⁄„œ…
    ActiveWindow.Zoom = ActiveWindow.Zoom + 1
    ActiveWindow.Zoom = ActiveWindow.Zoom - 1
End Sub
