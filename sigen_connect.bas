Attribute VB_Name = "sigen_connect"
'https://script.google.com/macros/s/AKfycbz8aNEkZQ6TOHRkrQgXVBweTe21dT3C0FN4n2kOfnckdSR2QY1oVPm31CafK-1YSQUK/exec





Public Sigen_NextPulse As Date

Public Sub Sigen_Main_Engine()
    On Error Resume Next
    Dim frm9 As Object: Set frm9 = UserForm9
    
    ' --- ≈÷«›… «· Õﬁﬁ „‰ CheckBox3 ---
    ' ≈–« ﬂ«‰ «·‹ CheckBox3 €Ì— „›⁄· (False)° ‰Êﬁ› «·—«œ«— Ê‰Œ—Ã
    If frm9.CheckBox3.value = False Then
        '  ‰ŸÌ› «·⁄‰«’— «·„ÊÃÊœ… Õ«·Ì«
        Dim iCtrl As Integer
        For iCtrl = frm9.Controls.count - 1 To 0 Step -1
            If Left(frm9.Controls(iCtrl).Name, 7) = "SigInd_" Or Left(frm9.Controls(iCtrl).Name, 7) = "SigTxt_" Then
                frm9.Controls.Remove frm9.Controls(iCtrl).Name
            End If
        Next iCtrl
        
        ' ≈·€«¡ «·ÃœÊ·… «·ﬁ«œ„… ·„‰⁄ ≈⁄«œ…  ‘€Ì· «·—«œ«—
        On Error Resume Next
        Application.OnTime Sigen_NextPulse, "sigen_connect.Sigen_Trigger", , False
        Exit Sub
    End If
    ' ----------------------------------

    ' ≈–« ﬂ«‰ «· ‘€Ì· «·√Ê·
    Static FirstRun As Boolean
    If FirstRun = False Then
        FirstRun = True
        Sigen_NextPulse = Now + TimeValue("00:00:05")
        Application.OnTime Sigen_NextPulse, "sigen_connect.Sigen_Trigger"
        Exit Sub
    End If
    
    ' (»ﬁÌ… «·ﬂÊœ «·Œ«’ »ﬂ Ì»ﬁÏ ﬂ„« ÂÊ...)
    Dim Sigen_URL As String: Sigen_URL = "https://script.google.com/macros/s/AKfycbz8aNEkZQ6TOHRkrQgXVBweTe21dT3C0FN4n2kOfnckdSR2QY1oVPm31CafK-1YSQUK/exec"
    Dim Sigen_MyName As String: Sigen_MyName = Trim(ThisWorkbook.Sheets("title_factory").Range("A2").value)
    
    Dim Sigen_Http As Object: Set Sigen_Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Sigen_MyName <> "" Then
        Sigen_Http.Open "POST", Sigen_URL, False
        Sigen_Http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        Sigen_Http.Send "station=" & Application.WorksheetFunction.EncodeURL(Sigen_MyName) & "&status=Active&t=" & Timer
    End If
    
    Sigen_Http.Open "GET", Sigen_URL & "?action=GET_ACTIVE&t=" & Timer & "&rnd=" & Int((1000 * Rnd) + 1), False
    Sigen_Http.Send
    Dim Sigen_Response As String: Sigen_Response = Sigen_Http.responseText
    
    Dim lastRow As Long
    lastRow = Sheets(5).Cells(rowS.count, "B").End(xlUp).row
    Dim Sigen_List As Variant
    Sigen_List = Application.Transpose(Sheets(5).Range("B1:B" & lastRow).value)
    
    '  ‰ŸÌ› «·⁄‰«’— ﬁ»· ≈⁄«œ… «·—”„
    For iCtrl = frm9.Controls.count - 1 To 0 Step -1
        If Left(frm9.Controls(iCtrl).Name, 7) = "SigInd_" Or Left(frm9.Controls(iCtrl).Name, 7) = "SigTxt_" Then
            frm9.Controls.Remove frm9.Controls(iCtrl).Name
        End If
    Next iCtrl
    
    ' —”„ «·⁄‰«’—
    Dim j As Integer
    Dim Sigen_TopPos As Single: Sigen_TopPos = 142
    For j = 0 To UBound(Sigen_List)
        Dim SName As String: SName = Sigen_List(j)
        Dim SInd As Object: Set SInd = frm9.Controls.Add("Forms.Label.1", "SigInd_" & j, True)
        SInd.Left = 10: SInd.Width = 10: SInd.Height = 10: SInd.Top = Sigen_TopPos + (j * 18): SInd.BorderStyle = 1: SInd.Caption = ""
        Dim STxt As Object: Set STxt = frm9.Controls.Add("Forms.Label.1", "SigTxt_" & j, True)
        STxt.Left = 25: STxt.Width = 180: STxt.Height = 15: STxt.Top = SInd.Top - 1: STxt.Caption = SName: STxt.BackStyle = 0: STxt.Font.Size = 9
        
       If InStr(1, Sigen_Response, Chr(34) & SName & Chr(34), vbTextCompare) > 0 Then
            ' «·„Õÿ… „ ’·…
            SInd.BackColor = vbYellow ' √’›—
            STxt.Font.Bold = True
            STxt.Font.Size = 12
            STxt.ForeColor = RGB(255, 255, 255)
            STxt.Font.Name = "Times New Roman"
            ' ≈÷«›… ﬂ·„… „ ’· Õ«·Ì«
            STxt.Caption = SName & " („ ’· Õ«·Ì«)"
        Else
            ' «·„Õÿ… €Ì— „ ’·…
            SInd.BackColor = vbRed    ' √Õ„—
             STxt.Font.Size = 11
            STxt.Font.Bold = True
            STxt.ForeColor = RGB(255, 255, 255)
            ' «·≈»ﬁ«¡ ⁄·Ï «”„ «·„Õÿ… ›ﬁÿ
            STxt.Caption = SName
        End If
        SInd.ZOrder 0: STxt.ZOrder 0
    Next j
    
    frm9.Repaint
    
    ' ≈⁄«œ… «·ÃœÊ·…
    Sigen_NextPulse = Now + TimeValue("00:00:40")
    Application.OnTime Sigen_NextPulse, "sigen_connect.Sigen_Trigger"
End Sub

Public Sub Sigen_Trigger()
    Call Sigen_Main_Engine
End Sub
