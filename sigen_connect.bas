Attribute VB_Name = "sigen_connect"
'https://script.google.com/macros/s/AKfycbz8aNEkZQ6TOHRkrQgXVBweTe21dT3C0FN4n2kOfnckdSR2QY1oVPm31CafK-1YSQUK/exec





Public Sigen_NextPulse As Date

' «·„Õ—ﬂ «·√”«”Ì ··‰Ÿ«„ - «·—«»ÿ „À»  œ«∆„Ì« Ê·« Ì €Ì— √»œ«
Public Sub Sigen_Main_Engine()
    On Error Resume Next
   Dim frm9 As Object: Set frm9 = UserForm9
    
    ' --- «·≈÷«›… «·ÃœÌœ… Â‰« ---
    ' ≈–« ﬂ«‰ Â–« ÂÊ «· ‘€Ì· «·√Ê· ⁄‰œ › Õ «·›Ê—„° «ÿ·» «· ÕœÌÀ »⁄œ 5 ÀÊ«‰Ì Ê«Œ—Ã ›Ê—«
    Static FirstRun As Boolean
    If FirstRun = False Then
        FirstRun = True
        Sigen_NextPulse = Now + TimeValue("00:00:05") ' ”Ì»œ√ «·⁄„· «·›⁄·Ì »⁄œ 5 ÀÊ«‰Ì „‰ «·› Õ
        Application.OnTime Sigen_NextPulse, "sigen_connect.Sigen_Trigger"
        Exit Sub
    End If
  '====================================================================
    
    ' «·—«»ÿ «·ﬂ«„· «·„À»  (·« Ì €Ì—)
    Dim Sigen_URL As String
    Sigen_URL = "https://script.google.com/macros/s/AKfycbz8aNEkZQ6TOHRkrQgXVBweTe21dT3C0FN4n2kOfnckdSR2QY1oVPm31CafK-1YSQUK/exec"
    
    ' 1. ﬁ—«¡… «”„ „Õÿ ﬂ «·Õ«·Ì „‰ «·‘Ì  (title_factory «·Œ·Ì… A2)
    Dim Sigen_MyName As String
    Sigen_MyName = Trim(ThisWorkbook.Sheets("title_factory").Range("A2").value)
    
    ' 2. ≈—”«· ‰»÷… «·Õ«·… (POST) · ”ÃÌ· ÊÃÊœﬂ ›Ì «·”Õ«»
    Dim Sigen_Http As Object: Set Sigen_Http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If Sigen_MyName <> "" Then
        Sigen_Http.Open "POST", Sigen_URL, False
        Sigen_Http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        Sigen_Http.Send "station=" & Application.WorksheetFunction.EncodeURL(Sigen_MyName) & "&status=Active&t=" & Timer
    End If
    
    ' 3. Ã·» »Ì«‰«  «·—«œ«— (GET) „⁄ ﬂ”— «·ﬂ«‘ ·÷„«‰ «·„’œ«ﬁÌ…
    Sigen_Http.Open "GET", Sigen_URL & "?action=GET_ACTIVE&t=" & Timer & "&rnd=" & Int((1000 * Rnd) + 1), False
    Sigen_Http.Send
    Dim Sigen_Response As String: Sigen_Response = Sigen_Http.responseText
    
    ' 4. „’›Ê›… «·„Õÿ«  «·—”„Ì… «·‹ 14 ( ‘„· «·ﬁ⁄ﬁ«⁄)
    Dim Sigen_List As Variant
    Sigen_List = Array("„ﬁ— «·ÂÌ∆…", "„ﬁ— ‘—ﬂ… «·’‰«⁄«  «·Õ—»ÌÂ", "„’‰⁄ «·ﬂ—«„… Ê«·Õ«—À", "„’‰⁄ «·—»Ì⁄", "„’‰⁄ «·‰Â—Ê«‰", "„’‰⁄ Õ„Ê—«»Ì", "„’‰⁄ «·Ì—„Êﬂ", "„’‰⁄ ⁄ﬁ»… Ê»œ—", "„’‰⁄ «·ﬁ«œ”ÌÂ", "„’‰⁄ «·—‘Ìœ", "„’‰⁄ ÕÿÌ‰", "„’‰⁄ Ã«»— »‰ ÕÌ«‰", "„’‰⁄  »Êﬂ", "„’‰⁄ «·ﬁ⁄ﬁ«⁄")
    
    ' 5.  ‰ŸÌ› «·⁄·«„«  «·”«»ﬁ… »—„ÃÌ« (Õ–› ⁄ﬂ”Ì ·÷„«‰ «·œﬁ…)
    Dim iCtrl As Integer
    For iCtrl = frm9.Controls.count - 1 To 0 Step -1
        If Left(frm9.Controls(iCtrl).Name, 7) = "SigInd_" Or Left(frm9.Controls(iCtrl).Name, 7) = "SigTxt_" Then
            frm9.Controls.Remove frm9.Controls(iCtrl).Name
        End If
    Next iCtrl
    
    ' 6. »‰«¡ «·—«œ«— «·‘«„· (⁄—÷ «·‹ 14 „Õÿ… œ«∆„«)
    Dim j As Integer
    Dim Sigen_TopPos As Single: Sigen_TopPos = 142 ' ≈“«Õ… 5 ”„ ·√”›·

    For j = 0 To UBound(Sigen_List)
        Dim SName As String: SName = Sigen_List(j)
        
        ' »‰«¡ «·œ«∆—… (Indicator)
        Dim SInd As Object: Set SInd = frm9.Controls.Add("Forms.Label.1", "SigInd_" & j, True)
        SInd.Left = 10: SInd.Width = 10: SInd.Height = 10
        SInd.Top = Sigen_TopPos + (j * 18)
        SInd.BorderStyle = 1: SInd.Caption = ""
        
        ' »‰«¡ «·‰’ (Station Name)
        Dim STxt As Object: Set STxt = frm9.Controls.Add("Forms.Label.1", "SigTxt_" & j, True)
        STxt.Left = 25: STxt.Width = 180: STxt.Height = 15
        STxt.Top = SInd.Top - 1
        STxt.Caption = SName: STxt.BackStyle = 0: STxt.Font.Size = 9
        
        ' --- „‰ÿﬁ «·√·Ê«‰ «·’«—„ («·„’œ«ﬁÌ…) ---
        ' ·‰ Ì ·Ê‰ »«·√’›— ≈·« ≈–« ﬂ«‰ «·«”„ „ÊÃÊœ« ›Ì "—œ «·”Õ«»" «·›⁄·Ì
        If InStr(1, Sigen_Response, Chr(34) & SName & Chr(34), vbTextCompare) > 0 Then
            SInd.BackColor = vbYellow ' √’›— („ ’· ”Õ«»Ì«)
            STxt.Font.Bold = True: STxt.ForeColor = vbBlack
        Else
            SInd.BackColor = vbRed    ' √Õ„— (€Ì— „ ’· √Ê „ÿ›Ì)
            STxt.Font.Bold = False: STxt.ForeColor = RGB(170, 170, 170)
        End If
        
        SInd.ZOrder 0: STxt.ZOrder 0
    Next j
    
    frm9.Repaint
    
    ' 7. ≈⁄«œ… «·ÃœÊ·… «· ·ﬁ«∆Ì… ﬂ· 40 À«‰Ì…
    Sigen_NextPulse = Now + TimeValue("00:00:40")
    Application.OnTime Sigen_NextPulse, "sigen_connect.Sigen_Trigger"
End Sub

Public Sub Sigen_Trigger()
    Call Sigen_Main_Engine
End Sub
