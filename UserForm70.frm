VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm70 
   Caption         =   "UserForm70"
   ClientHeight    =   8976.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9984.001
   OleObjectBlob   =   "UserForm70.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm70"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private MyURL As String

Private Sub CommandButton1_Click()
On Error Resume Next
    ' 1. «· √ﬂœ „‰ «Œ Ì«— «·„’‰⁄
    If Me.ComboBox1.value = "" Then
        MsgBox "Ì—ÃÏ «Œ Ì«— «”„ «·„’‰⁄ √Ê·« · ›⁄Ì· «·»—Ê ÊﬂÊ·", vbExclamation
        Exit Sub
    End If
    
    ' 2.  Õ›Ì“ «·« ’«· «·›Ê—Ì (ÌœÊÌ«) »œÊ‰ «‰ Ÿ«— «·„ƒﬁ 
    Me.Label1.Caption = "Ã«—Ì  ›⁄Ì· «·»—Ê ÊﬂÊ· ÊÕﬁ‰ «·»Ì«‰« ..."
    Call Auto_Pulse_Engine
    
    
End Sub

Private Sub ListBox1_Click()
    On Error Resume Next
    Dim Index As Long: Index = Me.ListBox1.ListIndex
    If Index <= 0 Then Exit Sub

    ' 1.  ‰ŸÌ› «·ﬁœÌ„
    Me.Controls.Remove "TextBox3"
    
    ' 2. ≈‰‘«¡ «·„—»⁄ (TextBox3) „⁄ ‘—Ìÿ «· „—Ì— «·⁄„ÊœÌ
    Dim txt As Object: Set txt = Me.Controls.Add("Forms.TextBox.1", "TextBox3", True)
    With txt
        .Width = Me.ListBox1.Width + 40
        .Left = Me.ListBox1.Left - 20
        .Height = 65 ' «— ›«⁄ „ Ê”ÿ Ê„—ÌÕ
        .Top = Me.ListBox1.Top - 70
        
        ' --- ≈⁄œ«œ«   ⁄œœ «·√”ÿ— Ê«· „—Ì— «·⁄„ÊœÌ ---
        .Multiline = True
        .WordWrap = True
        .ScrollBars = 2        ' <---  ›⁄Ì· fmScrollBarsVertical
        .HideSelection = False
        .EnterKeyBehavior = False ' Enter ··≈—”«· Ê·Ì” ··‰“Ê·
        
        ' --- «· ‰”Ìﬁ «·„ÿ·Ê» ---
        .Font.Name = "Times New Roman": .Font.Size = 12: .Font.Bold = True
        .TextAlign = 3: .BackColor = RGB(255, 255, 248): .BorderStyle = 1
        .SetFocus
    End With

    ' 3. Ê÷⁄ «·“— (btnSend) ›Ì “«ÊÌ… «·„—»⁄
    With Me.btnSend
        .Visible = True
        .Caption = "≈—”«·"
        .Left = txt.Left + 5
        .Top = txt.Top + txt.Height - 22
        .Width = 45: .Height = 20
        .BackColor = RGB(0, 120, 215): .ForeColor = vbWhite
        .ZOrder 0
    End With
End Sub
Private Sub ListBox1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
End Sub
' https://script.google.com/macros/s/AKfycbz6_NLoHuiQMUaJKYwcvKcUCDKLec2bM3s0gNr0LaZYHccexKrtIHP5Q7O0JU5-c4ni/exec
Private Sub UserForm_Initialize()
    On Error Resume Next
    ' —«»ÿ «·”ﬂ—»  «·”Õ«»Ì
    MyURL = "https://script.google.com/macros/s/AKfycbwP6aq9Nk0T2IB8yDjVJyEIzBdeifSdCLdvrz_C1dXz-nBDL7xsfX-jRcHhNcdo1eSh/exec"
    
    '  ‰”Ìﬁ «·Ê«ÃÂ…
    Me.Label1.BackColor = RGB(255, 192, 203) ' «··Ê‰ «·Ê—œÌ
    Me.ListBox1.ColumnCount = 2
    Me.ListBox2.ColumnCount = 2
    Me.ListBox2.ColumnWidths = "140;80"
    Me.ListBox1.ColumnWidths = "100;80"
    ' √”„«¡ «·„’«‰⁄
    Me.ComboBox1.List = Array("„ﬁ— «·ÂÌ∆…", "„ﬁ— ‘—ﬂ… «·’‰«⁄«  «·Õ—»ÌÂ", "„’‰⁄ «·ﬂ—«„… Ê«·Õ«—À", "„’‰⁄ «·—»Ì⁄", "„’‰⁄ «·‰Â—Ê«‰", "„’‰⁄ Õ„Ê—«»Ì", "„’‰⁄ «·Ì—„Êﬂ", "„’‰⁄ ⁄ﬁ»… Ê»œ—", "„’‰⁄ «·ﬁ«œ”ÌÂ", "„’‰⁄ «·—‘Ìœ", "„’‰⁄ ÕÿÌ‰", "„’‰⁄ Ã«»— »‰ ÕÌ«‰", "„’‰⁄  »Êﬂ")
    
    Call Refresh_Employee_List
    Call StartTimer ' ≈ÿ·«ﬁ √Ê· ‰»÷… „ƒﬁ 
End Sub

' 2. „Õ—ﬂ «· “«„‰ «· ·ﬁ«∆Ì (Ì⁄„· ﬂ· 10 ÀÊ«‰Ì)


' 3. ÊŸÌ›… ÃœÊ·… «·„ƒﬁ 
Public Sub StartTimer()
    On Error Resume Next
    ' ≈·€«¡ √Ì  ÊﬁÌ  ﬁœÌ„ · Ã‰» «· œ«Œ·
    Application.OnTime NextTick, "TriggerTimer", , False
    
    ' ÷»ÿ «·Êﬁ  «·ÃœÌœ: 40 À«‰Ì…
    NextTick = Now + TimeValue("00:00:40")
    Application.OnTime NextTick, "TriggerTimer"
End Sub

' 4. ›· —… «·„ÊŸ›Ì‰ («·»ÕÀ ›Ì «·⁄„Êœ E „‰ «·’› 9)
Sub Refresh_Employee_List()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim i As Long, lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row
    Me.ListBox2.Clear
    If lastRow < 9 Then Exit Sub
    For i = 9 To lastRow
        If InStr(1, ws.Cells(i, "E").value, Me.TextBox1.Text, vbTextCompare) > 0 Or Me.TextBox1.Text = "" Then
            Me.ListBox2.AddItem ws.Cells(i, "E").value
            Me.ListBox2.List(Me.ListBox2.ListCount - 1, 1) = ws.Cells(i, "B").value
        End If
    Next i
End Sub

' 5. Õﬁ‰ «·»Ì«‰«  «·„” ·„… ›Ì ¬Œ— ”Ã·


' 6. ≈—”«· «·—”«·… ⁄‰œ «·÷€ÿ ⁄·Ï ListBox1


' 7. ≈—”«· «·ﬁÌœ ⁄‰œ «·÷€ÿ ⁄·Ï ListBox2
Private Sub ListBox2_Click()
    On Error Resume Next
    Dim Target As String: Target = Me.ListBox1.value
    If Target = "" Or Target = "«·„Õÿ… «·‰‘ÿ… («·¬‰)" Then MsgBox "Õœœ «·ÊÃÂ… √Ê·«", vbExclamation: Exit Sub
    
     TextBox2.Text = "":  TextBox2.ForeColor = vbBlue:  TextBox2.Font.Bold = True
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim r As Range: Set r = ws.Range("E9:E" & ws.Rows.count).Find(What:=Me.ListBox2.value, LookAt:=xlWhole)
    If Not r Is Nothing Then
        Dim Pack As String, c As Integer: Pack = "[RECORD_TRANSFER]|"
        For c = 1 To ws.Cells(r.row, ws.Columns.count).End(xlToLeft).Column: Pack = Pack & ws.Cells(r.row, c).value & "|": Next c
        Call Send_Data(Target, Pack)
        
        Me.TextBox2.Text = "[" & Format(Now, "hh:mm") & "]  „ ‰ﬁ· ﬁÌœ " & Me.ListBox2.value & " ≈·Ï " & Target: Beep
        
        'Me.Label2.Caption = "[" & Format(Now, "hh:mm") & "]  „ ‰ﬁ· ﬁÌœ " & Me.ListBox2.value & " ≈·Ï " & Target: Beep
    End If
End Sub

' 8. —«œ«— «·„’«‰⁄
Sub Peer_Discovery_Scan()
    On Error Resume Next
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", MyURL, False: http.Send
    Dim r As String: r = http.responseText: Me.ListBox1.Clear: Me.ListBox1.AddItem "«·„Õÿ… «·‰‘ÿ… («·¬‰)"
    Dim i As Integer: For i = 0 To Me.ComboBox1.ListCount - 1
        If InStr(r, Chr(34) & Me.ComboBox1.List(i) & Chr(34)) > 0 Then
            Me.ListBox1.AddItem Me.ComboBox1.List(i): Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = "√Ê‰·«Ì‰"
        End If
    Next i
End Sub

' 9. ≈Ìﬁ«› «·„ƒﬁ  ⁄‰œ «·Œ—ÊÃ
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    Application.OnTime NextTick, "TriggerTimer", , False
End Sub

' ≈Ã—«¡«  ›—⁄Ì… ≈÷«›Ì…
Private Sub TextBox1_Change(): On Error Resume Next: Call Refresh_Employee_List: End Sub
Sub Send_Data(TargetFact As String, DataPack As String)
    On Error Resume Next
    Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", MyURL, False: http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send "action=SEND_MSG&from=" & WorksheetFunction.EncodeURL(Me.ComboBox1.value) & "&to=" & WorksheetFunction.EncodeURL(TargetFact) & "&msg=" & WorksheetFunction.EncodeURL(DataPack)
    
End Sub
Private Sub CommandButton2_Click(): Unload Me: End Sub


Public Sub Auto_Pulse_Engine()
    On Error Resume Next
    Dim CurrentStation As String: CurrentStation = Me.ComboBox1.value
    
    ' 1.  ‰‘Ìÿ "«·–«ﬂ—… «·„—∆Ì…" Ê„”Õ «·„Õ ÊÏ «·”«»ﬁ „‰ «··Ì»· ﬁ»· √Ì ⁄„·Ì… ÃœÌœ…
    
    
    If CurrentStation <> "" Then
        Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        
        ' ≈—”«· Õ«·… «·‰‘«ÿ
        http.Open "POST", MyURL, False
        http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.Send "station=" & WorksheetFunction.EncodeURL(CurrentStation) & "&status=Active"
        
        ' Ã·» «·—”«∆· «·ÃœÌœ…
        http.Open "GET", MyURL & "?action=GET_MSGS&myname=" & WorksheetFunction.EncodeURL(CurrentStation), False
        http.Send
        
        Dim resp As String: resp = http.responseText
        
        ' ≈–« ÊÃœ  »Ì«‰«  ÃœÌœ… Ì „ Õﬁ‰Â«° Ê≈·« ”Ì»ﬁÏ «··Ì»· ›«—€« (‰ŸÌ›«)
        If resp <> "" And InStr(resp, "EMPTY") = 0 Then
            Call Inject_Data(resp)
        End If
        
        Call Peer_Discovery_Scan
    End If
    
    ' ≈⁄«œ… ÃœÊ·… «·‰»÷… «· «·Ì…
    Call StartTimer
End Sub
Sub Inject_Data(RawData As String)
    On Error Resume Next
    Call Sync_Memory_UI(RawData)
    If InStr(RawData, "[CHAT_MSG]|") > 0 Then
        Dim Parts() As String: Parts = Split(RawData, "|")
        Me.TextBox2.Text = "—”«·… „‰ " & Parts(1) & ": " & Parts(2)
        Me.TextBox2.ForeColor = vbBlue:  TextBox2.Font.Bold = True
        Beep
        Exit Sub
    End If
    
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Sheet1")
    Dim NextRow As Long: NextRow = ws.Cells(ws.Rows.count, "E").End(xlUp).row + 1
    If NextRow < 9 Then NextRow = 9
    Dim arr() As String: arr = Split(Mid(RawData, InStr(RawData, "|") + 1), "|")
    Dim j As Integer: For j = 0 To UBound(arr) - 1: ws.Cells(NextRow, j + 1).value = arr(j): Next j
Beep:
Call Sync_Memory_UI(RawData)
End Sub

Public Sub Sync_Memory_UI(RawData As String)
    On Error Resume Next
    
    ' ≈–« ﬂ«‰  «·»Ì«‰«   Õ ÊÌ ⁄·Ï Ê”Ê„ «·≈—”«· (œ—œ‘… √Ê ﬁÌœ)
    If InStr(RawData, "[CHAT_MSG]|") > 0 Or InStr(RawData, "[RECORD_TRANSFER]|") > 0 Then
        
        ' 1.  ‰‘Ìÿ «·–«ﬂ—…: „”Õ «··Ì»· ·«” ﬁ»«· «·—”«·… «·ÃœÌœ…
        TextBox2.Text = ""
        
        ' 2.  ÊÃÌÂ «·—”«·… ··Ê«ÃÂ… »ÿ—Ìﬁ… ’ÕÌÕ… (»œÊ‰ „”Ã »Êﬂ”)
        If InStr(RawData, "[CHAT_MSG]|") > 0 Then
            Dim Parts() As String: Parts = Split(RawData, "|")
           TextBox2.Text = "?? —”«·… „‰ " & Parts(1) & ": " & Parts(2)
        Else
            Label1.Caption = "?  „ «” ·«„ ÊÕﬁ‰ ﬁÌœ ÃœÌœ  ·ﬁ«∆Ì«"
        End If
        
        Beep
    End If
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    
    ' «· Õﬁﬁ ≈–« ﬂ«‰ «·÷€ÿ œ«Œ· TextBox3 Êﬂ«‰ «·„› «Õ ÂÊ Enter
    If Me.ActiveControl.Name = "TextBox3" And KeyCode = 13 Then
        Dim UserMsg As String: UserMsg = Me.ActiveControl.Text
        Dim Target As String: Target = Me.ListBox1.value
        Dim MyStation As String: MyStation = Me.ComboBox1.value
        
        If UserMsg <> "" Then
            ' --- »—Ê ÊﬂÊ· «·≈—”«· «·”Õ«»Ì (ﬂÊœﬂ «·√’·Ì) ---
            Dim FullMsg As String: FullMsg = "[CHAT_MSG]|" & MyStation & "|" & UserMsg
            Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
            
            http.Open "POST", MyURL, False
            http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            http.Send "action=SEND_MSG&from=" & WorksheetFunction.EncodeURL(MyStation) & _
                      "&to=" & WorksheetFunction.EncodeURL(Target) & "&msg=" & WorksheetFunction.EncodeURL(FullMsg)
            
            '  ÕœÌÀ ”Ã· «·—”«∆· (TextBox2)
            If InStr(http.responseText, "MSG_SENT") > 0 Then
                Me.TextBox2.Text = "[" & Format(Now, "hh:mm:ss") & "] „‰ " & MyStation & " ≈·Ï " & Target & ": " & UserMsg & vbCrLf & Me.TextBox2.Text
                Beep
            End If
        End If
        
        ' ≈Œ›«¡ «·’‰œÊﬁ »⁄œ «·≈—”«·
        Me.Controls.Remove ("TextBox3")
    End If
End Sub

Private Sub btnSend_Click()
    On Error Resume Next
    
    ' 1. Ã·» «·»Ì«‰«  „‰ «· Ìﬂ”  »Êﬂ” «·„‰‘√ »—„ÃÌ«
    Dim txt As Object: Set txt = Me.Controls("TextBox3")
    Dim UserMsg As String: UserMsg = txt.Text
    Dim Target As String: Target = Me.ListBox1.value
    Dim MyStation As String: MyStation = Me.ComboBox1.value
    
    If UserMsg <> "" Then
        ' --- »—Ê ÊﬂÊ· «·≈—”«· «·”Õ«»Ì (ﬂÊœﬂ «·√’·Ì) ---
        Dim FullMsg As String: FullMsg = "[CHAT_MSG]|" & MyStation & "|" & UserMsg
        Dim http As Object: Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.Open "POST", MyURL, False
        http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.Send "action=SEND_MSG&from=" & WorksheetFunction.EncodeURL(MyStation) & _
                  "&to=" & WorksheetFunction.EncodeURL(Target) & "&msg=" & WorksheetFunction.EncodeURL(FullMsg)
        
        '  ÕœÌÀ «·”Ã· (TextBox2)
        If InStr(http.responseText, "MSG_SENT") > 0 Then
            Me.TextBox2.Text = "[" & Format(Now, "hh:mm") & "] √‰« -> " & Target & ": " & UserMsg & vbCrLf & Me.TextBox2.Text
            Beep
        End If
    End If
    
    ' 2. «· ŸÌ› «·‰Â«∆Ì (√Â„ ŒÿÊ…)
    Me.btnSend.Visible = False  ' <--- ≈Œ›«¡ «·“— «·ÕﬁÌﬁÌ
    Me.Controls.Remove "TextBox3" ' <--- Õ–› «·„—»⁄ «·»—„ÃÌ
End Sub

