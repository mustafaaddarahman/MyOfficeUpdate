Attribute VB_Name = "record_varible"
Option Explicit

' 1. „ —Ã„ «·√⁄„œ… «·œﬁÌﬁ (·—»ÿ ⁄‰«’— «·‰„Ê–Ã »’›Ê› «·»Ì«‰« )
Private Function GetColIndex(ByVal CtrlName As String) As Variant
    Select Case CtrlName
        Case "TextBox1": GetColIndex = 1: Case "TextBox2": GetColIndex = 2: Case "TextBox3": GetColIndex = 3
        Case "ComboBox4": GetColIndex = 4: Case "TextBox5": GetColIndex = 5: Case "ComboBox5": GetColIndex = 6
        Case "TextBox7": GetColIndex = 7
        Case "TextBox8": GetColIndex = 8  ' «·œ—Ã…
        Case "TextBox9": GetColIndex = 9  ' «·„—Õ·…
        Case "TextBox10": GetColIndex = 10: Case "TextBox11": GetColIndex = 11: Case "TextBox12": GetColIndex = 12
        Case "ComboBox6": GetColIndex = 13: Case "TextBox14": GetColIndex = 14: Case "ComboBox7": GetColIndex = 15
        Case "TextBox16": GetColIndex = 16: Case "ComboBox8": GetColIndex = 17: Case "TextBox18": GetColIndex = 18
        Case "TextBox19": GetColIndex = 19: Case "TextBox20": GetColIndex = 20: Case "TextBox21": GetColIndex = 21
        Case "TextBox22": GetColIndex = 22: Case "TextBox23": GetColIndex = 23: Case "TextBox24": GetColIndex = 24
        Case "TextBox25": GetColIndex = 25: Case "ComboBox10": GetColIndex = 26: Case "TextBox27": GetColIndex = 27
        Case "ComboBox11": GetColIndex = 28: Case "TextBox29": GetColIndex = 29: Case "TextBox30": GetColIndex = 30
        Case "TextBox31": GetColIndex = 31: Case "TextBox32": GetColIndex = 32: Case "TextBox33": GetColIndex = 33
        Case "TextBox34": GetColIndex = 34: Case "TextBox35": GetColIndex = 35: Case "TextBox36": GetColIndex = 36
        Case "TextBox37": GetColIndex = 37: Case "TextBox38": GetColIndex = 38: Case "TextBox39": GetColIndex = 39
        Case "TextBox40": GetColIndex = 40: Case "TextBox41": GetColIndex = 41: Case "TextBox42": GetColIndex = 42
        Case "TextBox43": GetColIndex = 43: Case "TextBox44": GetColIndex = 44: Case "TextBox45": GetColIndex = 45
        Case "TextBox46": GetColIndex = 46: Case "TextBox47": GetColIndex = 47: Case "TextBox48": GetColIndex = 48
        Case "TextBox49": GetColIndex = 49: Case "TextBox50": GetColIndex = 50: Case "ComboBox9": GetColIndex = 51
        Case "TextBox52": GetColIndex = 52: Case "TextBox53": GetColIndex = 53: Case "TextBox54": GetColIndex = 54
        Case "TextBox55": GetColIndex = 55: Case "TextBox56": GetColIndex = 56: Case "TextBox57": GetColIndex = 57
        Case "TextBox58": GetColIndex = 58: Case "TextBox59": GetColIndex = 59: Case "TextBox60": GetColIndex = 60
        Case "TextBox61": GetColIndex = 61: Case "TextBox62": GetColIndex = 62: Case "TextBox63": GetColIndex = 63
        Case "TextBox64": GetColIndex = 64: Case "TextBox82": GetColIndex = 66: Case "TextBox83": GetColIndex = 67
        Case "TextBox84": GetColIndex = 68: Case "TextBox85": GetColIndex = 69: Case "TextBox86": GetColIndex = 70
        Case "TextBox87": GetColIndex = 71: Case "TextBox88": GetColIndex = 72: Case "TextBox91": GetColIndex = 74
        Case "TextBox99": GetColIndex = "BU": Case "TextBox100": GetColIndex = "CA"
        Case Else: GetColIndex = 0
    End Select
End Function

' 2. «· ﬁ«ÿ Õ«·… «·»Ì«‰«  «·√’·Ì… (Ìı” œ⁄Ï ⁄‰œ › Õ «·‰„Ê–Ã √Ê «Œ Ì«— „ÊŸ›)
Public Sub CaptureSnapshot(ByRef frm As Object, ByRef dict As Object)
    Set dict = CreateObject("Scripting.Dictionary")
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.TextBox Or TypeOf ctrl Is MSForms.ComboBox Then
            dict(ctrl.Name) = ctrl.value
        End If
    Next ctrl
End Sub

' 3. „Õ—ﬂ  —ÕÌ· «· €ÌÌ—«  ≈·Ï BlackBox (»œÊ‰ Õ›Ÿ  ·ﬁ«∆Ì ··„·›)
Public Sub RecordToBlackBox(ByRef frm As Object, ByRef dict As Object)
    Dim wsData As Worksheet: Set wsData = Sheets(1)
    Dim wsLog As Worksheet: Set wsLog = Sheets("BlackBox")
    Dim wsAdmin As Worksheet: Set wsAdmin = Sheets("title_factory")
    Dim ctrl As Control, colIdx As Variant, lr As Long
    Dim ChangeDetected As Boolean: ChangeDetected = False
    
    ' «· √ﬂœ „‰ ÊÃÊœ »Ì«‰«  √’·Ì… ··„ﬁ«—‰…
    If dict Is Nothing Then Exit Sub

    For Each ctrl In frm.Controls
        If dict.exists(ctrl.Name) Then
            ' «· Õﬁﬁ „‰ ÊÃÊœ  €ÌÌ— ( Ã«Â· «·„”«›«  «·“«∆œ…)
            If Trim(ctrl.value) <> Trim(dict(ctrl.Name)) Then
                colIdx = GetColIndex(ctrl.Name)
                
                ' ≈–« ﬂ«‰ «·⁄‰’— „— »ÿ« »⁄„Êœ ›Ì Ê—ﬁ… «·»Ì«‰« 
                If colIdx <> 0 Then
                    lr = wsLog.Cells(wsLog.Rows.count, 1).End(xlUp).row + 1
                    With wsLog
                        .Cells(lr, 1) = frm.TextBox2.value             ' «·—ﬁ„ «·ÊŸÌ›Ì
                        .Cells(lr, 2) = frm.TextBox5.value             ' «”„ «·„ÊŸ›
                        .Cells(lr, 3) = wsData.Cells(8, colIdx).value  ' «”„ «·Õﬁ· „‰ Ê—ﬁ… «·»Ì«‰« 
                        .Cells(lr, 4) = dict(ctrl.Name)                ' «·ﬁÌ„… «·ﬁœÌ„…
                        .Cells(lr, 5) = ctrl.value                     ' «·ﬁÌ„… «·ÃœÌœ…
                        .Cells(lr, 6) = Format(Now, "hh:mm:ss AM/PM")  ' «·Êﬁ 
                        .Cells(lr, 7) = Format(Date, "dddd")           ' «·ÌÊ„
                        .Cells(lr, 8) = Format(Date, "mmmm")           ' «·‘Â—
                        .Cells(lr, 9) = Year(Date)                     ' «·”‰…
                        .Cells(lr, 10) = wsAdmin.Range("C2").value     ' «·„”ƒÊ· «·Õ«·Ì
                    End With
                    
                    '  ÕœÌÀ «·ﬁÌ„… ›Ì «·ﬁ«„Ê” ·„‰⁄ «· —ÕÌ· «·„ﬂ—— ·‰›” «· ⁄œÌ·
                    dict(ctrl.Name) = ctrl.value
                    ChangeDetected = True
                End If
            End If
        End If
    Next ctrl
    
    ' ≈ŸÂ«— —”«·… «·‰ ÌÃ… ··„” Œœ„
    If ChangeDetected Then
        MsgBox " „  ⁄„·Ì…  —ÕÌ· „«  „  €ÌÌ—Â »‰Ã«Õ ≈·Ï ”Ã· «·‰Ÿ«„.", vbInformation + vbMsgBoxRight, " √ﬂÌœ «· —ÕÌ·"
    Else
        MsgBox "·„ Ì „ «ﬂ ‘«› √Ì  €ÌÌ—«  ÃœÌœ… · —ÕÌ·Â«.", vbExclamation + vbMsgBoxRight, " ‰»ÌÂ"
    End If
End Sub
