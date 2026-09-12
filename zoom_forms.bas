Attribute VB_Name = "zoom_forms"
Sub FixAndCleanInjection()
    Dim vbComp As Object, CodeMod As Object
    Dim i As Long, LineCount As Long
    
    ' ÇáÊÓãíÇÊ ÇáãËÈÊÉ ááÍŞä
    Dim Declaration As String: Declaration = "Dim Zoomer As New clsZoomManager"
    Dim BindCall As String: BindCall = "    Zoomer.Bind Me, Me.SpinButton1, Me.az"
    
    On Error Resume Next
    
    ' ÇáãÑæÑ Úáì ßÇİÉ ÇáãßæäÇÊ (UserForms)
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = 3 Then ' UserForm
            Set CodeMod = vbComp.CodeModule
            LineCount = CodeMod.CountOfLines
            
            ' 1. ÇáÊäÙíİ: ÍĞİ Ãí ÓØÑ ÇÓÊÏÚÇÁ ŞÏíã ãæÌæÏ İí "ÇáããÑ" (ÎÇÑÌ ÇáÜ Sub)
            ' åĞÇ íãäÚ ÎØÃ Outside Procedure ÊãÇãÇğ
            For i = LineCount To 1 Step -1
                Dim CurrentLine As String
                CurrentLine = Trim(CodeMod.Lines(i, 1))
                If InStr(CurrentLine, "Zoomer.Bind") > 0 Or InStr(CurrentLine, "Zoomer.Init") > 0 Then
                    CodeMod.DeleteLines i, 1
                End If
            Next i

            ' 2. ÍŞä ÇáÊÚÑíİ: íæÖÚ İí Ãæá ÓØÑ ÈÇáãæÏíæá (General Declarations)
            If Not CodeMod.Find(Declaration, 1, 1, -1, -1) Then
                CodeMod.InsertLines 1, Declaration
            End If
            
            ' 3. ÍŞä ÇáÇÓÊÏÚÇÁ: ÏÇÎá ÍÏË Initialize ÍÕÑÇğ
            Dim StartLine As Long
            StartLine = CodeMod.ProcBodyLine("UserForm_Initialize", 0)
            
            If StartLine > 0 Then
                ' ÅĞÇ æÌÏ ÇáÜ Sub¡ äÖÚ ÇáßæÏ İí Ãæá ÓØÑ ÊäİíĞí ÈÏÇÎáå áíÚãá Me ÈÔßá ÕÍíÍ
                CodeMod.InsertLines StartLine, BindCall
            Else
                ' ÅĞÇ áã íÌÏ ÇáÅÌÑÇÁ¡ íäÔÆå ãä ÇáÕİÑ İí äåÇíÉ ÇáßæÏ
                LineCount = CodeMod.CountOfLines + 1
                CodeMod.InsertLines LineCount, ""
                CodeMod.InsertLines LineCount + 1, "Private Sub UserForm_Initialize()"
                CodeMod.InsertLines LineCount + 2, BindCall
                CodeMod.InsertLines LineCount + 3, "End Sub"
            End If
        End If
    Next vbComp
    
    MsgBox "ÊãÊ ÇáãÚÇáÌÉ ÈäÌÇÍ!" & vbCrLf & "ÇáãÓãíÇÊ: az æ SpinButton1" & vbCrLf & "ÇáãæÖÚ: ÏÇÎá Initialize", vbInformation
End Sub
