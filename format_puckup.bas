Attribute VB_Name = "format_puckup"
Sub BackupAndPrepareTw6een()
    Dim ws As Worksheet
    Dim backupPath As String
    Dim backupFile As String
    Dim lastRow As Long
    Dim frmLoaded As Boolean
    Dim frm As Object
    
    Set ws = ThisWorkbook.Sheets("tw6een")
    
    ' «· Õﬁﬁ Â· «·›Ê—„ „› ÊÕ
    frmLoaded = False
    For Each frm In VBA.UserForms
        If frm.Name = "UserForm62" Then
            frmLoaded = True
            Exit For
        End If
    Next frm
    
    ' ¬Œ— ’› »Ì«‰« 
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    ' ≈–« ·«  ÊÃœ »Ì«‰«  „‰ «·’› 9
    If lastRow < 9 Then
        MsgBox "‰Ÿ«„ «· ÊÿÌ‰ ·« ÌÕ ÊÌ ⁄·Ï √Ì »Ì«‰«  Õ«·Ì«." & vbCrLf & _
               "”Ì „ «·«” Ì—«œ Õ«·Ì«.", vbInformation
               
        ' ≈ŸÂ«— «·›Ê—„ ›ﬁÿ ≈–« ·„ Ìﬂ‰ „› ÊÕ
        If Not frmLoaded Then UserForm62.Show
        Exit Sub
    End If
    
    ' „”«— «·‰”Œ… «·«Õ Ì«ÿÌ…
    backupPath = "D:\export1\"
    backupFile = backupPath & "‰”Œ…_«Õ Ì«ÿÌ…_·»Ì«‰« _«· ÊÿÌ‰_" & _
                 Format(Now, "yyyymmdd_HHMMSS") & ".xlsx"
    
    ' ≈‰‘«¡ ‰”Œ… «Õ Ì«ÿÌ…
    ws.Copy
    ActiveWorkbook.SaveAs backupFile, xlOpenXMLWorkbook
    ActiveWorkbook.Close False
    
    ' „”Õ «·»Ì«‰«  „‰ «·’› 9
    ws.Rows("9:" & lastRow).ClearContents
    
    MsgBox " „ ≈‰‘«¡ ‰”Œ… «Õ Ì«ÿÌ… Ê„”Õ «·»Ì«‰«  »‰Ã«Õ.", vbInformation
    
    ' ≈ŸÂ«— «·›Ê—„ ›ﬁÿ ≈–« ·„ Ìﬂ‰ „› ÊÕ
    If Not frmLoaded Then UserForm62.Show
End Sub

Sub ReloadTw6eenForm()
    ' ≈€·«ﬁ «·›Ê—„ ≈–« ﬂ«‰ „› ÊÕ
    Dim uf As Object
    For Each uf In VBA.UserForms
        If uf.Name = "UserForm62" Then
            Unload uf
            Exit For
        End If
    Next uf
    
    ' ≈⁄«œ… › Õ «·›Ê—„
    UserForm62.Show
End Sub

