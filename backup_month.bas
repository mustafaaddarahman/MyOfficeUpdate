Attribute VB_Name = "backup_month"
Sub MonthlyBackupManager_DriveD()
    Dim fso As Object
    Dim rootDrive As String, backupRoot As String, backupPath As String, logFile As String
    Dim currentYear As Integer, currentMonth As Integer, currentDay As Integer
    Dim daysInMonth As Integer, lastFiveDaysStart As Integer
    Dim answer As VbMsgBoxResult
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. ÊÍÏíÏ ÇáÊÇÑíÎ æÇáÅÚÏÇÏÇÊ
    currentYear = Year(Date)
    currentMonth = Month(Date)
    currentDay = Day(Date)
    daysInMonth = Day(DateSerial(currentYear, currentMonth + 1, 0))
    lastFiveDaysStart = daysInMonth - 4
    
    ' ÇáãÓÇÑ Úáì ÇáÞÑÕ D
    rootDrive = "D:\"
    backupRoot = rootDrive & "ÇáäÓÎÉ ÇáÇÍÊíÇØíÉ áÌãíÚ ÇÔåÑ " & currentYear
    
    ' 2. ÇáÊÍÞÞ ãä ÇáÝÊÑÉ ÇáÒãäíÉ (ÂÎÑ 5 ÃíÇã)
    If currentDay >= lastFiveDaysStart And currentDay <= daysInMonth Then
        
        ' ÇáÊÃßÏ ãä æÌæÏ ÇáÞÑÕ D ÃæáÇð áÊÌäÈ ÇáÎØÃ
        If Not fso.DriveExists("D") Then
            MsgBox "ÇáÞÑÕ D ÛíÑ ãæÌæÏ¡ íÑÌì ÇáÊÍÞÞ ãä ÇáãÓÇÑ.", vbCritical
            Exit Sub
        End If

        ' ÅäÔÇÁ ÇáãÌáÏ ÇáÑÆíÓí ÅÐÇ áã íßä ãæÌæÏÇð
        If Not fso.FolderExists(backupRoot) Then fso.CreateFolder (backupRoot)
        
        ' ãáÝ ÇáÊÍÞÞ áãäÚ ÇáÊßÑÇÑ (Flag)
        logFile = backupRoot & "\done_" & currentMonth & ".log"
        
        ' 3. ÅÐÇ áã íæÇÝÞ ÇáãÓÊÎÏã ãÓÈÞÇð åÐÇ ÇáÔåÑ
        If Not fso.FileExists(logFile) Then
            answer = MsgBox("åá ÊÑíÏ ÇáÂä Úãá äÓÎÉ ÇÍÊíÇØíÉ ÔåÑíÉ¿", vbQuestion + vbOKCancel, "ÊäÈíå ÇáäÓÎ ÇáÇÍÊíÇØí")
            
            If answer = vbOK Then
                ' ÇÓã ÇáäÓÎÉ ãÚ ÇáÔåÑ æÇáÊÇÑíÎ
                backupPath = backupRoot & "\äÓÎÉ ÔåÑ " & monthName(currentMonth) & " - ÈÊÇÑíÎ " & Format(Date, "yyyy-mm-dd") & ".xlsm"
                
                ' ÊäÝíÐ ÇáäÓÎ
                ThisWorkbook.SaveCopyAs backupPath
                
                ' ÅäÔÇÁ ãáÝ ÚáÇãÉ ÇáÅÊãÇã áãäÚ ÇáÓÄÇá ãÌÏÏÇð
                fso.CreateTextFile(logFile).WriteLine "Backup Completed on " & Now
                
                MsgBox "Êã ÍÝÙ ÇáäÓÎÉ ÈäÌÇÍ Ýí: " & vbCrLf & backupPath, vbInformation, "äÌÍÊ ÇáÚãáíÉ"
            End If
        End If
    End If
End Sub
