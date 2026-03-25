VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm66 
   Caption         =   "UserForm66"
   ClientHeight    =   10740
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   7776
   OleObjectBlob   =   "UserForm66.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm66"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim folderPath As String


Private Sub CommandButton30_Click()
Unload Me
UserForm9.Show
End Sub





Private Sub UserForm_Initialize()
    folderPath = "D:\helpe\"
    With ListBox1
        .TextAlign = fmTextAlignRight
        .Font.Name = "Arial": .Font.Size = 22
        .ColumnCount = 2: .ColumnWidths = (.Width - 25) & ";0"
    End With
    Call LoadOnlyParents
End Sub

Sub LoadOnlyParents()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Helper")
    Dim i As Long: ListBox1.Clear
    For i = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If ws.Cells(i, 2).value = 0 Then
            ListBox1.AddItem "[+] " & ws.Cells(i, 1).value
            ListBox1.List(ListBox1.ListCount - 1, 1) = i
        End If
    Next i
End Sub

Private Sub ListBox1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Helper")
    Dim i As Long, StartIdx As Long, pRow As Long
    Dim SelectedText As String, RealName As String
    Dim VideoExts As Variant, ext As Variant
    
    ' «·„”«— «·„⁄ „œ
    Dim MyPath As String: MyPath = "D:\helpe\"

    If ListBox1.ListIndex = -1 Then Exit Sub
    StartIdx = ListBox1.ListIndex
    SelectedText = ListBox1.List(StartIdx, 0)
    
    ' --- 1. „‰ÿﬁ «·‘Ã—… (‰›” ≈Ã—«¡ «· Ê”Ì⁄ Ê«·ÿÌ) ---
    If Left(SelectedText, 3) = "[+]" Or Left(SelectedText, 3) = "[-]" Then
        pRow = val(ListBox1.List(StartIdx, 1))
        If Left(SelectedText, 3) = "[+]" Then
            ListBox1.List(StartIdx, 0) = "[-] " & Mid(SelectedText, 5)
            For i = pRow + 1 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
                If ws.Cells(i, 2).value = 1 Then
                    ListBox1.AddItem "  > " & ws.Cells(i, 1).value, StartIdx + 1
                    ListBox1.List(StartIdx + 1, 1) = "Child"
                    StartIdx = StartIdx + 1
                Else: Exit For: End If
            Next i
        Else
            ListBox1.List(StartIdx, 0) = "[+] " & Mid(SelectedText, 5)
            Do While StartIdx + 1 < ListBox1.ListCount
                If Left(ListBox1.List(StartIdx + 1, 0), 4) = "  > " Then ListBox1.RemoveItem StartIdx + 1 Else Exit Do
            Loop
        End If
        
    ' --- 2. „‰ÿﬁ › Õ «·„·›«  (‰›” ≈Ã—«¡ «·‹ TreeView Click) ---
    ElseIf Left(SelectedText, 4) = "  > " Then
        '  ‰ŸÌ› «·«”„: ≈“«·… "  > " ·«” —Ã«⁄ «”„ «·„·› ﬂ„« ÂÊ ›Ì «·„Ã·œ
        RealName = Trim(Mid(SelectedText, 5))
        
        On Error Resume Next
        ' › Õ «·ÊÊ—œ »«” Œœ«„ „‘€· «·‰Ÿ«„ «·«› —«÷Ì („À· «·‹ TreeView)
        If Dir(MyPath & RealName & ".docx") <> "" Then
            CreateObject("Shell.Application").Open (MyPath & RealName & ".docx")
        End If
        
        ' › Õ «·›ÌœÌÊ »Ã„Ì⁄ «„ œ«œ« Â
        VideoExts = Array("mp4", "avi", "wmv", "mov", "mkv", "mpg", "flv")
        For Each ext In VideoExts
            If Dir(MyPath & RealName & "." & ext) <> "" Then
                CreateObject("Shell.Application").Open (MyPath & RealName & "." & ext)
                Exit For
            End If
        Next ext
        On Error GoTo 0
    End If
End Sub
