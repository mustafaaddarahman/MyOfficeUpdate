Attribute VB_Name = "export_all_data54"
Sub export_all_data_2()
On Error Resume Next
Sheets(54).Range("bw8:ci10000").EntireColumn.Hidden = True
Dim PASS As Variant
Dim lastR, i As Integer
lastR = Sheets(45).Cells(Rows.count, 2).End(xlUp).row
For i = 1 To lastR
PASS = Sheets(45).Cells(i, 2).value

Next i


    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(54).Range("a2:ci7").ClearContents
    Sheets(54).Range("a2:ci7").Merge
    Range("a2:ci7").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(54).Range("a2:ci7").Font.Name = "Arial"
   Sheets(54).Range("a2").Font.Bold = True
       
     Sheets(54).Range("a2:c17").ClearContents
    Sheets(54).Range("a2:ci7").HorizontalAlignment = xlCenter
     Sheets(54).Range("a2:ci7").VerticalAlignment = xlCenter
       
    Sheets(54).Range("a2:ci7").value = UserForm36.ComboBox1.value & " " & "«·»Ì«‰«  «·ﬂ·Ì… ·’—›Ì«  «·—« » " & "   " & "·‘Â—" & "    " & monthName(Month(Date)) & "   " & Sheets _
    (1).Cells(9, "be").value
    
    'Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(54)
   
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Application.Visible = True
             Application.ActiveWorkbook.SaveAs fileName:="D:\export\" & " " & " ›Ê—„… «·»Ì«‰«  «·ﬂ·Ì…" & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
             Application.Workbooks(2).Close True
            
        End With
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False

End Sub

