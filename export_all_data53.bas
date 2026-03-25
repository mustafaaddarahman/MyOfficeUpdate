Attribute VB_Name = "export_all_data53"
Sub export_all_data_1()
On Error Resume Next
Sheets(53).Range("bw8:ci10000").EntireColumn.Hidden = True
Dim PASS As Variant
Dim lastR, i As Integer
lastR = Sheets(45).Cells(Rows.count, 2).End(xlUp).row
For i = 1 To lastR
PASS = Sheets(45).Cells(i, 2).value

Next i


    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(53).Range("a2:ci7").ClearContents
    Sheets(53).Range("a2:ci7").Merge
    Range("a2:ci7").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(53).Range("a2:ci7").Font.Name = "Arial"
   Sheets(53).Range("a2").Font.Bold = True
       
     Sheets(53).Range("a2:c17").ClearContents
    Sheets(53).Range("a2:ci7").HorizontalAlignment = xlCenter
     Sheets(53).Range("a2:ci7").VerticalAlignment = xlCenter
       
    Sheets(53).Range("a2:ci7").value = UserForm36.ComboBox1.value & " " & UserForm36.ComboBox2.value & " " & "·‘Â—" & " " & UserForm36.ComboBox3.value _
    & " " & UserForm36.ComboBox4.value
    'Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(53)
   
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Application.Visible = True
             Application.ActiveWorkbook.SaveAs fileName:="D:\export\" & UserForm36.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm36.ComboBox3.value & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
             Application.Workbooks(2).Close True
            
        End With
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False

End Sub
