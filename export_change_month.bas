Attribute VB_Name = "export_change_month"
Sub export_change_month1()
On Error Resume Next
On Error GoTo emad
Dim PASS As Variant
Dim lastR, i As Integer
lastR = Sheets(45).Cells(Rows.count, 2).End(xlUp).row
For i = 1 To lastR
PASS = Sheets(45).Cells(i, 2).value

Next i


    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(36).Range("a2:ce5").ClearContents
    Sheets(36).Range("a2:ce5").Merge
    Range("a2:ce5").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(36).Range("a2:ce5").Font.Name = "Arial"
   Sheets(36).Range("a2").Font.Bold = True
       
    
    Sheets(36).Range("a2:ce5").HorizontalAlignment = xlCenter
     Sheets(36).Range("a2:ce5").VerticalAlignment = xlCenter
       
    Sheets(36).Range("a2:ce5").value = UserForm37.ComboBox1.value & " " & UserForm37.ComboBox2.value & " " & "·‘Â—" & " " & UserForm37.ComboBox3.value _
    & " " & UserForm37.ComboBox4.value
    'Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(36)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Application.Visible = True
             Application.ActiveWorkbook.SaveAs fileName:="D:\export\" & UserForm37.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm37.ComboBox3.value & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
            
        End With
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
emad:






End Sub
