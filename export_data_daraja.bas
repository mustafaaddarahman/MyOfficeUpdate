Attribute VB_Name = "export_data_daraja"
Sub export_data_daraja1()
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
     Sheets(18).Range("a2:m5").ClearContents
    Sheets(18).Range("a2:m5").Merge
    Range("a2:m5").WrapText = True
     
     
    Sheets(18).Range("a2:m5").Font.Name = "Arial"
    Range("a2").Font.Bold = True
       
    
     Range("a2:m5").HorizontalAlignment = xlCenter
      Range("a2:m5").VerticalAlignment = xlCenter
       
    Sheets(18).Range("a2:m5").value = UserForm36.ComboBox1.value & " " & UserForm36.ComboBox2.value & " " & "·‘Â—" & " " & UserForm36.ComboBox3.value _
    & " " & UserForm36.ComboBox4.value
    Range("a2").Font.Size = 28
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(18)
    
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

emad:







End Sub
