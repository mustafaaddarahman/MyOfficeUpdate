Attribute VB_Name = "expory_delete_leave"
Sub export_delete_leave1()
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
     Sheets(39).Range("a2:bq5").ClearContents
    Sheets(39).Range("a2:bq5").Merge
    Range("a2:bq5").WrapText = True
     Range("a2").Font.Size = 18
     
    Sheets(39).Range("a2:bq5").Font.Name = "Arial"
    Range("a2").Font.Bold = True
       
    
     Range("a2:bq5").HorizontalAlignment = xlCenter
      Range("a2:bq5").VerticalAlignment = xlCenter
       
    Sheets(39).Range("a2:bq5").value = UserForm38.ComboBox1.value & " " & UserForm38.ComboBox2.value & " " & "·‘Â—" & " " & UserForm38.ComboBox3.value _
    & " " & UserForm38.ComboBox4.value
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(39)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Application.Visible = True
             Application.ActiveWorkbook.SaveAs fileName:="D:\export\" & UserForm38.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm38.ComboBox3.value & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
              Application.Workbooks(2).Close True
           
            
        End With
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
emad:
End Sub
