Attribute VB_Name = "expory_year_month_44"
Sub expory_year_month_44_1()


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
     Sheets(44).Range("a2:bq5").ClearContents
    Sheets(44).Range("a2:m5").Merge
    Range("a2:m5").WrapText = True
     Range("a2").Font.Size = 18
     
    Sheets(44).Range("a2:m5").Font.Name = "Arial"
   Sheets(44).Range("a2").Font.Bold = True
       
    
     Sheets(44).Range("a2:m5").HorizontalAlignment = xlCenter
     Sheets(44).Range("a2:m5").VerticalAlignment = xlCenter
       
    Sheets(44).Range("a2:m5").value = UserForm40.ComboBox1.value & " " & UserForm40.ComboBox2.value & " " & "·‘Â—" & " " & UserForm40.ComboBox3.value _
    & " " & UserForm38.ComboBox4.value
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(44)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Application.Visible = True
             Application.ActiveWorkbook.SaveAs fileName:="D:\export\" & UserForm40.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm40.ComboBox3.value & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled
              Application.Workbooks(2).Close True
           
            
        End With
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
emad:

End Sub




