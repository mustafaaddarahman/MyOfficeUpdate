Attribute VB_Name = "export_pdf_draja"
Sub export_pdf_draja_1()
On Error Resume Next
On Error GoTo emad
Dim lstr As Integer
lstr = Sheets(18).Cells(Rows.count, "a").End(xlUp).row
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
            Dim fName As Variant
            fName = "D:\export\" & UserForm36.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm36.ComboBox3.value
           
             Application.Range("a2:m" & lstr).ExportAsFixedFormat Type:=xlTypePDF, _
             fileName:=fName & ".pdf"
            
           
            
        End With
     Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
   Application.Workbooks(2).Close True
emad:








End Sub
