Attribute VB_Name = "export_pdf_form_32"
Sub export_pdf_form_32_1()
On Error Resume Next
On Error GoTo emad
Dim lstr As Integer
lstr = Sheets(39).Cells(Rows.count, "a").End(xlUp).row
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
            Dim fName As Variant
            fName = "D:\export\" & UserForm38.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm38.ComboBox3.value
           
          Application.Range("a2:bq" & lstr).ExportAsFixedFormat Type:=xlTypePDF, _
          fileName:=fName & ".pdf"
                    
            
        End With
  Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
     Application.Workbooks(2).Close True
emad:
 
End Sub
