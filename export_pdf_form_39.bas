Attribute VB_Name = "export_pdf_form_39"
Sub export_pdf_form_39_1()
On Error Resume Next
'On Error GoTo emad
Sheets(44).Activate
Dim lstr As Integer

lstr = Sheets(44).Cells(Rows.count, "a").End(xlUp).row
    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(44).Range("a2:m5").ClearContents
    Sheets(44).Range("a2:m5").Merge
    Sheets(44).Range("a2:m5").WrapText = True
      Sheets(44).Range("a2").Font.Size = 18
     
    Sheets(44).Range("a2:m5").Font.Name = "Arial"
    Sheets(44).Range("a2").Font.Bold = True
       
    
     Sheets(44).Range("a2:m5").HorizontalAlignment = xlCenter
      Sheets(44).Range("a2:m5").VerticalAlignment = xlCenter
       
    Sheets(44).Range("a2:m5").value = UserForm40.ComboBox1.value & " " & UserForm40.ComboBox2.value & " " & "·‘Â—" & " " & UserForm40.ComboBox3.value _
    & " " & UserForm40.ComboBox4.value
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(44)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
        
        .Copy
            Dim fName As Variant
            fName = "D:\export\" & UserForm40.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm40.ComboBox3.value
           
          Application.Range("a2:m" & lstr).ExportAsFixedFormat Type:=xlTypePDF, _
          fileName:=fName & ".pdf"
                    
            
        End With
  Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
     Application.Workbooks(2).Close True
'emad:
 
End Sub


