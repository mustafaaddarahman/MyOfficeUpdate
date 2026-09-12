Attribute VB_Name = "export_pdf_form_53"
Sub export_pdf_form_53_1()
On Error Resume Next
'Sheets(53).Activate
Dim lstr As Integer
lstr = Sheets(53).Cells(Rows.count, "a").End(xlUp).row
    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(53).Range("a2:ci7").ClearContents
    Sheets(53).Range("a2:ci7").Merge
    Range("a2:ci7").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(53).Range("a2:c17").Font.Name = "Arial"
    Range("a2").Font.Bold = True
       
    
     Range("a2:ci7").HorizontalAlignment = xlCenter
      Range("a2:ci7").VerticalAlignment = xlCenter
       
    Sheets(53).Range("a2:ci7").value = UserForm36.ComboBox1.value & " " & UserForm36.ComboBox2.value & " " & "·‘Â—" & " " & UserForm36.ComboBox3.value _
    & " " & UserForm36.ComboBox4.value
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(53)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
        
           .Copy
            Dim fName As Variant
            fName = "D:\export\" & UserForm36.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm36.ComboBox3.value
                             
           
           
          Application.Range("a2:cd" & lstr).ExportAsFixedFormat Type:=xlTypePDF, _
             fileName:=fName & ".pdf"
                    
            
        End With
  Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
     Application.Workbooks(2).Close True


End Sub
