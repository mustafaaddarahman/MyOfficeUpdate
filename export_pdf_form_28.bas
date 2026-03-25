Attribute VB_Name = "export_pdf_form_28"
Sub export_pdf_form_28_1()

On Error Resume Next
On Error GoTo emad
Dim lstr As Integer
lstr = Sheets(36).Cells(Rows.count, "a").End(xlUp).row
    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(36).Range("a2:cd5").ClearContents
    Sheets(36).Range("a2:cd5").Merge
    Range("a2:cd5").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(36).Range("a2:cd5").Font.Name = "Arial"
    Range("a2").Font.Bold = True
       
    
     Range("a2:cd5").HorizontalAlignment = xlCenter
      Range("a2:cd5").VerticalAlignment = xlCenter
       
    Sheets(36).Range("a2:cd5").value = UserForm37.ComboBox1.value & " " & UserForm37.ComboBox2.value & " " & "·‘Â—" & " " & UserForm37.ComboBox3.value _
    & " " & UserForm37.ComboBox4.value
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(36)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
        
           .Copy
            Dim fName As Variant
            fName = "D:\export\" & UserForm37.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm37.ComboBox3.value
                             
           
           
          Application.Range("a2:cd" & lstr).ExportAsFixedFormat Type:=xlTypePDF, _
             fileName:=fName & ".pdf"
                    
            
        End With
  Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
     Application.Workbooks(2).Close True
emad:


End Sub

