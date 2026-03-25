Attribute VB_Name = "export_pdf_form_17"
Sub export_pdf_form_17_1()

On Error GoTo emad
 
Dim lstr As Integer
lstr = Sheets(28).Cells(Rows.count, "a").End(xlUp).row
    Dim filePath As String
    Dim ws As Worksheet
    filePath = Application.ActiveWorkbook.path
     Sheets(28).Range("a2:bz5").ClearContents
    Sheets(28).Range("a2:bz5").Merge
    Range("a2:az5").WrapText = True
     Range("a2").Font.Size = 28
     
    Sheets(28).Range("a2:br5").Font.Name = "Arial"
    Range("a2").Font.Bold = True
       
    
     Range("a2:az5").HorizontalAlignment = xlCenter
      Range("a2:bz5").VerticalAlignment = xlCenter
       
    Sheets(28).Range("a2:bz5").value = UserForm36.ComboBox1.value & " " & UserForm36.ComboBox2.value & " " & "·‘Â—" & " " & UserForm36.ComboBox3.value _
    & " " & UserForm36.ComboBox4.value
    Cells(2, "a").EntireColumn.AutoFit
    
    Set ws = Sheets(28)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
        With ws
            .Copy
            Dim fName As Variant
            fName = "D:\export\" & UserForm36.ComboBox1.value & " " & " ·Â–« «·‘Â—" & UserForm36.ComboBox3.value
           
             Application.Range("a2:bz" & lstr).ExportAsFixedFormat Type:=xlTypePDF, _
             fileName:=fName & ".pdf"
            
           
            
        End With
     Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Visible = False
   Application.Workbooks(2).Close True
emad:
End Sub

