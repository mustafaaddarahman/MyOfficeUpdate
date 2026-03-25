Attribute VB_Name = "delete_change_month"
Sub delete_change_month1()

Dim lastR As Integer
lastR = Sheets(35).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(35).Cells(Y, 2).value = UserForm1.TextBox2.Text And Sheets(35).Cells(Y, 5).value = UserForm1.TextBox5.Text Then
    Cells(Y, 1).Select
ActiveCell.EntireRow.Delete Shift:=xlUp
   MsgBox " „ Õ–› «·ﬁÌœ »‰Ã«Õ", vbCritical, "—”«·…  ‰»ÌÂ"
   
    Exit For
    End If
    Next Y






End Sub
