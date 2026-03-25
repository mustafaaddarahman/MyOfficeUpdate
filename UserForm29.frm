VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm29 
   Caption         =   "UserForm29"
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6816
   OleObjectBlob   =   "UserForm29.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub ComboBox1_AfterUpdate()

End Sub

Private Sub ComboBox1_Change()
On Error Resume Next

If ComboBox1.value = "«·„Êﬁ› «·Õ«·Ì" Then

Label25.Visible = False
'=======================================================
Dim lastR As Integer
lastR = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For Y = 9 To lastR
    
    If Sheets(1).Cells(Y, 2) = TextBox68.Text And Sheets(1).Cells(Y, 5) = TextBox60.Text Then
   
    
     ComboBox1.Text = Sheets(1).Cells(Y, 58)
     TextBox67.Text = Sheets(1).Cells(Y, 49)
    


    End If
    Next Y

Dim num As Long

num = val(TextBox67.value / 30) * val(TextBox3.value)
TextBox66.value = num

  TextBox4.value = val(TextBox65) + val(TextBox66)
ElseIf ComboBox1.value = "„Õ«· ⁄·Ï «· ﬁ«⁄œ" Then
Label25.Visible = True
TextBox4.value = val(TextBox65)
End If
Dim lastr1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For y1 = 9 To lastr1
    
    If Sheets(1).Cells(y1, 2) = TextBox68.Text And Sheets(1).Cells(y1, 5) = TextBox60.Text Then
        TextBox67.Text = Sheets(1).Cells(y1, 49)
    









End If
Next


End Sub





Private Sub CommandButton1_Click()
Dim num As Long
num = val(TextBox66.value) + val(TextBox65.value)
TextBox4.value = num
If TextBox5.Text = "" Or TextBox64.Text = "" Then
MsgBox "ÌÃ» «œŒ«· Ï «—ÌŒ «·–Ì Ì” ÕﬁÂ ·Â–« «·‘Â— ·Ì ”‰Ï ··‰Ÿ«„ «Õ ”«»Â Ÿ„‰« Ê«Õ ”«» ﬁÌœÂ ÷„‰ Â–« «· «—ÌŒ", vbCritical, "—”«·…  ‰»ÌÂ"
Exit Sub
End If



'===========================

Dim iRow As Long, i As Long

    Sheets(37).Activate

    

      
      
       
'============================================================================

'=======================================================
 lastR = Sheets(37).Cells(Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(37).Cells(Y, 1) = TextBox68.Text And Sheets(37).Cells(Y, 2) = TextBox60.Text Then
   Cells(Y, 1).Select
ActiveCell.EntireRow.Delete
    Exit For
    End If
    Next Y
    
    
    
    
  
  
   
   
   
   
    lastr1 = Sheets(37).Cells(Rows.count, 1).End(xlUp).row
    For y1 = 2 To lastr1
      If Sheets(37).Cells(y1, 1) <> TextBox68.Text And Sheets(37).Cells(y1, 2) <> TextBox60.Text Then
    
    Exit For
    End If
     Next
    iRow = Range("A" & Rows.count).End(xlUp).row
    Range("A" & iRow + 1).Offset(0, 0).value = TextBox68.value
    Range("A" & iRow + 1).Offset(0, 1).value = TextBox60.value
    Range("A" & iRow + 1).Offset(0, 2).value = TextBox1.value
    Range("A" & iRow + 1).Offset(0, 3).value = TextBox61.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox2.value
    Range("A" & iRow + 1).Offset(0, 5).value = TextBox65.value
    Range("A" & iRow + 1).Offset(0, 6).value = ComboBox1.value
    Range("A" & iRow + 1).Offset(0, 7).value = TextBox67.value
    Range("A" & iRow + 1).Offset(0, 8).value = TextBox3.value
    Range("A" & iRow + 1).Offset(0, 9).value = TextBox66.value
    Range("A" & iRow + 1).Offset(0, 10).value = TextBox4.value
    
    Range("A" & iRow + 1).Offset(0, 11).value = Format(TextBox5.Text, "yyyy/mm/dd")
    Range("A" & iRow + 1).Offset(0, 12).value = Format(TextBox64.Text, "yyyy/mm/dd")
        UserForm1.TextBox49.Text = TextBox4.value
    UserForm1.TextBox58.Text = "—« » Ã“∆Ì"
 UserForm1.ComboBox3.value = "—« » Ã“∆Ì"
 

 lastr2 = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For y2 = 9 To lastr2
      If Sheets(1).Cells(y2, 2) = TextBox68.Text And Sheets(1).Cells(y2, 5) = TextBox60.Text Then
    Sheets(1).Cells(y2, 49) = TextBox4.Text
    Sheets(1).Cells(y2, 58) = UserForm1.TextBox58.Text
    End If
     Next

'=====================================??? ???????? ??? ??? 46
 Sheets(46).Activate
Dim iRow3 As Long
Dim i3 As Long
Dim lastr3 As Long
Dim y3 As Long

       

 lastr3 = Sheets(46).Cells(Rows.count, 1).End(xlUp).row
    For y3 = 2 To lastr3
    
    If Sheets(46).Cells(y3, 1) = TextBox68.Text And Sheets(46).Cells(y3, 2) = TextBox60.Text Then
   Cells(y3, 1).Select
ActiveCell.EntireRow.Delete
    Exit For
    End If
    Next y3
    
    lastr3 = Sheets(46).Cells(Rows.count, 1).End(xlUp).row
    For y3 = 2 To lastr3
      If Sheets(46).Cells(y3, 1) <> TextBox68.Text And Sheets(46).Cells(y3, 2) <> TextBox60.Text Then
    
    Exit For
    End If
     Next
    iRow = Range("A" & Rows.count).End(xlUp).row
    Range("A" & iRow + 1).Offset(0, 0).value = TextBox68.value
    Range("A" & iRow + 1).Offset(0, 1).value = TextBox60.value
    Range("A" & iRow + 1).Offset(0, 2).value = TextBox1.value
    Range("A" & iRow + 1).Offset(0, 3).value = TextBox61.value
    Range("A" & iRow + 1).Offset(0, 4).value = TextBox2.value
    Range("A" & iRow + 1).Offset(0, 5).value = TextBox65.value
    Range("A" & iRow + 1).Offset(0, 6).value = ComboBox1.value
    
Dim month_date As Date
Dim name_month As Variant
month_date = Month(Sheets(46).Cells(1, "m"))
name_month = monthName(month_date)
 Range("A" & iRow + 1).Offset(0, 7).value = name_month
   












End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub



Private Sub TextBox2_AfterUpdate()
TextBox65.value = val(TextBox61.value) / 30 * val(TextBox2.value)
 TextBox3.value = 30 - val(TextBox2.value)

End Sub

Private Sub TextBox2_Change()
'On Error Resume Next
TextBox65.value = val(TextBox61.value) / 30 * val(TextBox2.value)
 TextBox3.value = 30 - val(TextBox2.value)

End Sub

Private Sub TextBox3_AfterUpdate()
On Error Resume Next


TextBox66.value = val(TextBox67.value) / 30 * val(TextBox3.value)
End Sub

Private Sub TextBox3_Change()
On Error Resume Next

TextBox66.value = val(TextBox67.value) / 30 * val(TextBox3.value)

End Sub

Private Sub TextBox60_Change()

End Sub

Private Sub TextBox62_AfterUpdate()
If TextBox62.Text = "" Then
TextBox60.Text = ""
End If
End Sub

Private Sub TextBox62_Change()

Dim lastR, r As Integer
lastR = Sheets(38).Cells(Rows.count, 1).End(xlUp).row
For r = 2 To lastR
If Sheets(38).Cells(r, 2).value Like "*" & TextBox62.Text & "*" Then
TextBox68.Text = Sheets(38).Cells(r, 1)
 TextBox60.Text = Sheets(38).Cells(r, 2)
 TextBox1.Text = Sheets(38).Cells(r, 3)
 TextBox61.Text = Sheets(38).Cells(r, 4)
  TextBox69.Text = Sheets(38).Cells(r, 5)
 
 
 
 
Exit For
End If
Next
End Sub

Private Sub TextBox66_AfterUpdate()
 TextBox4.value = Format(val(TextBox65) + val(TextBox66), "00000000")
End Sub

Private Sub TextBox66_Change()

 TextBox4.value = val(TextBox65) + val(TextBox66)

End Sub

Private Sub TextBox67_Change()

End Sub

Private Sub TextBox68_Change()

End Sub

Private Sub UserForm_Activate()
On Error Resume Next
Label25.Visible = False
ComboBox1.List = Array("«·„Êﬁ› «·Õ«·Ì", "„Õ«· ⁄·Ï «· ﬁ«⁄œ")
TextBox62.Text = UserForm1.TextBox5.Text





End Sub


