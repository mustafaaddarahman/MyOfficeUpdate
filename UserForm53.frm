VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm53 
   Caption         =   "UserForm53"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4896
   OleObjectBlob   =   "UserForm53.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm53"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub ListBox1_Click()
If ListBox1.ListIndex <> -1 Then
        
        ' ????? ?????? ???????? ??? ??????? ???? ?? ??????? ??????
        UserForm1.TextBox72.value = ListBox1.value
         
        ' ????? ??????? ?????? (??????? ??? ?? ??? ??????)
       Unload Me
      
      If UserForm1.ListBox1.ListCount > 0 Then
        UserForm1.ListBox1.ListIndex = 0
    End If
      
    End If
     UserForm1.ComboBox1.DropDown
    UserForm1.ComboBox1.ListIndex = 4
    
End Sub

Private Sub UserForm_Activate()
Dim lastr1, y1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For y1 = 9 To lastr1
        If Sheets(1).Cells(y1, "x") = "ÍÓÈ ÇáÊÞÇÊØÚ ÇáÌÏæá ÇáÖÑíÈí Çä íßæä ÇáÚÏÏ ÕÝÑ" Then
  
  ListBox1.AddItem (Sheets(1).Cells(y1, "e"))
     '  ListBox1.List(k, 0) = Sheets(1).Cells(y1, "e").Value
       
        
   
   
   
  ' MsgBox "åäÇß ÎØÃ Ýí ÇÍÊÓÇÈ ãäÙæãÉ ÇáÇÓÊÞØÇÚ ÇáÖÑíÈí ÍÇæá ÇÕáÇÍÉ ÞÏ íßæä Ýí ÇáÞíÏ " & Sheets(1).Cells(y, "e"), vbCritical, "ÑÓÇáÉ ÊäÈíå"

End If
Next

Dim lastr2, y2 As Integer
lastr2 = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
    For y2 = 9 To lastr2
If Sheets(1).Cells(y2, "d") = "ÇäËì" And Sheets(1).Cells(y2, "w") = "ÇáãæÙÝ ÇáãÊÒæÌ æ ÒæÌÊå ÑÈÉ ÈíÊ" Then
 
  ListBox1.AddItem (Sheets(1).Cells(y2, "e"))
     '  ListBox1.List(k, 0) = Sheets(1).Cells(y1, "e").Value
        
   End If
Next

Dim lastr3, y3 As Integer
lastr3 = Sheets(1).Cells(Rows.count, 1).End(xlUp).row
For y3 = 9 To lastr3
If Sheets(1).Cells(y3, "w") = "ÇáÇÚÒÈ" And Sheets(1).Cells(y3, "t") > 0 Then
 ListBox1.AddItem (Sheets(1).Cells(y3, "e"))
     '  ListBox1.List(k, 0) = Sheets(1).Cells(y1, "e").Value
  End If
Next

End Sub

Private Sub UserForm_Click()

End Sub
