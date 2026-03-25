VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "ãÌÇãíÚ ÍÓÈ ÇáİÆÇÊ"
   ClientHeight    =   12285
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   18324
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const WS_THICKFRAME As Long = &H40000
Const MIN_BOX As Long = &H20000
Const MAX_BOX As Long = &H10000
Private Declare PtrSafe Function DrawMenuBar Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As Long

Public Sub AddToForm(ByVal Box_Type As Long)
Dim BisMask As Long
Dim Window_Handle As Long
Dim WindowStyle As Long
Dim Ret As Long
If Box_Type = MIN_BOX Or Box_Type = MAX_BOX Then
    Window_Handle = GetForegroundWindow()
    WindowStyle = GetWindowLong(Window_Handle, GWL_STYLE)
    BisMask = WindowStyle Or Box_Type
    Ret = SetWindowLong(Window_Handle, GWL_STYLE, BisMask)
    Ret = DrawMenuBar(Window_Handle)
End If
End Sub
Private Sub ComboBox1_Change()
With Me.ListView1
'Setting  ÇÚÏÇÏÇÊ ÇáãÙåÑ
.FullRowSelect = True 'áÊÙáíá ßÇãá ÇáÕİ
.Gridlines = True
.View = lvwReport  'ÚÑÖ ÎØæØ ÇáÔÈßÉ
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' áæä ÇáÎØ
.Font.Size = 11 'ÍÌã ÇáÎØ
.Font.Bold = False  'ÛÇãŞ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
Next frw1
'=============================================
 Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next

    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
  
   

End With
End With
End With
ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
Me.Label66.Caption = "ÇáÚÏÏ Çáßáí áÚäæÇä" & Me.ComboBox1.value
Me.TextBox65.Text = ListView1.ListItems.count
End Sub

Private Sub ComboBox3_Change()
Me.Label67.Caption = "ÇáÚÏÏ Çáßáí áÚäæÇä" & Me.ComboBox3.value
End Sub

Private Sub CommandButton1_Click()
Unload Me
UserForm9.Show
End Sub

Private Sub CommandButton2_Click()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "n") >= TextBox1.Text And Sheets(1).Cells(frw2, "p") >= TextBox2.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox11.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("n9:n571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("n9:n571"), ">=" & Me.TextBox1.Text)
TextBox12.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("p9:p571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("p9:p571"), "=" & Me.TextBox2.Text)
TextBox13.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("r9:r571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("r9:r571"), "=" & Me.TextBox3.Text)
TextBox14.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("s9:s571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("s9:s571"), "=" & Me.TextBox4.Text)
TextBox15.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("t9:t571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("t9:t571"), "=" & Me.TextBox5.Text)
TextBox16.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("u9:u571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("u9:u571"), "=" & Me.TextBox6.Text)
TextBox17.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("j9:j571"), "=" & Me.TextBox8.Text)
TextBox18.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aa9:aa571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("aa9:aa571"), "=" & Me.TextBox9.Text)
TextBox19.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ac9:ac571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("ac9:ac571"), "=" & Me.TextBox10.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count


End Sub

Private Sub CommandButton3_Click()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================








'====================================================

'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
Me.TextBox66.value = ListView2.ListItems.count
End Sub

Private Sub CommandButton4_Click()
With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")

item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(46) = Sheets(1).Cells(frw1, "au")
item1.SubItems(47) = Sheets(1).Cells(frw1, "av")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
Next frw1
Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next

    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
       ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    
    
    
   
    
    
    
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
   
    
    
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
   
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
  
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
   
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
   
    
   
   
 
    
   
   
  
    
'=================================================
End With
End With

ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
End Sub

Private Sub OptionButton1_Click()
If OptionButton1.value = True Then
OptionButton2.value = False
Call arb
End If
End Sub

Private Sub OptionButton2_Click()
On Error Resume Next
If OptionButton2.value = True Then
OptionButton1.value = False
If Not Intersect(Target, UserForm1.ActiveControl) Is Nothing Then
Call englsh
End If
End If
End Sub

Private Sub TextBo29_Change()
On Error Resume Next
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "at") >= val(TextBox29.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox73.value = Application.WorksheetFunction.SumIfs(Sheets(1).Range("at9:at571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("at9:at571"), ">=" & TextBox29.Text)






'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox1_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "n") >= TextBox1.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox11.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("n9:n571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("n9:n571"), ">=" & Me.TextBox1.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox10_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer


 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "y") >= TextBox10.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox67.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("y9:y571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("y9:y571"), "=" & Me.TextBox10.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox11_Change()

End Sub

Private Sub TextBox2_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "p") >= TextBox2.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================

TextBox12.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("p9:p571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("p9:p571"), ">=" & Me.TextBox2.Text)






'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox20_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "af") >= val(TextBox20.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox32Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("af9:af571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("af9:af571"), ">=" & Me.TextBox20.Text)







'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox21_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "ag") >= val(TextBox21.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox31.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ag9:ag571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("ag9:ag571"), ">=" & Me.TextBox21.Text)







'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox22_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "ah") >= val(TextBox22.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox33.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ah9:ah571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("ah9:ah571"), ">=" & Me.TextBox22.Text)







'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox23_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "ai") >= val(TextBox23.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox34Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ai9:ai571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("ai9:ai571"), ">=" & Me.TextBox23.Text)






'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox24_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "aj") >= val(TextBox24.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox35.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aj9:aj571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aj9:aj571"), ">=" & Me.TextBox24.Text)







'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox25_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "al") >= val(TextBox25.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox36.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("al9:al571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("al9:al571"), ">=" & Me.TextBox25.Text)





'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox26_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "an") >= val(TextBox26.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox64.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("an9:an571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("an9:an571"), ">=" & Me.TextBox26.Text)







'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox27_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "ap") >= val(TextBox27.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox71.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ap9:ap571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("ap9:ap571"), ">=" & Me.TextBox27.Text)







'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox28_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "ar") >= val(TextBox28.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox72.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ar9:ar571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("ar9:ar571"), ">=" & Me.TextBox28.Text)





'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox3_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "r") >= TextBox3.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================

TextBox13.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("r9:r571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("r9:r571"), ">=" & Me.TextBox3.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox4_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "s") >= TextBox4.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox14.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("s9:s571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("s9:s571"), ">=" & Me.TextBox4.Text)





'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox5_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "t") >= TextBox5.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox15.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("t9:t571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("t9:t571"), ">=" & Me.TextBox5.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox6_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "u") >= TextBox6.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox16.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("u9:u571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("u9:u571"), ">=" & Me.TextBox6.Text)






'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox67_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "y") >= TextBox10.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox11.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("n9:n571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("n9:n571"), ">=" & Me.TextBox1.Text)
TextBox12.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("p9:p571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("p9:p571"), "=" & Me.TextBox2.Text)
TextBox13.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("r9:r571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("r9:r571"), "=" & Me.TextBox3.Text)
TextBox14.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("s9:s571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("s9:s571"), "=" & Me.TextBox4.Text)
TextBox15.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("t9:t571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("t9:t571"), "=" & Me.TextBox5.Text)
TextBox16.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("u9:u571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("u9:u571"), "=" & Me.TextBox6.Text)
TextBox17.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("j9:j571"), "=" & Me.TextBox8.Text)
TextBox18.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aa9:aa571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("aa9:aa571"), "=" & Me.TextBox9.Text)
TextBox19.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ac9:ac571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("ac9:ac571"), "=" & Me.TextBox10.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox68_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "x") >= val(TextBox68.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox30.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("x9:x571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("x9:x571"), ">=" & Me.TextBox68.Text)





'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox7_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "j") >= TextBox7.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox17.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("j9:j571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("j9:j571"), ">=" & Me.TextBox7.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox70_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer

 With Me.ListView2
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f") = Me.ComboBox3.value And Sheets(1).Cells(frw2, "au") >= val(TextBox70.Text) Then
Set item2 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")


'===========
item2.SubItems(23) = Sheets(1).Cells(frw2, "x")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(31) = Sheets(1).Cells(frw2, "af")
item2.SubItems(32) = Sheets(1).Cells(frw2, "ag")
item2.SubItems(33) = Sheets(1).Cells(frw2, "ah")
item2.SubItems(34) = Sheets(1).Cells(frw2, "ai")
item2.SubItems(35) = Sheets(1).Cells(frw2, "aj")
item2.SubItems(37) = Sheets(1).Cells(frw2, "al")
item2.SubItems(39) = Sheets(1).Cells(frw2, "an")
item2.SubItems(41) = Sheets(1).Cells(frw2, "ap")
item2.SubItems(43) = Sheets(1).Cells(frw2, "ar")
item2.SubItems(45) = Sheets(1).Cells(frw2, "at")
item2.SubItems(47) = Sheets(1).Cells(frw2, "au")
item2.SubItems(48) = Sheets(1).Cells(frw2, "av")
item2.SubItems(49) = Sheets(1).Cells(frw2, "aw")
'============================================

TextBox74.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("au9:au571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("au9:au571"), ">=" & Me.TextBox70.Text)






'====================================================
Me.TextBox66.value = ListView2.ListItems.count
'TextBox66.Text = Application.WorksheetFunction.CountIfs(Sheets(1).Range("f9:f571"), "*" & Me.ComboBox3.Text & "*", Sheets(1).Range("aw9:aw571"), ">=" & Sheets(1).Cells(frw2, "aw"))

End If
Next frw2
End With
End Sub

Private Sub TextBox8_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "aa") >= TextBox8.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox18.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("aa9:aa571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("aa9:aa571"), ">=" & Me.TextBox8.Text)





'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub TextBox9_Change()
Dim item2 As ListItem
Dim last2, frw2 As Integer
Me.TextBox10.Text = ""
Me.TextBox11.Text = ""
 With Me.ListView1
.ListItems.Clear
Sheets(1).Activate
last2 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw2 = 9 To last2
If Sheets(1).Cells(frw2, "f").value = Me.ComboBox1.value And Sheets(1).Cells(frw2, "ac") >= TextBox9.Text Then
Set item2 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw2, "A"))
'==================================================
item2.SubItems(1) = Sheets(1).Cells(frw2, "b")
item2.SubItems(4) = Sheets(1).Cells(frw2, "e")
item2.SubItems(5) = Sheets(1).Cells(frw2, "f")
item2.SubItems(6) = Sheets(1).Cells(frw2, "g")
item2.SubItems(9) = Sheets(1).Cells(frw2, "j")
item2.SubItems(13) = Sheets(1).Cells(frw2, "n")
item2.SubItems(15) = Sheets(1).Cells(frw2, "p")
item2.SubItems(17) = Sheets(1).Cells(frw2, "r")
item2.SubItems(18) = Sheets(1).Cells(frw2, "s")
item2.SubItems(19) = Sheets(1).Cells(frw2, "t")
item2.SubItems(20) = Sheets(1).Cells(frw2, "u")
item2.SubItems(24) = Sheets(1).Cells(frw2, "y")
item2.SubItems(26) = Sheets(1).Cells(frw2, "aa")
item2.SubItems(28) = Sheets(1).Cells(frw2, "ac")
item2.SubItems(29) = Sheets(1).Cells(frw2, "ad")
item2.SubItems(30) = Sheets(1).Cells(frw2, "ae")
item2.SubItems(48) = Sheets(1).Cells(frw2, "aw")
'===========================================
TextBox19.Text = Application.WorksheetFunction.SumIfs(Sheets(1).Range("ac9:ac571"), Sheets(1).Range("f9:f571"), "*" & Me.ComboBox1.Text & "*", Sheets(1).Range("ac9:ac571"), ">=" & Me.TextBox9.Text)







'===================================================

End If
Next frw2
End With
Me.TextBox65.Text = ListView1.ListItems.count

End Sub

Private Sub UserForm_Activate()
On Error Resume Next

'==============================
Call AddToForm(MIN_BOX)
Call AddToForm(MAX_BOX)

Sheets(5).Activate
ComboBox1.List = Sheets(5).Range("a1:a373").value
ComboBox3.List = Sheets(5).Range("a1:a373").value

With ListView1
.Gridlines = True
.View = lvwReport
.FullRowSelect = True
.ColumnHeaders.Add , , "ÇáÑŞã", 40
.ColumnHeaders.Add , , "ÇáÑŞã ÇáæÙíİí", 120
.ColumnHeaders.Add , , "ÑŞã ÇáãÇÓÊÑ", 0
.ColumnHeaders.Add , , "ÇáÌäÓ", 0
.ColumnHeaders.Add , , "ÇÓã ÇáãæÙİ", 140
.ColumnHeaders.Add , , "ÇáÚäæÇä ÇáæÙíİí", 100
.ColumnHeaders.Add , , " ÇáŞÓã", 100
.ColumnHeaders.Add , , " ÇáÏÑÌÉ", 0
.ColumnHeaders.Add , , "ÇáãÑÍáÉ", 0
.ColumnHeaders.Add , , "ÇáÑÇÊÈ ÇáÇÓãí ", 140
.ColumnHeaders.Add , , "ÇíÇã ÇáÛíÇÈ", 0
.ColumnHeaders.Add , , " ÇíÇã ÇáÏæÇã", 0
.ColumnHeaders.Add , , " äÓÈÉ ÇáÔåÇÏÉ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÇáÔåÇÏÉ", 120
.ColumnHeaders.Add , , "äÓÈÉ ÇáåäÏÓíå ", 0
.ColumnHeaders.Add , , " ãÎÕÕÇÊ ÇáåäÏÓíå", 120
.ColumnHeaders.Add , , "äÓÈÉ ÇáãäÕÈ ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÇáãäÕÈ ", 120
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ãæŞÚ ÌÛÑÇİí", 120
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÇØİÇá  ", 140
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÒæÌíå", 120
.ColumnHeaders.Add , , " ÚÏÏ ÇáÇØİÇá", 0
.ColumnHeaders.Add , , "ÇáÍÇáå ÇáÇÌÊãÇÚíå ", 0
.ColumnHeaders.Add , , "ÇáÇÓÊŞØÇÚ ÇáÖÑíÈí", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÎØæÑÉ ãÈáÛ ", 140
.ColumnHeaders.Add , , " äÓÈÉãÎÕÕÇÊ ãåäíå ", 0
.ColumnHeaders.Add , , "  ãÎÕÕÇÊ ãåäíå ", 120
.ColumnHeaders.Add , , "äÓÈÉ ÇáÎØæÑå", 0
.ColumnHeaders.Add , , " ãÎÕÕÇÊ ÇáÎØæÑÉ ", 140
.ColumnHeaders.Add , , "ãÌãæÚ ÇáãÎÕÕÇÊ", 120
.ColumnHeaders.Add , , "ãÌãæÚ ÇáÇÓÊÍŞÇŞÇÊ", 120
.ColumnHeaders.Add , , " ÊæŞİÇÊ ÊŞÇÚÏíå 10", 0
.ColumnHeaders.Add , , "ÇÊæŞİÇÊ ÊŞÇÚÏíå 15%", 0
.ColumnHeaders.Add , , "ÇáÊÇãíä ÇáÕÍí ", 0
.ColumnHeaders.Add , , " ÇÌæÑ ÇáäŞá", 0
.ColumnHeaders.Add , , " ÑÓã ÇáØÇÈÚ", 0
.ColumnHeaders.Add , , " æÒÇÑÉ ÇáãÇáíå", 0
.ColumnHeaders.Add , , "ÇáÇÓÊŞØÇÚ1", 0
.ColumnHeaders.Add , , " ÌåÉ ÇáÇÓÊŞØÇÚ1  ", 0
.ColumnHeaders.Add , , "ÇÓÊŞØÇÚ ãÊİÑŞ", 0
.ColumnHeaders.Add , , " ÌåÉ ÇáÇÓÊŞØÇÚ ", 0
.ColumnHeaders.Add , , " äÇÏí ÇáåíÃÉ ÇáÕäÇÚÇÊ ÇáÍÑÈí", 0
.ColumnHeaders.Add , , "ÌåÉ ÇáÇÓÊŞØÇÚ ", 0
.ColumnHeaders.Add , , "ÇÓÊŞØÇÚ 2 ", 0
.ColumnHeaders.Add , , "ÌåÉ ÇáÇÓÊŞØÇÚ 2 ", 0
.ColumnHeaders.Add , , "ÇÓÊŞØÇÚÇÊ ÑÚÇíå", 0
.ColumnHeaders.Add , , " ãÈáÛ ÇáÛíÇÈ", 0
.ColumnHeaders.Add , , " ãÌãæÚ ÇáÇÓÊŞØÇÚÇÊ", 0
.ColumnHeaders.Add , , "ÇáÑÇÊÈ ÇáÕÇİí", 150

ListView1.Font = Bold
ListView1.Font.Size = 16
ListView1.Font.Name = "PT hrading"
End With
'======================================== ãÑÍáÉ ÇãáÇÁ ÇáÈíÇäÇÊ
With Me.ListView1
'Setting  ÇÚÏÇÏÇÊ ÇáãÙåÑ
.FullRowSelect = True 'áÊÙáíá ßÇãá ÇáÕİ
.Gridlines = True
.View = lvwReport  'ÚÑÖ ÎØæØ ÇáÔÈßÉ
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' áæä ÇáÎØ
.Font.Size = 11 'ÍÌã ÇáÎØ
.Font.Bold = False  'ÛÇãŞ
Font.Name = "Calibri"

With Me.ListView1
Dim item1 As ListItem
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView1.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")
item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(13) = Sheets(1).Cells(frw1, "n")
item1.SubItems(15) = Sheets(1).Cells(frw1, "p")
item1.SubItems(17) = Sheets(1).Cells(frw1, "r")
item1.SubItems(18) = Sheets(1).Cells(frw1, "s")
item1.SubItems(19) = Sheets(1).Cells(frw1, "t")
item1.SubItems(20) = Sheets(1).Cells(frw1, "u")
item1.SubItems(24) = Sheets(1).Cells(frw1, "y")
item1.SubItems(26) = Sheets(1).Cells(frw1, "aa")
item1.SubItems(28) = Sheets(1).Cells(frw1, "ac")
item1.SubItems(29) = Sheets(1).Cells(frw1, "ad")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(48) = Sheets(1).Cells(frw1, "aw")
Next frw1
'=============================================
 Dim Columns As Long
Dim i As Long
With ListView1
Columns = ListView1.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next

    ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(7).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(8).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(9).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(10).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(11).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(12).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(13).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(14).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(15).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(16).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(17).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(18).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(19).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(20).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(21).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(22).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(23).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(24).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(25).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(26).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(27).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(28).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(29).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(30).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(36).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(37).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(38).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(39).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(40).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(41).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(42).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(43).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(44).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(45).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(49).Alignment = lvwColumnCenter
   ListView1.ColumnHeaders(50).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(52).Alignment = lvwColumnCenter
   ListView1.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(55).Alignment = lvwColumnCenter
   ListView1.ColumnHeaders(56).Alignment = lvwColumnCenter
   ListView1.ColumnHeaders(57).Alignment = lvwColumnCenter

End With
End With
End With
ListView1.Font = Bold
ListView1.Font.Size = 14
ListView1.Font.Name = "PT hrading"
End Sub

Private Sub UserForm_Initialize()
On Error Resume Next
Sheets(5).Activate
ComboBox1.List = Sheets(5).Range("a1:a373").value
With ListView2
.Gridlines = True
.View = lvwReport
.FullRowSelect = True
.ColumnHeaders.Add , , "ÇáÑŞã", 40
.ColumnHeaders.Add , , "ÇáÑŞã ÇáæÙíİí", 120
.ColumnHeaders.Add , , "ÑŞã ÇáãÇÓÊÑ", 0
.ColumnHeaders.Add , , "ÇáÌäÓ", 0
.ColumnHeaders.Add , , "ÇÓã ÇáãæÙİ", 140
.ColumnHeaders.Add , , "ÇáÚäæÇä ÇáæÙíİí", 100
.ColumnHeaders.Add , , " ÇáŞÓã", 100
.ColumnHeaders.Add , , " ÇáÏÑÌÉ", 0
.ColumnHeaders.Add , , "ÇáãÑÍáÉ", 0
.ColumnHeaders.Add , , "ÇáÑÇÊÈ ÇáÇÓãí ", 140
.ColumnHeaders.Add , , "ÇíÇã ÇáÛíÇÈ", 0
.ColumnHeaders.Add , , " ÇíÇã ÇáÏæÇã", 0
.ColumnHeaders.Add , , " äÓÈÉ ÇáÔåÇÏÉ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÇáÔåÇÏÉ", 0
.ColumnHeaders.Add , , "äÓÈÉ ÇáåäÏÓíå ", 0
.ColumnHeaders.Add , , " ãÎÕÕÇÊ ÇáåäÏÓíå", 0
.ColumnHeaders.Add , , "äÓÈÉ ÇáãäÕÈ ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÇáãäÕÈ ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ãæŞÚ ÌÛÑÇİí", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÇØİÇá  ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÒæÌíå", 0
.ColumnHeaders.Add , , " ÚÏÏ ÇáÇØİÇá", 0
.ColumnHeaders.Add , , "ÇáÍÇáå ÇáÇÌÊãÇÚíå ", 0
.ColumnHeaders.Add , , "ÇáÇÓÊŞØÇÚ ÇáÖÑíÈí", 120
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ÎØæÑÉ ãÈáÛ ", 0
.ColumnHeaders.Add , , "ãÎÕÕÇÊ ãåäíå ", 0
.ColumnHeaders.Add , , " ãäÓÈÉ ãÎÕÕÇÊ ãåäíå ", 0
.ColumnHeaders.Add , , "äÓÈÉ ÇáÎØæÑå", 0
.ColumnHeaders.Add , , " ãÎÕÕÇÊ ÇáÎØæÑÉ ", 0
.ColumnHeaders.Add , , "ãÌãæÚ ÇáãÎÕÕÇÊ", 0
.ColumnHeaders.Add , , "ãÌãæÚ ÇáÇÓÊÍŞÇŞÇÊ", 120
.ColumnHeaders.Add , , " ÊæŞİÇÊ ÊŞÇÚÏíå 10", 100
.ColumnHeaders.Add , , "ÇÊæŞİÇÊ ÊŞÇÚÏíå 15%", 120
.ColumnHeaders.Add , , "ÇáÊÇãíä ÇáÕÍí ", 0
.ColumnHeaders.Add , , " ÇÌæÑ ÇáäŞá", 120
.ColumnHeaders.Add , , " ÇáäİŞÉ", 120
.ColumnHeaders.Add , , " ÑÓã ÇáØÇÈÚ", 0
.ColumnHeaders.Add , , " æÒÇÑÉ ÇáãÇáíå", 0
.ColumnHeaders.Add , , "ÇáÇÓÊŞØÇÚ1", 120
.ColumnHeaders.Add , , " ÌåÉ ÇáÇÓÊŞØÇÚ1  ", 0
.ColumnHeaders.Add , , "ÇÓÊŞØÇÚ ãÊİÑŞ", 120
.ColumnHeaders.Add , , " ÌåÉ ÇáÇÓÊŞØÇÚ ", 0
.ColumnHeaders.Add , , " äÇÏí ÇáåíÃÉ ÇáÕäÇÚÇÊ ÇáÍÑÈí", 0
.ColumnHeaders.Add , , "ÌåÉ ÇáÇÓÊŞØÇÚ ", 0
.ColumnHeaders.Add , , "ÇÓÊŞØÇÚ 2 ", 0
.ColumnHeaders.Add , , "ÌåÉ ÇáÇÓÊŞØÇÚ 2 ", 0
.ColumnHeaders.Add , , "ÇÓÊŞØÇÚÇÊ ÑÚÇíå", 0
.ColumnHeaders.Add , , " ãÈáÛ ÇáÛíÇÈ", 120
.ColumnHeaders.Add , , " ãÌãæÚ ÇáÇÓÊŞØÇÚÇÊ", 120
.ColumnHeaders.Add , , "ÇáÑÇÊÈ ÇáÕÇİí", 150

End With
'======================================== ãÑÍáÉ ÇãáÇÁ ÇáÈíÇäÇÊ
With Me.ListView2
'Setting  ÇÚÏÇÏÇÊ ÇáãÙåÑ
.FullRowSelect = True 'áÊÙáíá ßÇãá ÇáÕİ
.Gridlines = True
.View = lvwReport  'ÚÑÖ ÎØæØ ÇáÔÈßÉ
.AllowColumnReorder = True
.ForeColor = RGB(84, 53, 110)  ' áæä ÇáÎØ
.Font.Size = 11 'ÍÌã ÇáÎØ
.Font.Bold = False  'ÛÇãŞ
Font.Name = "Calibri"

With Me.ListView2
Dim item1 As ListItem
Dim last1, frw1 As Integer
last1 = Sheets(1).Range("A" & Rows.count).End(xlUp).row
For frw1 = 9 To last1
Set item1 = ListView2.ListItems.Add(, , Sheets(1).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(1).Cells(frw1, "b")

item1.SubItems(4) = Sheets(1).Cells(frw1, "e")
item1.SubItems(5) = Sheets(1).Cells(frw1, "f")
item1.SubItems(6) = Sheets(1).Cells(frw1, "g")
item1.SubItems(9) = Sheets(1).Cells(frw1, "j")
item1.SubItems(23) = Sheets(1).Cells(frw1, "x")
item1.SubItems(30) = Sheets(1).Cells(frw1, "ae")
item1.SubItems(31) = Sheets(1).Cells(frw1, "af")
item1.SubItems(32) = Sheets(1).Cells(frw1, "ag")
item1.SubItems(33) = Sheets(1).Cells(frw1, "ah")
item1.SubItems(34) = Sheets(1).Cells(frw1, "ai")
item1.SubItems(35) = Sheets(1).Cells(frw1, "aj")
item1.SubItems(37) = Sheets(1).Cells(frw1, "al")
item1.SubItems(39) = Sheets(1).Cells(frw1, "an")
item1.SubItems(41) = Sheets(1).Cells(frw1, "ap")
item1.SubItems(43) = Sheets(1).Cells(frw1, "ar")
item1.SubItems(45) = Sheets(1).Cells(frw1, "at")
item1.SubItems(47) = Sheets(1).Cells(frw1, "au")
item1.SubItems(48) = Sheets(1).Cells(frw1, "av")
item1.SubItems(49) = Sheets(1).Cells(frw1, "aw")
Next frw1
Dim Columns As Long
Dim i As Long
With ListView2
Columns = ListView2.ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next

    ListView2.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView2.ColumnHeaders(2).Alignment = lvwColumnCenter
    
    ListView2.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(7).Alignment = lvwColumnCenter
       ListView2.ColumnHeaders(10).Alignment = lvwColumnCenter
    
    
    
   
    
    
    
    ListView2.ColumnHeaders(24).Alignment = lvwColumnCenter
   
    
    
    ListView2.ColumnHeaders(31).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(32).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(33).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(34).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(35).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(36).Alignment = lvwColumnCenter
   
    ListView2.ColumnHeaders(38).Alignment = lvwColumnCenter
   
    ListView2.ColumnHeaders(40).Alignment = lvwColumnCenter
    
    
    ListView2.ColumnHeaders(42).Alignment = lvwColumnCenter
  
    ListView2.ColumnHeaders(44).Alignment = lvwColumnCenter
   
    ListView2.ColumnHeaders(46).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(47).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(48).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(49).Alignment = lvwColumnCenter
   
    
    ListView2.ColumnHeaders(51).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(52).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(53).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(54).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(55).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(56).Alignment = lvwColumnCenter
   ListView2.ColumnHeaders(57).Alignment = lvwColumnCenter
    ListView2.ColumnHeaders(58).Alignment = lvwColumnCenter
'=================================================
End With
End With
End With
ListView2.Font = Bold
ListView2.Font.Size = 14
ListView2.Font.Name = "PT hrading"
End Sub
