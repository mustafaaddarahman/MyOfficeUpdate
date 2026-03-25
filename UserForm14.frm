VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm14 
   Caption         =   "UserForm14"
   ClientHeight    =   11025
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   14376
   OleObjectBlob   =   "UserForm14.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton19_Click()
Dim lastR As Integer
lastR = Sheets(21).Cells(Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(21).Cells(Y, 1) = TextBox1.Text And Sheets(21).Cells(Y, 3) = TextBox60.Text Then
   
    Exit For
    End If
    Next Y
    Sheets(21).Cells(Y, 1) = TextBox1.Text
    Sheets(21).Cells(Y, 2) = TextBox2.Text
    Sheets(21).Cells(Y, 3) = TextBox60.Text
    Sheets(21).Cells(Y, 4) = TextBox3.Text
    Sheets(21).Cells(Y, 5) = Format(TextBox4.Text, "yyyy/mm/dd")
    Sheets(21).Cells(Y, 6) = TextBox5.Text


MsgBox " „   ⁄œÌ· «·”Ã· »‰Ã«Õ", vbMsgBoxRight + vbOKOnly, "⁄„«œ «·‰⁄Ì„Ì"
Call auto_num
Call refrash_listview

End Sub

Private Sub CommandButton20_Click()
On Error GoTo err_command30_click
Dim pathscaner As String
pathscaner = "c:\windows\system32\wiaacmgr.exe"
Call Shell(pathscaner, 1)
exit_command30_click:
Exit Sub
err_command30_click:
 MsgBox Err.Description
Resume exit_command30_click

End Sub

Private Sub CommandButton21_Click()
On Error GoTo emad
Dim get_path As String
Dim newfold As FileSystemObject
Set newfold = New FileSystemObject
newfold.CreateFolder ("e:\" & TextBox60.Text & "\")

With Application.FileDialog(msoFileDialogFilePicker)
If .Show <> 0 Then
get_path = .SelectedItems(1)

End If
TextBox5.Text = get_path
End With
emad:
End Sub

Private Sub CommandButton22_Click()
On Error GoTo emad
Dim lastR As Integer
Dim file_path As String
lastR = Sheets(21).Cells(Rows.count, 1).End(xlUp).row
    For Y = 2 To lastR
    
    If Sheets(21).Cells(Y, 1) = TextBox1.Text And Sheets(21).Cells(Y, 2) = TextBox2.Text Then
   
    Exit For
    End If
    Next Y
    
  
   
   'Call emad_pro

 '=========================================
 file_path = TextBox5.Text
 
 TextBox5.Text = file_path
 SavePicture Image1.Picture, TextBox5.Text

Sheets(21).Cells(Y, 6) = TextBox5.Text
MsgBox " „ Õ›Ÿ «·’Ê—…"


emad:

End Sub

Private Sub CommandButton23_Click()

Dim lastRow As Integer

lastRow = Sheets(21).Cells(Rows.count, 1).End(xlUp).row

 
    Sheets(21).Range("A" & lastRow + 1).Offset(0, 0).value = TextBox1.Text
    Sheets(21).Range("a" & lastRow + 1).Offset(0, 1).value = TextBox2.Text
    Sheets(21).Range("a" & lastRow + 1).Offset(0, 2).value = TextBox60.Text
    Sheets(21).Range("a" & lastRow + 1).Offset(0, 3).value = TextBox3.Text
    Sheets(21).Range("a" & lastRow + 1).Offset(0, 4).value = Format(TextBox4.Text, "yyyy/mm/dd")
    Sheets(21).Range("a" & lastRow + 1).Offset(0, 5).value = TextBox5.Text
    



TextBox60.Text = ""
TextBox2.Text = ""
TextBox61.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
TextBox1.Text = ""

'Dim frw1 As Integer
'Dim lastrow1 As Integer
'Dim item1 As ListItem
'lastrow1 = Sheets(21).Cells(Rows.Count, 2).End(xlUp).Row
'For frw1 = 2 To lastrow1

'Set item1 = ListView1.ListItems.Add(, , Sheets(21).Cells(frw1, "A"))
'item1.SubItems(1) = Sheets(21).Cells(frw1, "b")
'item1.SubItems(2) = Sheets(21).Cells(frw1, "c")
'item1.SubItems(3) = Sheets(21).Cells(frw1, "d")
'item1.SubItems(4) = Format(Sheets(21).Cells(frw1, "e"), "yyyy/mm/dd")
Call auto_num
'item1.SubItems(5) = Sheets(21).Cells(frw1, "f")
Call refrash_listview
'Next

TextBox60.Text = ""

End Sub

Private Sub CommandButton24_Click()
On Error GoTo emad
'Create and set dialog box as variable
Dim dialogBox As FileDialog
Set dialogBox = Application.FileDialog(msoFileDialogOpen)

'Do not allow multiple files to be selected
dialogBox.AllowMultiSelect = False

'Set the title of the DialogBox
dialogBox.title = "Select a file"

'Set the default folder to open
dialogBox.InitialFileName = UserForm14.TextBox5.Text

'Clear the dialog box filters
dialogBox.Filters.Clear
'Apply file filters - use ; to separate filters for the same name
dialogBox.Filters.Add "image_document", "*.JPG;*.JPEG;*.PNG;*.TIFF;*.GIF;*.PDF;*.AI;*.RAW;*.xlsx;*.xls;*.xlsm"

'Show the dialog box and output full file name
If dialogBox.Show = -1 Then
  UserForm14.TextBox5.Text = dialogBox.SelectedItems(1)
  Dim last1, frw1 As Integer
last1 = Sheets(21).Range("A" & Rows.count).End(xlUp).row
For frw1 = 6 To last1
If Sheets(21).Cells(frw1, "b") = TextBox2.Text And Sheets(21).Cells(frw1, "c") = TextBox3.Text Then


End If
Next
End If
Image1.Picture = LoadPicture(TextBox5.Text)
emad:
End Sub

Private Sub CommandButton25_Click()


Dim lastRow As Integer

lastRow = Sheets(21).Cells(Rows.count, 2).End(xlUp).row
TextBox1.value = lastRow
Call auto_num
End Sub

Private Sub CommandButton26_Click()
Dim lastR, r As Integer
lastR = Sheets(21).Cells(Rows.count, 1).End(xlUp).row
For r = 2 To lastR
If TextBox1.Text = Sheets(21).Cells(r, 1).value Then

Sheets(21).Cells(r, 1).Select
Exit For
End If
Next
ActiveCell.EntireRow.Resize(, 6).Delete
MsgBox " „ Õ–› «·”Ã· »‰Ã«Õ", vbMsgBoxRight + vbOKOnly, "⁄„«œ «·‰⁄Ì„Ì"
Call auto_num
Call refrash_listview
End Sub

Private Sub CommandButton27_Click()
Unload Me
End Sub

Private Sub ListView1_Click()
On Error GoTo emad
TextBox1.Text = ListView1.SelectedItem

TextBox2.Text = ListView1.SelectedItem.SubItems(1)
TextBox3.Text = ListView1.SelectedItem.SubItems(3)
TextBox4.Text = Format(ListView1.SelectedItem.SubItems(4), "yyyy/mm/dd")
TextBox5.Text = ListView1.SelectedItem.SubItems(5)
TextBox60.Text = ListView1.SelectedItem.SubItems(2)
Image1.Picture = LoadPicture(TextBox5.Text)
emad:
End Sub







Private Sub ListView1_KeyDown(KeyCode As Integer, ByVal Shift As Integer)
On Error GoTo emad
TextBox1.Text = ListView1.SelectedItem

TextBox2.Text = ListView1.SelectedItem.SubItems(1)
TextBox3.Text = ListView1.SelectedItem.SubItems(3)
TextBox4.Text = Format(ListView1.SelectedItem.SubItems(4), "yyyy/mm/dd")
TextBox5.Text = ListView1.SelectedItem.SubItems(5)
TextBox60.Text = ListView1.SelectedItem.SubItems(2)
Image1.Picture = LoadPicture(TextBox5.Text)
emad:
End Sub

Private Sub TextBox61_Change()
On Error GoTo emad
Sheets(21).Cells(2, 9) = TextBox61.Text
TextBox60.Text = Sheets(21).Cells(2, 10).Text
TextBox2.Text = TextBox61.Text
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""




emad:
End Sub

Private Sub TextBox61_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ListView1.ListItems.Clear
TextBox60.Text = ""
TextBox2.Text = ""
TextBox61.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
TextBox1.Text = ""
 Image1.Picture = LoadPicture(TextBox5.Text)
End Sub

Private Sub TextBox61_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
TextBox60.Text = ""
TextBox61.Text = ""
End Sub



Private Sub TextBox62_Change()
'On Error GoTo emad

Dim item1 As ListItem
Dim last2, frw1 As Integer

 With Me.ListView1
.ListItems.Clear
Sheets(21).Activate
last2 = Sheets(21).Range("A" & Rows.count).End(xlUp).row
For frw1 = 2 To last2

If Sheets(21).Cells(frw1, "c") Like "*" & TextBox62.Text & "*" Then
'==================================================
Set item1 = ListView1.ListItems.Add(, , Sheets(21).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(21).Cells(frw1, "b")
item1.ListSubItems(1).ForeColor = RGB(30, 70, 180)
item1.SubItems(2) = Sheets(21).Cells(frw1, "c")
item1.ListSubItems(2).Bold = True
item1.ListSubItems(2).ForeColor = RGB(84, 130, 53)

item1.SubItems(3) = Sheets(21).Cells(frw1, "d")
item1.SubItems(4) = Format(Sheets(21).Cells(frw1, "e"), "yyyy/mm/dd")
item1.SubItems(5) = Sheets(21).Cells(frw1, "f")
Image1.Picture = LoadPicture(TextBox5.Text)
End If
Next


End With

'===========================


'emad:
End Sub

Private Sub TextBox62_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call refrash_listview
TextBox60.Text = ""
TextBox62.Text = ""
TextBox63.Text = ""
TextBox2.Text = ""
TextBox61.Text = ""
TextBox3.Text = ""
TextBox4.Text = ""
TextBox5.Text = ""
TextBox1.Text = ""
Dim Columns As Long
Dim i As Long
With ListView1
Columns = .ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next
End With

'Next
End Sub



Private Sub TextBox63_Change()
'On Error GoTo emad
'TextBox62.Text = "-"
Dim frw2 As Integer
Dim lastrow1 As Integer

'lastrow1 = Sheets(21).Cells(Rows.Count, 2).End(xlUp).Row
'For frw2 = 2 To lastrow1
'If Sheets(21).Cells(frw2, 4) = TextBox63.Text Then
'TextBox1.Text = Sheets(21).Cells(frw2, "a")
'TextBox2.Text = Sheets(21).Cells(frw2, "b")
'TextBox3.Text = Sheets(21).Cells(frw2, "d")
'TextBox4.Text = Sheets(21).Cells(frw2, "e")
'TextBox5.Text = Sheets(21).Cells(frw2, "f")
'TextBox60.Text = Sheets(21).Cells(frw2, "c")
'Image1.Picture = LoadPicture(TextBox5.Text)
  
  'End If
 ' Next
'========================================
Dim r As Integer
Dim last_row As Integer
ListView1.ListItems.Clear
last_row = Sheets(21).Range("a" & Rows.count).End(xlUp).row
For r = 2 To last_row
If Sheets(21).Cells(r, 4) Like "*" & TextBox63.Text & "*" Then
 TextBox1.Text = Sheets(21).Cells(r, "a").value
TextBox2.Text = Sheets(21).Cells(r, "b").value
 TextBox60.Text = Sheets(21).Cells(r, "c").value
 TextBox3.Text = Sheets(21).Cells(r, "d").value
  TextBox4.Text = Format(Sheets(21).Cells(r, "e"), "yyyy/mm/dd")
  TextBox5.Text = Sheets(21).Cells(r, "f").value
 Image1.Picture = LoadPicture(TextBox5.Text)
  
  
 End If
 
Next




'========================================

 


'===============================
Dim frw1 As Integer
Dim lastRow As Integer
Dim item1 As ListItem
lastRow = Sheets(21).Cells(Rows.count, 2).End(xlUp).row
For frw1 = 2 To lastRow
If Sheets(21).Cells(frw1, 4) Like "*" & TextBox63.Text & "*" Then

Set item1 = ListView1.ListItems.Add(, , Sheets(21).Cells(frw1, "A"))

item1.SubItems(1) = Sheets(21).Cells(frw1, "b")
item1.ListSubItems(1).ForeColor = RGB(30, 70, 180)
item1.SubItems(2) = Sheets(21).Cells(frw1, "c")
item1.ListSubItems(2).Bold = True
item1.ListSubItems(2).ForeColor = RGB(84, 130, 53)

item1.SubItems(3) = Sheets(21).Cells(frw1, "d")
item1.SubItems(4) = Format(Sheets(21).Cells(frw1, "e"), "yyyy/mm/dd")
item1.SubItems(5) = Sheets(21).Cells(frw1, "f")
Image1.Picture = LoadPicture(TextBox5.Text)
End If
Next
'=======================

  '=========================
 
'  If TextBox62.Text = "" Then
  
'  TextBox60.Text = ""
'TextBox2.Text = ""
'TextBox61.Text = ""
'TextBox3.Text = ""
'TextBox4.Text = ""
'TextBox5.Text = ""
' Image1.Picture = LoadPicture(TextBox5.Text)
'End If

'emad:
End Sub

Private Sub TextBox63_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Call refrash_listview
End Sub

Private Sub UserForm_Activate()

Sheets(21).Activate
 ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter
    ListView1.ListItems.Clear
Dim frw1 As Integer
Dim lastRow As Integer
Dim item1 As ListItem
lastRow = Sheets(21).Cells(Rows.count, 2).End(xlUp).row
For frw1 = 2 To lastRow

Set item1 = ListView1.ListItems.Add(, , Sheets(21).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(21).Cells(frw1, "b")
item1.ListSubItems(1).ForeColor = RGB(30, 70, 180)

item1.SubItems(2) = Sheets(21).Cells(frw1, "c")
item1.ListSubItems(2).Bold = True
item1.ListSubItems(2).ForeColor = RGB(84, 130, 53)

item1.SubItems(3) = Sheets(21).Cells(frw1, "d")
item1.SubItems(4) = Format(Sheets(21).Cells(frw1, "e"), "yyyy/mm/dd")
item1.SubItems(5) = Sheets(21).Cells(frw1, "f")

Next
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

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Unload Me
UserForm14.Show

End Sub

Private Sub UserForm_Initialize()
TextBox4.Text = Format(TextBox4.Text, "yyyy/mm/dd")
ListView1.Gridlines = True
ListView1.View = lvwReport
ListView1.FullRowSelect = True
ListView1.ColumnHeaders.Add 1, , "«·—ﬁ„", 40
ListView1.ColumnHeaders.Add 2, , "«·—ﬁ„ «·ÊŸÌ›Ì", 100
ListView1.ColumnHeaders.Add 3, , "«”„ «·„ÊŸ›", 200
ListView1.ColumnHeaders.Add 4, , "—ﬁ„ «·ﬂ «»", 60
ListView1.ColumnHeaders.Add 5, , " «—ÌŒ «·ﬂ «» ", 100
ListView1.ColumnHeaders.Add 6, , "„”«— «·’Ê—Â ", 300
'==============================
 ListView1.ColumnHeaders(1).Alignment = lvwColumnLeft
    ListView1.ColumnHeaders(2).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(3).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(4).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(5).Alignment = lvwColumnCenter
    ListView1.ColumnHeaders(6).Alignment = lvwColumnCenter

'===============================

Dim Columns As Long
Dim i As Long
With ListView1
Columns = .ColumnHeaders.count
For i = 1 To Columns
.ColumnHeaders(i).Position = Columns - i + 1
Next
For i = 2 To Columns
.ColumnHeaders(i).Alignment = lvwColumnRight
Next
End With

'===============================
ListView1.ListItems.Clear
Dim frw1 As Integer
Dim lastRow As Integer
Dim item1 As ListItem
lastRow = Sheets(21).Cells(Rows.count, 2).End(xlUp).row
For frw1 = 2 To lastRow

Set item1 = ListView1.ListItems.Add(, , Sheets(21).Cells(frw1, "A"))
item1.SubItems(1) = Sheets(21).Cells(frw1, "b")
item1.ListSubItems(1).ForeColor = RGB(30, 70, 180)

item1.SubItems(2) = Sheets(21).Cells(frw1, "c")
item1.ListSubItems(2).Bold = True
item1.ListSubItems(2).ForeColor = RGB(84, 130, 53)

item1.SubItems(3) = Sheets(21).Cells(frw1, "d")
item1.SubItems(4) = Format(Sheets(21).Cells(frw1, "e"), "yyyy/mm/dd")
item1.SubItems(5) = Sheets(21).Cells(frw1, "f")

Next
End Sub
