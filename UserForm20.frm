VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm20 
   Caption         =   "UserForm20"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   6972
   OleObjectBlob   =   "UserForm20.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()
Call Hideit
If UserForm20.CheckBox1.value = True Then
UserForm20.CheckBox2.value = False
Unload Me
Call getimporfilename
Call ImportData

ThisWorkbook.Application.Visible = False
Sheets(13).Range("a2:ce10000").ClearContents
Application.ScreenUpdating = False

'Sheets(5).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select

Selection.Copy Sheets(13).Range("A8")

Application.ScreenUpdating = True
[a8].Select
Sheets(13).Select
Sheets(13).Activate
Range("a:ce").EntireColumn.Hidden = False
Application.ScreenUpdating = True
'=======================================

Sheets(30).Range("a1:ce10000").ClearContents
Application.ScreenUpdating = False
'Sheets(5).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select

Selection.Copy Sheets(30).Range("A1")

Application.ScreenUpdating = True

Sheets(30).Select
[a8].Select
Sheets(30).Activate
Application.ScreenUpdating = True


'============================================ ⁄„Ì„ œ«·… «·÷—Ì»Â ⁄·Ï‘Ì  30
'Sheets(30).Range("x9").Formula = "=IF(AND(VLOOKUP(W9,zwjea4,2,0)<4,V9>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(J9,astktaat,VLOOKUP(W9,zwjea4,2,0)+V9,1))"
'Sheets(30).Range("x9:x3000").FillDown

'============================================
'============================================ ⁄„Ì„ œ«·… «·÷—Ì»Â ⁄·Ï‘Ì  1
'Sheets(1).Range("x9").Formula = "=IF(AND(VLOOKUP(W9,zwjea4,2,0)<4,V9>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(J9,astktaat,VLOOKUP(W9,zwjea4,2,0)+V9,1))"
'Sheets(1).Range("x9:x3000").FillDown

'============================================
'============================================ ⁄„Ì„ œ«·… «·÷—Ì»Â ⁄·Ï‘Ì  13
'Sheets(13).Range("x9").Formula = "=IF(AND(VLOOKUP(W9,zwjea4,2,0)<4,V9>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(J9,astktaat,VLOOKUP(W9,zwjea4,2,0)+V9,1))"
'Sheets(13).Range("x9:x3000").FillDown

'============================================




UserForm7.Show

Sheets(13).Range("a1:ce7").ClearContents
Sheets(30).Range("a1:ce7").ClearContents

End If








End Sub

Private Sub CheckBox2_Click()
Call Hideit
If UserForm20.CheckBox2.value = True Then
UserForm20.CheckBox1.value = False
Unload Me
ThisWorkbook.Application.Visible = False
Sheets(13).Range("a1:ce10000").ClearContents
Application.ScreenUpdating = False
'Sheets(5).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select

Selection.Copy Sheets(30).Range("A1")

Application.ScreenUpdating = True

Sheets(13).Select
[a8].Select
Sheets(13).Activate
Application.ScreenUpdating = True
'=======================================

Sheets(30).Range("a1:ce10000").ClearContents
Application.ScreenUpdating = False
'Sheets(5).Range("a8").CurrentRegion.Delete
Sheets(1).Select

Range("a8").CurrentRegion.Select

Selection.Copy Sheets(30).Range("A1")

Application.ScreenUpdating = True

Sheets(30).Select
[a8].Select
Sheets(30).Activate
Application.ScreenUpdating = True


'============================================ ⁄„Ì„ œ«·… «·÷—Ì»Â ⁄·Ï‘Ì  30
'Sheets(30).Range("x9").Formula = "=IF(AND(VLOOKUP(W9,zwjea4,2,0)<4,V9>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(J9,astktaat,VLOOKUP(W9,zwjea4,2,0)+V9,1))"
'Sheets(30).Range("x9:x3000").FillDown

'============================================
'============================================ ⁄„Ì„ œ«·… «·÷—Ì»Â ⁄·Ï‘Ì  1
'Sheets(1).Range("x9").Formula = "=IF(AND(VLOOKUP(W9,zwjea4,2,0)<4,V9>0),""Õ”» «· ﬁ« ÿ⁄ «·ÃœÊ· «·÷—Ì»Ì «‰ ÌﬂÊ‰ «·⁄œœ ’›—"",VLOOKUP(J9,astktaat,VLOOKUP(W9,zwjea4,2,0)+V9,1))"
'Sheets(1).Range("x9:x3000").FillDown

'============================================
'============================================ ⁄„Ì„ œ«·… «·÷—Ì»Â ⁄·Ï‘Ì  13

'============================================



UserForm7.Show

Sheets(13).Range("a1:ce7").ClearContents
Sheets(30).Range("a1:ce7").ClearContents

End If
End Sub

Private Sub UserForm_Activate()
Sheets("sheet25").Range("a1:bm1000").ClearContents
End Sub

Private Sub UserForm_Click()

End Sub
