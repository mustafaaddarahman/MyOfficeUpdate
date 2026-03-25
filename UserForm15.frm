VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm15 
   Caption         =   "ÊÕÝíÉ"
   ClientHeight    =   3795
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   5316
   OleObjectBlob   =   "UserForm15.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
'On Error GoTo emad
Sheets(27).Range("a2:bi200").ClearContents
Dim arr1, arr2, lastR As Variant, i As Long, j As Long, k As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets("sheet3").Cells(Rows.count, 1).End(xlUp).row
arr1 = Sheets("sheet3").Range("a2:bf" & lastR)
j = 0
ReDim arr2(1 To UBound(arr1, 1), 1 To UBound(arr1, 2))

For i = 1 To UBound(arr1, 1)

 If arr1(i, 58) Like "*" & TextBox1.Text & "*" And arr1(i, 56) >= CDate(TextBox2.Text) And arr1(i, 56) <= CDate(TextBox3.Text) Then

 j = j + 1
 For k = 1 To UBound(arr2, 2)
 arr2(j, k) = arr1(i, k)
  
  Next k
  
End If

Next

Sheets(27).Range("a2").Resize(j, UBound(arr2, 2)) = arr2
Sheets(27).Activate
'emad:
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
Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Activate()

End Sub

Private Sub UserForm_Click()

End Sub
