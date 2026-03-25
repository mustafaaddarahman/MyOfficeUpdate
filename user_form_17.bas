Attribute VB_Name = "user_form_17"
Sub all_data_userform17()

On Error Resume Next
'Sheets(28).Range("a9:bv2000").ClearContents
'ActiveSheet.Range("a9: bv2000").UnMerge
'Dim arr001 As Variant, i As Long, j As Long, k As Long
'Dim arr002 As Variant
'Dim lastr As Long

'Sheets(27).Range("a1:be200").ClearContents
'lastr = Sheets(13).Cells(Rows.Count, 2).End(xlUp).Row

'arr001 = Sheets(13).Range("a9:bv" & lastr)
'j = 8
'ReDim arr002(9 To UBound(arr001, 1), 1 To UBound(arr001, 2))

'For i = 1 To UBound(arr001, 1)

' If arr00 = Sheets(13).Range("a9:bv" & lastr) Then
' j = j + 1
' For k = 1 To 74
' arr002(j, k) = arr001(i, k)
  
 ' Next k
  
'End If

'Next


'Sheets(28).Range("a9").Resize(j, 74).Value = arr002

'=============================
Sheets(28).Range("a9: bz8000").UnMerge
Sheets(28).Range("a9:bz8000").ClearContents

Sheets(28).Range("a9:bz8000").Interior.Color = &HFFFFFF
Sheets(28).Range("a9:bz8000").Font.Color = vbBlack



Dim arr001 As Variant, i As Long, j As Long, k As Long
Dim arr002 As Variant
Dim lastR As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets(13).Cells(Rows.count, 2).End(xlUp).row + 1
arr001 = Sheets(13).Range("a9:bz" & lastR)
j = 0
ReDim arr002(1 To UBound(arr001, 1), 1 To UBound(arr001, 2))

For i = 1 To UBound(arr001, 1)

 
 j = j + 1
 For k = 1 To 78
 arr002(j, k) = arr001(i, k)
  
  Next k
  


Next

Sheets(28).Range("a9").Resize(j, 78).value = arr002
Sheets(28).Activate


End Sub

