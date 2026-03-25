Attribute VB_Name = "user_form_28"
Sub all_data_userform28()

On Error Resume Next

Sheets(36).Range("a9:ce8000").ClearContents
ActiveSheet.Range("a9: ce6000").UnMerge

Sheets(36).Range("a9:ce8000").Interior.Color = &HFFFFFF
Sheets(36).Range("a9:ce8000").Font.Color = vbBlack



Dim arr001 As Variant, i As Long, j As Long, k As Long
Dim arr002 As Variant
Dim lastR As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets(35).Cells(Rows.count, 1).End(xlUp).row + 1
arr001 = Sheets(35).Range("a9:ce" & lastR)
j = 0
ReDim arr002(1 To UBound(arr001, 1), 1 To UBound(arr001, 2))

For i = 1 To UBound(arr001, 1)


 j = j + 1
 For k = 1 To 83
 arr002(j, k) = arr001(i, k)
  
  Next k
  


Next


Sheets(36).Range("a9").Resize(j, 83).value = arr002
Sheets(36).Activate
'=============================







End Sub
