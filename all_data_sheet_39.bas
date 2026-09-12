Attribute VB_Name = "all_data_sheet_39"
Sub all_data()
On Error Resume Next

Sheets(39).Range("a9:cc8000").ClearContents
ActiveSheet.Range("a2: cc6000").UnMerge
Range("a9:cc8000").Interior.Color = &HFFFFFF
Range("a9:cc8000").Font.Color = vbBlack
Dim arr001 As Variant, i As Long, j As Long, k As Long
Dim arr002 As Variant
Dim lastR As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets(3).Cells(Rows.count, 1).End(xlUp).row
arr001 = Sheets(3).Range("a1:cc" & lastR)
j = 0
ReDim arr002(1 To UBound(arr001, 1), 1 To UBound(arr001, 2))

For i = 1 To UBound(arr001, 1)


 j = j + 1
 For k = 1 To 69
 arr002(j, k) = arr001(i, k)
  
  Next k
  


Next


Sheets(39).Range("a8").Resize(j, 69).value = arr002
Sheets(39).Activate
'=============================







End Sub
