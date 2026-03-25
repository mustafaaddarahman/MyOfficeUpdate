Attribute VB_Name = "all_data_sheet_44"
Sub all_data_user39()
On Error Resume Next
Sheets(43).Activate
Sheets(44).Activate
Sheets(44).Range("a1:m100").ClearContents
Sheets(44).Range("a1: m100").UnMerge
Sheets(44).Range("a9:m100").Interior.Color = &HFFFFFF
Sheets(44).Range("a9:m100").Font.Color = vbBlack
Dim arr001 As Variant, i As Long, j As Long, k As Long
Dim arr002 As Variant
Dim lastR As Long
'Sheets(27).Range("a1:be200").ClearContents
lastR = Sheets(43).Cells(Rows.count, 1).End(xlUp).row
arr001 = Sheets(43).Range("a8:m" & lastR)
j = 0
ReDim arr002(1 To UBound(arr001, 1), 1 To UBound(arr001, 2))

For i = 1 To UBound(arr001, 1)


 j = j + 1
 For k = 1 To 13
 arr002(j, k) = arr001(i, k)
  
  Next k
  


Next


Sheets(44).Range("a8").Resize(j, 13).value = arr002

'=============================






End Sub
