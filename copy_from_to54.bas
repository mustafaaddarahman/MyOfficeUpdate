Attribute VB_Name = "copy_from_to54"
Sub copy_from_to54_1()

On Error Resume Next
Sheets(1).Activate


'===============================================================


Dim arr1 As Variant
Dim arr2 As Variant
Dim arr3 As Variant
arr1 = Sheets(1).Range("a8").CurrentRegion.value
For Each arr2 In Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 22, 20, 21, 32, 23, 24, 25, , 26, 27, 28, 29, 30, 31, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 51, 52, 53, 56, 61, 73, 57, 57)
arr3 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, , 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56)
Sheets(54).Cells(1, arr3(i)).Resize(UBound(arr1, 1)).value = Application.Index(arr1, , arr2)
i = i + 1
Next arr2
Sheets(54).Activate

End Sub

