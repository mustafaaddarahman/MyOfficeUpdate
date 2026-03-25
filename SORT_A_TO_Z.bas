Attribute VB_Name = "SORT_A_TO_Z"
Sub SORT_A_TO_Z1()
 Range("A8:CB6485").Select
    ActiveWorkbook.Worksheets("sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("sheet1").Sort.SortFields.Add2 key:=Range( _
        "E9:E3341"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("sheet1").Sort
        .SetRange Range("A8:CB3341")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
