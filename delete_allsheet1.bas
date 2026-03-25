Attribute VB_Name = "delete_allsheet1"
Sub DeleteDataToColumn84()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' ÊÍÏíÏ ÇáÔíÊ
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ÅíŞÇİ ÊÍÏíË ÇáÔÇÔÉ áÒíÇÏÉ ÇáÓÑÚÉ
    Application.ScreenUpdating = False
    
    ' 1. ÊÍÏíÏ ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ (ááÊÃßÏ ãä Ôãæá ßá ÔíÁ)
    ' Ãæ íãßäß ÇÓÊÎÏÇã ws.Rows.Count áãÓÍ Çáãáíæä Õİ ÈÇáßÇãá
    lastRow = ws.Rows.count
    
    ' 2. ãÓÍ ÇáäØÇŞ ãä ÇáÎáíÉ A9 Åáì ÂÎÑ Õİ æÇáÚãæÏ 84 (ÇáĞí åæ ÇáÚãæÏ CF)
    ' ãáÇÍÙÉ: ÇÓÊÎÏãäÇ Clear áÖãÇä ãÓÍ ÇáÈíÇäÇÊ æÇáÊäÓíŞÇÊ
    On Error Resume Next
    ws.Range(ws.Cells(9, 1), ws.Cells(lastRow, 84)).Clear
    On Error GoTo 0
    
    ' ÅÚÇÏÉ ÊÍÏíË ÇáÔÇÔÉ
    Application.ScreenUpdating = True
    
    ' MsgBox "Êã ãÓÍ ÇáÈíÇäÇÊ æÇáÊäÓíŞÇÊ áÛÇíÉ ÇáÚãæÏ 84 ãä ÇáÕİ 9 äÒæáÇğ.", vbInformation
End Sub
