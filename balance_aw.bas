Attribute VB_Name = "balance_aw"
Sub RecalculateAnWithEntitlements()
  Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets("Sheet1")

    

    Dim lastRow As Long

    Dim i As Long

    Dim payType As Variant

    Dim cashCount As Long

    

    cashCount = 0

    lastRow = ws.Cells(ws.rowS.count, "aw").End(xlUp).row

    

    ' 1. ÍáŞÉ ÇáãÓÍ ãÚ ãØÇÈŞÉ ÇáÍÑæİ æÊÔÇÈååÇ

    For i = 9 To lastRow

        payType = ws.Cells(i, "ay").value

        

        ' ÇÓÊÎÏÇã Trim áÅÒÇáÉ ÇáİÑÇÛÇÊ æ InStr áÖãÇä ÇáãØÇÈŞÉ æÇáÊÔÇÈå

        If InStr(1, Trim(CStr(payType)), "ßÇÔ ÕİÑí", vbTextCompare) > 0 Then

            cashCount = cashCount + 1

            ws.Cells(i, "an").value = ""

        End If

    Next i

    

    MsgBox "ÚÏÏ ÇáÕİæİ ÇáÊí ÊØÇÈŞÊ ãÚ ÔÑØ 'ßÇÔ ÕİÑí' æÊã ãÓÍ an İíåÇ åæ: " & cashCount, vbInformation, "İÍÕ ÇáÊÔÛíá"

    

    ' 2. ÊÔÛíá ãæÏíæá count_all Ãæá ãÑÉ

    On Error Resume Next

    Call count_all

    On Error GoTo 0

    

    ' 3. ÍáŞÉ ÅÓäÇÏ ŞíãÉ aw Åáì an ááÕİæİ ÇáãØÇÈŞÉ

    For i = 9 To lastRow

        payType = ws.Cells(i, "ay").value

        

        If InStr(1, Trim(CStr(payType)), "ßÇÔ ÕİÑí", vbTextCompare) > 0 Then

            ws.Cells(i, "an").value = ws.Cells(i, "aw").value

        End If

    Next i

    

    ' 4. ÇáÊÍÏíË ÇáÃÎíÑ

    On Error Resume Next

    Call count_all

    On Error GoTo 0

    

   
    
    MsgBox "áŞÏ ÊãÊ ÚãáíÉ ÇáÇÍÊÓÇÈ ÈÔßá ÏŞíŞ ãÚ ÊäÙíã ÇáÑæÇÊÈ ÇáÕİÑíå æØÑíŞÉ ÏİÚåÇ ßÇÔ áã ÊÊÇËÑ İí Êáß ÇáÇÓÊŞØÇÚÇÊ." & vbCrLf & _
           "(ÚÏÏ ÇáÓÌáÇÊ ÇáãÍÏËÉ: " & cashCount & ")", vbInformation, "äÊíÌÉ ÇáÇÍÊÓÇÈ ÇáäåÇÆí"
  
  
  
  
  End Sub
