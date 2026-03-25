Attribute VB_Name = "sheets_raight_to_lift"
Sub AllSheetsRTL()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.DisplayRightToLeft = True
    Next ws
End Sub
