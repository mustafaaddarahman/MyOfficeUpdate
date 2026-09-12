Attribute VB_Name = "mdl_SingleRowProcessor"
Sub Process_Vault_Array()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(1)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    If lastRow < 9 Then Exit Sub

    ' 1.  Õ„Ì· »Ì«‰«  «·√⁄„œ… AW Ê CH ≈·Ï „’›Ê›«  ›Ì «·–«ﬂ—…
    Dim arrAW As Variant, arrCH As Variant
    arrAW = ws.Range("AW1:AW" & lastRow).value
    arrCH = ws.Range("CH1:CH" & lastRow).value

    Dim i As Long
    Dim NetVal As Double, VaultVal As Double, TotalWithFract As Double

    ' 2. «·„⁄«·Ã… œ«Œ· «·–«ﬂ—… (”—Ì⁄… Ãœ«)
    For i = 9 To lastRow
        If IsNumeric(arrAW(i, 1)) Then
            NetVal = val(arrAW(i, 1))
            VaultVal = val(arrCH(i, 1))
            
            ' ¬·Ì… «·⁄“· Ê«·Ã„⁄
            TotalWithFract = Round(VaultVal + (NetVal - Fix(NetVal)), 4)
            
            If TotalWithFract >= 1 Then
                arrAW(i, 1) = Fix(NetVal) + 1
                arrCH(i, 1) = Round(TotalWithFract - 1, 4)
            Else
                arrAW(i, 1) = Fix(NetVal)
                arrCH(i, 1) = TotalWithFract
            End If
        End If
    Next i

    ' 3. ≈—Ã«⁄ «·‰ «∆Ã »«·ﬂ«„· ··‘Ì  »÷€ÿ… Ê«Õœ…
    ws.Range("AW1:AW" & lastRow).value = arrAW
    ws.Range("CH1:CH" & lastRow).value = arrCH
    ws.Range("AW9:AW" & lastRow).NumberFormat = "0"
End Sub
