Attribute VB_Name = "count_all_salalry"
Sub count_all()

Dim rng As Range
Dim cel As Range
Dim lastr1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, "b").End(xlUp).row
Set rng = Sheets(1).Range("au9:au" & lastr1)
For Each cel In rng
If IsEmpty(cel) Or Not IsNumeric(cel) Then
cel = 0
End If
Next cel




'????? ????????
'=N10+P10+R10+S10+T10+U10+Y10+AA10+AC10+BH10+BP10+BR10
Dim lastR As Long

lastR = Sheets(13).Cells(Rows.count, 2).End(xlUp).row



Ê—ﬁ…1.Cells(9, "n").Formula = "= j9*m9"
Ê—ﬁ…1.Range("n9:n" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "p").Formula = "= j9*o9"
Ê—ﬁ…1.Range("p9:p" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "r").Formula = "= j9*q9"
Ê—ﬁ…1.Range("r9:r" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "t").Formula = "= v9*10000"
Ê—ﬁ…1.Range("t9:t" & lastR).FillDown




Ê—ﬁ…1.Cells(9, "aa").Formula = "= j9*z9"
Ê—ﬁ…1.Range("aa9:aa" & lastR).FillDown

Ê—ﬁ…1.Cells(9, "ac").Formula = "= j9*ab9"
Ê—ﬁ…1.Range("ac9:ac" & lastR).FillDown


Ê—ﬁ…1.Range("ad9").Formula = "=N9+P9+R9+S9+T9+U9+Y9+AA9+AC9+BH9+BP9+BR9"
Ê—ﬁ…1.Range("ad9:ad" & lastR).FillDown





'????? ???????????
'=N9+P9+R9+S9+T9+U9+Y9+AA9+AC9+BH9+BP9+BR9
Ê—ﬁ…1.Range("ae9").Formula = "=ad9+j9+ce9"
Ê—ﬁ…1.Range("ae9:ae" & lastR).FillDown










' ????? ???????????

'=x9+af9+ag9+au9+ai9+aj9+al9+an9+ap9+ar9+at9+au9+bo9+bq9


Ê—ﬁ…1.Cells(9, "af").Formula = "=j9*0.1"
Ê—ﬁ…1.Range("af9:af" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "ap").Formula = "=IF(J9<=500000,1000,1500)"
Ê—ﬁ…1.Range("ap9:ap" & lastR).FillDown


 ' ????1.Cells(9, "al").Formula = "0"
 ' ????1.Range("al9:al" & lastr).FillDown

  
'????1.Range("au9").Formula = 0

'????1.Range("au9:au" & lastr).FillDown



Ê—ﬁ…1.Range("av9").Formula = "=x9+af9+ag9+ah9+ai9+aj9+al9+an9+ap9+ar9+at9+au9+bj9+bo9+bq9+ax9"
Ê—ﬁ…1.Range("av9:av" & lastR).FillDown



' ?????? ??????
'=IF(OR(BF9="????? ???????";BF9="????? ??? 6 ????";BF9="????? ???? 6 ????");(AE9+BP9+BR9)-(AV9+BO9+BQ9);AE9-AV9)


' ????1.Cells(9, "au").Formula = "=AW9-(AW9/30)*k9"
'  ????1.Range("au9:au" & lastr).FillDown

Ê—ﬁ…1.Range("aw9").Formula = "=ae9-av9"

Ê—ﬁ…1.Range("aw9:aw" & lastR).FillDown




End Sub

