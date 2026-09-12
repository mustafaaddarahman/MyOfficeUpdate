Attribute VB_Name = "count_all_salalry"
Sub count_all()

Application.Calculation = xlCalculationAutomatic

Dim rng As Range
Dim cel As Range
Dim lastr1 As Integer
lastr1 = Sheets(1).Cells(Rows.count, "b").End(xlUp).row
Set rng = Sheets(1).Range("au9:au" & lastr1)
For Each cel In rng
If IsEmpty(cel) Then
cel = 0
End If
Next cel





Dim lastR As Long

lastR = Sheets(1).Cells(Rows.count, 2).End(xlUp).row

Ê—ﬁ…1.Cells(9, "n").Formula = "=INT(j9*m9)"
Ê—ﬁ…1.Range("n9:n" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "p").Formula = "=INT(j9*o9)"
Ê—ﬁ…1.Range("p9:p" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "r").Formula = "=INT(j9*q9)"
Ê—ﬁ…1.Range("r9:r" & lastR).FillDown

Dim lastr3 As Long
 lastr3 = Sheets(1).Cells(Rows.count, "B").End(xlUp).row
     Sheets(1).Cells(9, "T").Formula = "=INT(IF(BF9=""Ì „ ⁄ »«Ã«“Â Œ„” ”‰Ê« "", V9*0, IF(V9>4, 4*10000, V9*10000)))"
    Sheets(1).Range("T9:T" & lastr3).FillDown



Ê—ﬁ…1.Cells(9, "aa").Formula = "=INT(j9*z9)"
Ê—ﬁ…1.Range("aa9:aa" & lastR).FillDown

Ê—ﬁ…1.Cells(9, "ac").Formula = "= INT(j9*ab9)"
Ê—ﬁ…1.Range("ac9:ac" & lastR).FillDown


Ê—ﬁ…1.Range("ad9").Formula = "=INT(N9+P9+R9+S9+T9+U9+Y9+AA9+AC9+BH9+BP9+BR9)"
Ê—ﬁ…1.Range("ad9:ad" & lastR).FillDown


Ê—ﬁ…1.Range("ae9").Formula = "=INT(ad9+j9+ce9)"
Ê—ﬁ…1.Range("ae9:ae" & lastR).FillDown


Range("AT9:AT" & lastR).Formula = "=INT(J9*0.0025)"
Sheets(1).Range("AT9:AT" & lastr3).FillDown


Ê—ﬁ…1.Cells(9, "af").Formula = "=INT(j9*0.1)"
Ê—ﬁ…1.Range("af9:af" & lastR).FillDown


Ê—ﬁ…1.Cells(9, "AP").Formula = "=INT(IF(H9=0,0,IF(H9<=6,1500,1000)))"
Ê—ﬁ…1.Range("AP9:AP" & lastR).FillDown



Ê—ﬁ…1.Range("av9").Formula = "=INT(X9+AF9+AG9+AH9+AI9+AJ9+AL9+AN9+AP9+AR9+AT9+AU9+BJ9+BO9+BQ9+AX9)"
Ê—ﬁ…1.Range("av9:av" & lastR).FillDown

Ê—ﬁ…1.Range("aw9").Formula = "=INT(ae9-av9)"
Ê—ﬁ…1.Range("aw9:aw" & lastR).FillDown





End Sub

