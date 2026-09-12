Attribute VB_Name = "t6ab8"
Public Sub Force_Sync_Sheets()
    ' ÅíŞÇİ ÊÍÏíË ÇáÔÇÔÉ áÖãÇä Ãä ÇáÚãáíÉ ÊÊã Îáİ ÇáßæÇáíÓ
    Application.ScreenUpdating = False
    
    ' ÊİÚíá æÖÚ ÇáÍÓÇÈ ÇáÊáŞÇÆí áÖãÇä ÇÓÊÌÇÈÉ ÇáÏæÇá
    Application.Calculation = xlCalculationAutomatic
    
    ' "ÇáŞÏÍÉ": ÊäÔíØ ÇáÔíÊÇÊ ÈÇáÊÑÊíÈ áÖãÇä ÊÓáÓá ÇáÍÓÇÈ
    Sheets(1).Activate   ' ÇáÔíÊ ÇáãÑßÒí (ŞÇÚÏÉ ÇáÈíÇäÇÊ)
    Sheets(18).Activate  ' ÔíÊ ÇáãÚÇÏáÇÊ
    Sheets(26).Activate  ' ÔíÊ ÇáÎáÇÕÇÊ
    
    ' ÅÌÈÇÑ ÇáãÍÑß Úáì ÅÚÇÏÉ ÇáÍÓÇÈ Çáßáí áÑÈØ ÇáäÊÇÆÌ
    Application.CalculateFull
    
    ' ÇáÓãÇÍ ááÈÑäÇãÌ ÈÅäåÇÁ ÇáÚãáíÇÊ ÇáãÚáŞÉ
    DoEvents
    
    ' (ÇÎÊíÇÑí) ÇáÚæÏÉ áÔíÊ ÇáæÇÌåÉ ÅĞÇ ßäÊ ÊİÖá Ğáß¡ Ãæ ÇáÈŞÇÁ İí ÇáÎáİíÉ
    ' Sheets("ÇáæÇÌåÉ").Activate
    
    Application.ScreenUpdating = True
End Sub
