Attribute VB_Name = "Module15"
Sub automotive()
On Error Resume Next

Dim srchcell As Range
 
 
For Each srchcell In Range("f9:f800")
 '=VLOOKUP(B9;sheet25!A8:AG772;2;FALSE) «·⁄‰Ê«‰ «·ÊŸÌ›Ì

 
 
srchcell = Application.WorksheetFunction.VLookup(srchcell.Offset(0, -4), Sheets("sheet25").Range("a8:af800"), 2, 0)
 
 Next srchcell
 '========================================
  '=VLOOKUP(B9;sheet25!A8:AG772;26;FALSE) «·ﬁ”„
 Dim srchcell1 As Range
For Each srchcell1 In Range("g9:g800")
srchcell1 = Application.WorksheetFunction.VLookup(srchcell1.Offset(0, -5), Sheets("sheet25").Range("a8:af800"), 26, 0)
Next srchcell1
'================================================
'=VLOOKUP(B9;sheet25!A8:AG772;20;FALSE) «·œ—Ã…

Dim srchcell2 As Range
For Each srchcell2 In Range("h9:h800")
srchcell2 = Application.WorksheetFunction.VLookup(srchcell2.Offset(0, -6), Sheets("sheet25").Range("a8:af800"), 20, 0)
 

Next srchcell2


'==================================================

 '=VLOOKUP(B9;sheet25!A8:AG772;23;FALSE   «·„—Õ·…
Dim srchcell3 As Range
For Each srchcell3 In Range("i9:i800")
srchcell3 = Application.WorksheetFunction.VLookup(srchcell3.Offset(0, -7), Sheets("sheet25").Range("a8:af800"), 23, 0)

Next srchcell3








'===========================================
  '=VLOOKUP(B9;sheet25!A8:AG772;24;FALSE)  «·—« » «·«”„Ì
Dim srchcell4 As Range
For Each srchcell4 In Range("j9:j800")
srchcell4 = Application.WorksheetFunction.VLookup(srchcell4.Offset(0, -8), Sheets("sheet25").Range("a8:af800"), 24, 0)
 Next srchcell4









'==================================
'=VLOOKUP(B9;sheet25!A8:AG772;14;FALSE)  «—ÌŒ Õ’Ê»Â ⁄·Ï «·⁄·«Ê…

Dim srchcell5 As Range
For Each srchcell5 In Range("bb9:bb800")
srchcell5 = Application.WorksheetFunction.VLookup(srchcell5.Offset(0, -52), Sheets("sheet25").Range("a8:af800"), 14, 0)

Next srchcell5







'=========================================
 ' =VLOOKUP(B9;sheet25!A8:AG772;19;FALSE) «Ìﬁ«› «·⁄·«Ê…
Dim srchcell6 As Range
For Each srchcell6 In Range("bk9:bk800")
srchcell6 = Application.WorksheetFunction.VLookup(srchcell6.Offset(0, -61), Sheets("sheet25").Range("a8:af800"), 19, 0)

Next srchcell6

MsgBox "·ﬁœ  „  ⁄„·Ì…  ÕœÌÀ «·»Ì«‰«  «·Œ«’… »«·„ÊŸ›Ì‰ »⁄œ « ’«·Â« »‰Ÿ«„ «·⁄·«Ê« ", vbDefaultButton2, "—”«·…  ‰»ÌÂ"




End Sub
