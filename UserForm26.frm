VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm26 
   Caption         =   "UserForm26"
   ClientHeight    =   3744
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10848
   OleObjectBlob   =   "UserForm26.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub ComboBox1_AfterUpdate()

End Sub

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

Sheets("title_factory").Cells(2, 1) = ComboBox1.value
Sheets("title_factory").Cells(2, 2) = TextBox1.Text
Sheets("title_factory").Cells(2, 3) = TextBox2.Text
Sheets("title_factory").Cells(2, 4) = Format(TextBox3.Text, "yyyy/mm/dd")
MsgBox "ÊãÊ ÊåíÆÉ ÇáÈíÇäÇÊ æÇÕÈÍ ÇáäÙÇã ÌÇåÒ ááÚãá", vbDefaultButton1, "ÑÓÇáÉ ÊæÌíå  "


End Sub

Private Sub CommandButton2_Click()
If Sheets("title_factory").Cells(2, 1) = "" And Sheets("title_factory").Cells(2, 2) = "" And Sheets("title_factory").Cells(2, 3) = "" And Sheets("title_factory").Cells(2, 4) = "" Then
MsgBox "áã ÊÞã ÈÊåíÆÉ Çí ÈíÇäÇÊ ÇÓÇÓíå áãÚÑÝÉ ÇÓã ÇáÊÔßíá Çæ ÇáãÕäÚ æãä åã ÇáãÓÊÎÏãíä áíÊÓäì ÇáÏÎæá Çáì ãäÕÉ ÇáäÙÇã", vbCritical, "ÑÓÇáÉ ÊäÈíå"
Unload Me
Application.Quit
Else
Unload Me
UserForm20.Show

End If
End Sub

Private Sub CommandButton3_Click()
If Trim(TextBox4.Text) = "" Then
        MsgBox "íÑÌì ßÊÇÈÉ ÇáÊÔßíá ÃæáÇð", vbExclamation, "ÊäÈíå"
        Exit Sub
    End If

    ' 2. ÅÖÇÝÉ ÇáÞíãÉ Åáì ÂÎÑ ÕÝ Ýí ÇáÚãæÏ B ÈÇáÔíÊ 5
    lastRow = Sheets(5).Cells(rowS.count, "B").End(xlUp).row + 1
    Sheets(5).Range("B" & lastRow).value = TextBox4.Text

    ' 3. ÊÍÏíË ComboBox1 ÝæÑÇð ãä ÇáÈíÇäÇÊ ÇáÌÏíÏÉ
    ' äÚíÏ ÍÓÇÈ ÂÎÑ ÕÝ ÈÚÏ ÇáÅÖÇÝÉ áÖãÇä Ôãæá ÇáÞíãÉ ÇáÌÏíÏÉ
    newRow = Sheets(5).Cells(rowS.count, "B").End(xlUp).row
    
    If newRow > 1 Then
        ComboBox1.List = Application.Transpose(Sheets(5).Range("B1:B" & newRow).value)
    Else
        ComboBox1.List = Array(Sheets(5).Range("B1").value)
    End If

    ' 4. ÊÝÑíÛ ÇáÊíßÓÊ ÈæßÓ æÅÚáÇã ÇáãÓÊÎÏã ÈÇáäÌÇÍ
    TextBox4.value = ""
    MsgBox "ÊãÊ ÅÖÇÝÉ ÇáÊÔßíá æÊÍÏíË ÇáÞÇÆãÉ ÈäÌÇÍ", vbInformation, "ÊÍÏíË"
End Sub

Private Sub CommandButton4_Click()
' 1. ÇáÊÃßÏ ãä Ãä ÇáãÓÊÎÏã ÇÎÊÇÑ ÊÔßíáÇð áÊÚÏíáå æÃä ÇáÊíßÓÊ ÈæßÓ áíÓ ÝÇÑÛÇð
    If ComboBox1.ListIndex = -1 Then
        MsgBox "íÑÌì ÇÎÊíÇÑ ÇáÊÔßíá ÇáãÑÇÏ ÊÍÏíËå ãä ÇáÞÇÆãÉ ÃæáÇð", vbExclamation
        Exit Sub
    End If
    
    If TextBox5.Text = "" Then
        MsgBox "íÑÌì ßÊÇÈÉ ÇáÇÓã ÇáÌÏíÏ Ýí ÇáÊíßÓÊ ÈæßÓ", vbExclamation
        Exit Sub
    End If

    ' 2. ÊÍÏíÏ ÑÞã ÇáÕÝ ÇáãÑÇÏ ÊÚÏíáå Ýí ÇáÔíÊ 5
    ' ÈãÇ Ãä ÇáÞÇÆãÉ ÊÈÏÃ ãä B1¡ ÝÅä ÑÞã ÇáÕÝ åæ Index + 1
    Dim TargetRow As Long
    TargetRow = ComboBox1.ListIndex + 1

    ' 3. ÊÍÏíË ÇáÞíãÉ Ýí ÇáÔíÊ 5 ãÈÇÔÑÉ
    Sheets(5).Cells(TargetRow, "B").value = TextBox5.Text

    ' 4. ÊÍÏíË ComboBox1 ÝæÑÇð áíÚßÓ ÇáÇÓã ÇáÌÏíÏ
    lastRow = Sheets(5).Cells(rowS.count, "B").End(xlUp).row
    If lastRow > 1 Then
        ComboBox1.List = Application.Transpose(Sheets(5).Range("B1:B" & lastRow).value)
    Else
        ComboBox1.List = Array(Sheets(5).Range("B1").value)
    End If

    ' 5. ãÓÍ ÇáÊíßÓÊ ÈæßÓ æÅÚáÇã ÇáãÓÊÎÏã
    TextBox5.value = ""
    MsgBox "Êã ÊÍÏíË ÇáÊÔßíá ÈäÌÇÍ", vbInformation
End Sub

Private Sub CommandButton5_Click()
Dim TargetRow As Long
    
    ' 1. ÇáÊÃßÏ ãä ÇÎÊíÇÑ ÊÔßíá ãä ÇáÞÇÆãÉ
    If ComboBox1.ListIndex = -1 Then
        MsgBox "íÑÌì ÇÎÊíÇÑ ÇáÊÔßíá ÇáãÑÇÏ ÍÐÝå ÃæáÇð", vbExclamation
        Exit Sub
    End If

    ' 2. ÊÍÏíÏ ÑÞã ÇáÕÝ (Index + 1 áÃä ÇáÈíÇäÇÊ ÊÈÏÃ ãä B1)
    TargetRow = ComboBox1.ListIndex + 1

    ' 3. ÍÐÝ ÇáÎáíÉ Ýí ÇáÚãæÏ B ÝÞØ æÅÒÇÍÉ ÇáÎáÇíÇ ÇáÊí ÊÍÊåÇ ááÃÚáì
    ' åÐÇ ÇáÓØÑ áÇ íÍÐÝ ÇáÓØÑ ÈÇáßÇãá¡ Èá íÍÐÝ ÇáÎáíÉ B ÝÞØ
    Sheets(5).Range("B" & TargetRow).Delete Shift:=xlUp

    ' 4. ÊÍÏíË ComboBox1 ÝæÑÇð áíÚßÓ ÇáÞÇÆãÉ ÇáÌÏíÏÉ
    lastRow = Sheets(5).Cells(rowS.count, "B").End(xlUp).row
    
    If lastRow >= 1 And Sheets(5).Range("B1").value <> "" Then
        ComboBox1.List = Application.Transpose(Sheets(5).Range("B1:B" & lastRow).value)
    Else
        ComboBox1.Clear
    End If

    MsgBox "Êã ÍÐÝ ÇáÞíãÉ æÊÍÏíË ÇáÞÇÆãÉ", vbInformation
End Sub

Private Sub UserForm_Initialize()
 Dim lastRow As Long
Dim ws1 As Worksheet

' ÊÚííä ÇáæÑÞÉ ÇáÎÇãÓÉ ááãÊÛíÑ ws1
Set ws1 = Sheets(5)

' ÊÍÏíÏ ÑÞã ÂÎÑ ÕÝ íÍÊæí Úáì ÈíÇäÇÊ Ýí ÇáÚãæÏ B
' ÇáßæÏ íÈÏÃ ãä ÂÎÑ ÕÝ Ýí ÇáÅßÓíá æíÕÚÏ ááÃÚáì ÍÊì íÌÏ Ãæá ÎáíÉ ÈåÇ ÞíãÉ
lastRow = ws1.Cells(ws1.rowS.count, "B").End(xlUp).row

' ÊÚÈÆÉ ComboBox4 ÈÇáäØÇÞ ãä ÇáÎáíÉ B1 Åáì ÂÎÑ ÕÝ æÌÏäÇå
If lastRow >= 1 Then
    ComboBox1.List = ws1.Range("B1:B" & lastRow).value
End If
End Sub

