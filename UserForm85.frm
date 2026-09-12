VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm85 
   Caption         =   "UserForm85"
   ClientHeight    =   3852
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7956
   OleObjectBlob   =   "UserForm85.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm85"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub CommandButton2_Click()



End Sub

Private Sub CommandButton3_Click()




End Sub

Private Sub CommandButton4_Click()
Unload Me
End Sub

Private Sub CommandButton5_Click()
Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Dim lastRow As Long
    Dim i As Long
    Dim valToSpread As Double
    Dim ColDataArr As Variant
    
    ' 1. ŞÑÇÁÉ ÇáŞíãÉ ãä TextBox52
    valToSpread = val(Me.TextBox52.Text)
    
    ' 2. ãÚÑİÉ ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ İí ÇáÔíÊ
    lastRow = ws.Cells(ws.rowS.count, "aw").End(xlUp).row
    
    ' ÅíŞÇİ ÊÍÏíË ÇáÔÇÔÉ áÊÓÑíÚ ÇáÃÏÇÁ
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    ' 3. ÌáÈ äØÇŞ ÇáÚãæÏ al ÈÇáßÇãá Åáì ÇáĞÇßÑÉ ÏİÚÉ æÇÍÏÉ
    ColDataArr = ws.Range("ck9:ck" & lastRow).value
    
    ' ÊÚãíã ÇáŞíãÉ ÏÇÎá ÇáãÕİæİÉ
    For i = 1 To UBound(ColDataArr, 1)
        ColDataArr(i, 1) = valToSpread
    Next i
    
    ' ßÊÇÈÉ ÇáãÕİæİÉ ÇáãÍÏËÉ Åáì ÇáÔíÊ ÏİÚÉ æÇÍÏÉ
    ws.Range("ck9:ck" & lastRow).value = ColDataArr
    
    ' ÅÚÇÏÉ ÅİÚíá ÅÚÏÇÏÇÊ ÇáÅßÓá
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
   
   
   Call RecalculateAnWithEntitlements
   
    
   
    MsgBox "Êã ÊÚãíã ÇáŞíãÉ Úáì ÚãæÏ al æÅÚÇÏÉ ÇáÇÍÊÓÇÈ ÈäÌÇÍ", vbInformation, "ÊÃßíÏ"



End Sub

Private Sub Label1_Click()

End Sub

Private Sub TextBox52_Change()

Call englsh



End Sub

Private Sub UserForm_Activate()
 
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
 TextBox53.value = Sheets(1).Cells(9, "ck").value
 
 
End Sub
Sub ClearSheet4Range()
    Dim ws As Object
    Dim lastRow As Long
    
    ' ÊÍÏíÏ ÇáÔíÊ ÇáÑÇÈÚ
    Set ws = ThisWorkbook.Sheets("Sheet4")
    
    ' ÅíŞÇİ ÊÍÏíË ÇáÔÇÔÉ áÊÓÑíÚ ÇáÊäİíĞ
    Application.ScreenUpdating = False
    
    ' ÇáÈÍË Úä ÂÎÑ Õİ íÍÊæí Úáì ÈíÇäÇÊ Öãä ÇáÃÚãÏÉ ãä S Åáì V
    On Error Resume Next
    lastRow = ws.Columns("S:W").Find(What:="*", _
                                    After:=ws.Range("S2"), _
                                    LookAt:=xlPart, _
                                    LookIn:=xlFormulas, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious).row
    On Error GoTo 0
    
    ' ÇáÊÍŞŞ ããÇ ÅĞÇ ßÇä åäÇß Õİæİ ÈíÇäÇÊ ÃÕáÇğ ÈÚÏ ÇáÕİ ÇáÃæá (S2)
    If lastRow >= 2 Then
        ' ãÓÍ ÇáãÍÊæíÇÊ ãä ÇáÕİ 2 Åáì ÂÎÑ Õİ Êã ÇáÚËæÑ Úáíå ááäØÇŞ S:V
        ws.Range("S2:V" & lastRow).ClearContents
           Else
        MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ ááãÓÍ İí ÇáäØÇŞ ÇáãÍÏÏ.", vbExclamation, "ÊäÈíå"
    End If
    
    ' ÅÚÇÏÉ ÊÔÛíá ÊÍÏíË ÇáÔÇÔÉ
    Application.ScreenUpdating = True
End Sub



