VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   3888
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7896
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
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
    
    ' 1. чягаи гАчМЦи ЦД TextBox52
    valToSpread = val(Me.TextBox52.Text)
    
    ' 2. Цзящи бня ущ МмйФМ зАЛ хМгДгй щМ гАтМй
    lastRow = ws.Cells(ws.rowS.count, "aw").End(xlUp).row
    
    ' еМчгщ ймоМк гАтгти АйсяМз гАцога
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
    End With
    
    ' 3. лАх Дьгч гАзЦФо al хгАъгЦА еАЛ гАпгъяи ощзи Фгмои
    ColDataArr = ws.Range("al9:al" & lastRow).value
    
    ' йзЦМЦ гАчМЦи огнА гАЦущФщи
    For i = 1 To UBound(ColDataArr, 1)
        ColDataArr(i, 1) = valToSpread
    Next i
    
    ' ъйгхи гАЦущФщи гАЦмоки еАЛ гАтМй ощзи Фгмои
    ws.Range("al9:al" & lastRow).value = ColDataArr
    
    ' езгои ещзМА езогогй гАеъсА
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With
    
   
   
   Call RecalculateAnWithEntitlements
   
    
   
    MsgBox "йЦ йзЦМЦ гАчМЦи зАЛ зЦФо al Фезгои гАгмйсгх хДлгм", vbInformation, "йцъМо"


End Sub



Private Sub CommandButton6_Click()

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
 TextBox53.value = Sheets(1).Cells(9, "al").value
 
 
End Sub
Sub ClearSheet4Range()
    Dim ws As Object
    Dim lastRow As Long
    
    ' ймоМо гАтМй гАягхз
    Set ws = ThisWorkbook.Sheets("Sheet4")
    
    ' еМчгщ ймоМк гАтгти АйсяМз гАйДщМп
    Application.ScreenUpdating = False
    
    ' гАхмк зД бня ущ МмйФМ зАЛ хМгДгй жЦД гАцзЦои ЦД S еАЛ V
    On Error Resume Next
    lastRow = ws.Columns("S:W").Find(What:="*", _
                                    After:=ws.Range("S2"), _
                                    LookAt:=xlPart, _
                                    LookIn:=xlFormulas, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious).row
    On Error GoTo 0
    
    ' гАймчч ЦЦг епг ъгД ЕДгъ ущФщ хМгДгй цуАгП хзо гАущ гАцФА (S2)
    If lastRow >= 2 Then
        ' Цсм гАЦмйФМгй ЦД гАущ 2 еАЛ бня ущ йЦ гАзкФя зАМЕ ААДьгч S:V
        ws.Range("S2:V" & lastRow).ClearContents
        MsgBox "йЦ Цсм гАДьгч хДлгм мйЛ гАущ: " & lastRow, vbInformation, "йЦй гАзЦАМи"
    Else
        MsgBox "Аг йФло хМгДгй ААЦсм щМ гАДьгч гАЦмоо.", vbExclamation, "йДхМЕ"
    End If
    
    ' езгои йтшМА ймоМк гАтгти
    Application.ScreenUpdating = True
End Sub

