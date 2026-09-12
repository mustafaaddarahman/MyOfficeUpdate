VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm82 
   Caption         =   "UserForm82"
   ClientHeight    =   9384.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7044
   OleObjectBlob   =   "UserForm82.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm82"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GetLineByHeader(HeaderName As String) As String
  
  Dim ws As Worksheet
    Dim ColIndex As Long
    
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ÇáÈÍË Úä ÑŞã ÇáÚãæÏ ÈäÇÁğ Úáì ÇáåíÏÑ
    On Error Resume Next
    ColIndex = ws.rowS(8).Find(What:=HeaderName, LookAt:=xlWhole).Column
    On Error GoTo 0
    
    If ColIndex > 0 Then
        ' ÌáÈ ÇáãÚÇÏáÉ ãä ÇáÕİ 9 İí Ğáß ÇáÚãæÏ
        GetFormulaFromSheet = ws.Cells(9, ColIndex).Formula
    Else
        GetFormulaFromSheet = "áã íÊã ÇáÚËæÑ Úáì ÇáÚãæÏ!"
    End If
  
  
  
  
  
  
End Function
Private Sub CommandButton1_Click()
Dim ws As Worksheet
    Dim FoundCell As Range
    Dim lastRow As Long
    Dim ColNum As Long
    Sheets(4).Cells(2, "R").value = TextBox3.value
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' 1. ÇáÚËæÑ Úáì ÇáÚãæÏ ÈäÇÁğ Úáì ãÇ ÇÎÊÑÊå İí ComboBox1
    Set FoundCell = ws.rowS(8).Find(What:=ComboBox1.value, LookAt:=xlWhole)
    
    If FoundCell Is Nothing Then
        MsgBox "áã ÃÌÏ ÇáÚãæÏ ÇáãÓãì: " & ComboBox1.value
        Exit Sub
    End If
    
    ColNum = FoundCell.Column
    
    ' 2. ÊÛííÑ ÇÓã ÇáåíÏÑ (ÇáÕİ 8) Åáì ãÇ åæ ãßÊæÈ İí TextBox1
    FoundCell.value = TextBox1.value
    
    ' 3. ÊÍÏíÏ ÂÎÑ Õİ ÈäÇÁğ Úáì ÇáÚãæÏ B
    lastRow = ws.Cells(ws.rowS.count, 2).End(xlUp).row
    
    
    ' 4. æÖÚ ÇáÕíÛÉ ãä TextBox2 İí ÇáÎáíÉ 9 æİí ßá ÇáÚãæÏ
    ' ãáÇÍÙÉ: äÍä äÃÎĞ ÇáäÕ ãä TextBox2 ßãÇ åæ æäÖÚå İí ÇáÎáÇíÇ
    ws.Range(ws.Cells(9, ColNum), ws.Cells(lastRow, ColNum)).Formula = TextBox2.value
    
    MsgBox "ÊãÊ ÇáÚãáíÉ: Êã ÊÍÏíË ÇáåíÏÑ æÊÚÈÆÉ ÇáÕíÛÉ ãä ÇáÕİ 9 Åáì " & lastRow
   Call count_all
End Sub

Private Sub CommandButton2_Click()
Unload Me


End Sub

Private Sub UserForm_Initialize()
    Dim i As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ComboBox1.Clear
    
    ' äÈÍË İí ÇáÕİ 8
    For i = 1 To 200
        ' ÔÑØ ãÒÏæÌ:
        ' 1. ÇáÎáíÉ áíÓÊ İÇÑÛÉ
        ' 2. ÇáãÍÊæì åæ äÕ (vbString)
        If Len(Trim(ws.Cells(8, i).value)) > 0 Then
            If VarType(ws.Cells(8, i).value) = vbString Then
                ComboBox1.AddItem ws.Cells(8, i).value
            End If
        End If
    Next i
End Sub

' ÚäÏ ÇÎÊíÇÑ ÚãæÏ¡ ÌáÈ ÇáãÚÇÏáÉ ÇáÍÇáíÉ æÇáÇÓã ÇáÍÇáí
Private Sub ComboBox1_Change()
   Dim ws As Worksheet
    Dim ColIndex As Long
    Dim HeaderName As String
    
    If ComboBox1.value = "" Then Exit Sub
    
    HeaderName = ComboBox1.value
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' ÇáÈÍË Úä ÑŞã ÇáÚãæÏ ÈäÇÁğ Úáì ÇáåíÏÑ ÇáãæÌæÏ İí ÇáÕİ 8
    On Error Resume Next
    ColIndex = ws.rowS(8).Find(What:=HeaderName, LookAt:=xlWhole).Column
    On Error GoTo 0
    
    If ColIndex > 0 Then
        ' ÌáÈ ÕíÛÉ ÇáÏÇáÉ ãä ÇáÕİ 9 İí åĞÇ ÇáÚãæÏ ææÖÚåÇ İí TextBox2
        ' äÓÊÎÏã Formula áÖãÇä ÇáÍÕæá Úáì ÇáãÚÇÏáÉ æáíÓ ÇáäÊíÌÉ
        TextBox2.value = ws.Cells(9, ColIndex).Formula
    Else
        MsgBox "áã íÊã ÇáÚËæÑ Úáì ÈíÇäÇÊ İí ÇáÕİ 9 áåĞÇ ÇáÚãæÏ!", vbExclamation
    End If
End Sub
Public Sub UpdateModuleFormula(HeaderName As String, NewFormula As String)
    Dim VBProj As Object, VBComp As Object, CodeMod As Object
    Dim i As Long
    
    ' ÇáæÕæá Åáì ÇáãæÏíæá
    Set VBProj = ThisWorkbook.VBProject
    Set VBComp = VBProj.VBComponents("count_all_salalry")
    Set CodeMod = VBComp.CodeModule
    
    ' ÇáÈÍË Úä ÇáÓØÑ ÇáÈÑãÌí ÇáãÑÊÈØ ÈÇÓã ÇáÚãæÏ
    For i = 1 To CodeMod.CountOfLines
        ' äÈÍË Úä ÇÓã ÇáÚãæÏ ÏÇÎá ÇáãæÏíæá áÖãÇä ÃääÇ äÍÏË ÇáÓØÑ ÇáÕÍíÍ
        If InStr(1, CodeMod.Lines(i, 1), HeaderName, vbTextCompare) > 0 Then
            ' äŞæã ÈÊÍÏíË ÇáÓØÑ áíÍãá ÇáãÚÇÏáÉ ÇáÌÏíÏÉ (ÈİÇÕáÉ ãäŞæØÉ)
            ' äÖÚåÇ ßÊÚáíŞ Ãæ ßãÊÛíÑ áÊÊãßä ãä ÇÓÊÑÌÇÚåÇ áÇÍŞÇğ
            CodeMod.ReplaceLine i, "' " & HeaderName & " Formula: " & NewFormula
            Exit Sub
        End If
    Next i
    
    ' ÅĞÇ áã äÌÏ ÇáÓØÑ¡ äŞæã ÈÅÖÇİÊå İí äåÇíÉ ÇáãæÏíæá
    CodeMod.InsertLines CodeMod.CountOfLines + 1, "' " & HeaderName & " Formula: " & NewFormula
End Sub
