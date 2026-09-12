VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsZoomManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

' ÊÚÑíİ ÇáÃÏæÇÊ áÑÈØ ÃÍÏÇËåÇ (WithEvents)
Public WithEvents SpinButton1 As MSForms.SpinButton
Attribute SpinButton1.VB_VarHelpID = -1
Private TargetForm As Object
Private az As MSForms.Label

' ãÊÛíÑÇÊ ÍİÙ ÃÈÚÇÏ ÇáÊÕãíã ÇáÃÓÇÓíÉ
Private BaseWidth As Double
Private BaseHeight As Double
Private BaseZoom As Integer

' æÙíİÉ ÇáÑÈØ: ÊÑÈØ ÇáßÇÆäÇÊ ÈÈÚÖåÇ æÊÓÌá ÇáÃÈÚÇÏ ÇáÃÕáíÉ
Public Sub Bind(Frm As Object, SBtn As MSForms.SpinButton, Lbl As MSForms.Label)
    Set TargetForm = Frm
    Set SpinButton1 = SBtn
    Set az = Lbl
    
    ' ÊÓÌíá ÇáŞíã ÇáÃÕáíÉ ÚäÏ áÍÙÉ ÇáÊÔÛíá
    If TargetForm.Zoom = 0 Then TargetForm.Zoom = 100 ' ÊÃãíä ÖÏ ÇáŞíã ÇáÕİÑíÉ
    
    BaseZoom = TargetForm.Zoom
    BaseWidth = TargetForm.Width
    BaseHeight = TargetForm.Height
    
    ' ÅÚÏÇÏÇÊ ÃÏÇÉ ÇáÜ SpinButton1
    With SpinButton1
        .Min = 10
        .Max = 400
        .value = BaseZoom
        .SmallChange = 5
    End With
    
    ' ÊÍÏíË ÇáäÕ ÇáÇÈÊÏÇÆí ááíÈá az
    az.Caption = "Zoom: " & BaseZoom & "%"
End Sub

' ãÍÑß ÇáÒææã: íäİĞ ÇáÊÛííÑ ÚäÏ ÇáÖÛØ Úáì ÃÓåã ÇáÓÈíä-ÈæÊæä
Private Sub SpinButton1_Change()
    On Error Resume Next
    Dim NewZoom As Integer
    NewZoom = SpinButton1.value
    
    ' 1. ÊØÈíŞ ÇáÒææã Úáì ãÍÊæíÇÊ ÇáİæÑã
    TargetForm.Zoom = NewZoom
    
    ' 2. ÊÚÏíá ÍÌã ÅØÇÑ ÇáİæÑã áíÊäÇÓÈ ãÚ ÇáÒææã ÇáÌÏíÏ (äÓÈÉ æÊäÇÓÈ)
    TargetForm.Width = BaseWidth * (NewZoom / BaseZoom)
    TargetForm.Height = BaseHeight * (NewZoom / BaseZoom)
    
    ' 3. ÊÍÏíË ŞíãÉ ÇááíÈá az
    az.Caption = "Zoom: " & NewZoom & "%"
End Sub
