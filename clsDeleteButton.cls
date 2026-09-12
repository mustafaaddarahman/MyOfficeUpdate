VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDeleteButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

' ÊÚÑíİ ÇáÃÍÏÇË æÇáãÑÌÚ ááíæÒÑ İæÑã
Public WithEvents DelLabel As MSForms.Label
Attribute DelLabel.VB_VarHelpID = -1
Public ParentForm As Object

Private Sub DelLabel_Click()
    Dim ctrl As MSForms.Control
    Dim chkTarget As MSForms.CheckBox
    Dim targetID As String
    Dim i As Long
    Dim hasDynamic As Boolean
    Dim maxTop As Double
    
    ' 1. ÊÍæíá ÇáÊÇÌ Åáì äÕ áÊÌäÈ ÎØÃ Type Mismatch (ÎÇÕÉ ÅĞÇ ßÇä gID ÑŞãÇğ)
    targetID = CStr(DelLabel.Tag)
    
    ' 2. ãÍÇæáÉ ÇáæÕæá ááÌíß ÈæßÓ ÇáãÑÊÈØ ÈåĞÇ ÇáŞíÏ
    On Error Resume Next
    Set chkTarget = ParentForm.Controls("chkDel_" & targetID)
    On Error GoTo 0
    
    ' ÅĞÇ áã íÊã ÇáÚËæÑ Úáì ÇáÌíß ÈæßÓ (ÑÈãÇ áã íäÔÃ ÈÚÏ Ãæ ÊÛíÑ ÇÓãå)
    If chkTarget Is Nothing Then
        MsgBox "ÎØÃ ÈÑãí: áã íÊã ÇáÚËæÑ Úáì ãÑÈÚ ÇáÇÎÊíÇÑ ÇáãÓãì " & "chkDel_" & targetID, vbCritical
        Exit Sub
    End If
    
    ' 3. ÇáİÍÕ ÇáĞßí: åá ãÑÈÚ ÇáÕÍ (CheckBox) ãİÚøá¿
    If chkTarget.value = True Then
        
        ' --- [ Ã: ÚãáíÉ ÇáÍĞİ ÇáÚßÓí áÌãíÚ ãßæäÇÊ ÇáŞíÏ ] ---
        For i = ParentForm.Controls.count - 1 To 0 Step -1
            ' ãŞÇÑäÉ ÇáÊÇÌ ßäÕæÕ áÖãÇä ÇáÏŞÉ
            If CStr(ParentForm.Controls(i).Tag) = targetID Then
                ParentForm.Controls.Remove ParentForm.Controls(i).Name
            End If
        Next i
        
        ' --- [ È: ÅÚÇÏÉ ÍÓÇÈ ÇáÇÑÊİÇÚ æÇáÚÏÇÏÇÊ ] ---
        hasDynamic = False
        maxTop = 120 ' ÇáŞãÉ ÇáÇİÊÑÇÖíÉ áÃæá ŞíÏ
        
        For Each ctrl In ParentForm.Controls
            ' İÍÕ ÇáßÇÆäÇÊ ÇáÊí áåÇ ÊÇÛ İŞØ (ÇÓÊÈÚÇÏ ÇáÃÏæÇÊ ÇáËÇÈÊÉ)
            If CStr(ctrl.Tag) <> "" Then
                hasDynamic = True
                ' ÇáÈÍË Úä ÃÎÑ äŞØÉ (Top) æÕá ÅáíåÇ ÃÎÑ ÚäÕÑ ãæÌæÏ
                If ctrl.Top > maxTop Then maxTop = ctrl.Top
            End If
        Next ctrl
        
        ' --- [ Ì: ÊÍÏíË ãÊÛíÑÇÊ ÇáíæÒÑ İæÑã ÇáÚÇãÉ ] ---
        ' ãáÇÍÙÉ: íÌÈ Ãä íßæä NextTop æ GroupID æ DelButtons ãÚÑøİíä ßÜ Public İí ÇáİæÑã
        If Not hasDynamic Then
            ' ÅĞÇ ÃÕÈÍ ÇáİæÑã İÇÑÛÇğ ÊãÇãÇğ
            ParentForm.NextTop = 120
            ParentForm.GroupID = 0
            ' ÊÕİíÑ ÇáãÌãæÚÉ (Collection) áãÓÍ ÇáãİÇÊíÍ ÇáŞÏíãÉ ãä ÇáĞÇßÑÉ
            Set ParentForm.DelButtons = New Collection
            ParentForm.ScrollTop = 0
            ParentForm.ScrollHeight = ParentForm.InsideHeight
            Application.StatusBar = "Êã ÍĞİ ÇáŞíÏ æÊÕİíÑ ÇáäÙÇã ÈÇáßÇãá."
        Else
            ' ÅĞÇ ÈŞíÊ ŞíæÏ ÃÎÑì¡ äÌÚá ãæŞÚ ÇáÅÖÇİÉ ÇáŞÇÏã ÈÚÏ ÃÎÑ ÚäÕÑ ÈãÓÇİÉ 90
            ParentForm.NextTop = maxTop + 90
            ParentForm.ScrollHeight = ParentForm.NextTop + 50
            Application.StatusBar = "Êã ÍĞİ ÇáŞíÏ ÈäÌÇÍ."
        End If
        
        ' ÊÍÏíË ÇáÔÇÔÉ İæÑÇğ
        ParentForm.Repaint
        
    Else
        ' ÊäÈíå ÇáãÓÊÎÏã İí ÍÇá ÖÛØ ÅáÛÇÁ ÈÏæä ÊİÚíá ÇáÕÍ
        MsgBox "ÃÓÊÇĞ ÓÚÏ íÑÌì ÊİÚíá ãÑÈÚ (ÇáÕÍ) ÃæáÇğ áÊÃßíÏ ÍĞİ åĞÇ ÇáãæÙİ.", vbInformation, "ÊÃßíÏ ÇáÅÌÑÇÁ"
    End If
End Sub
