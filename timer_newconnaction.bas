Attribute VB_Name = "timer_newconnaction"
Option Explicit
Public NextTick As Date

' ÇáãÇßÑæ ÇáĞí íÈÍË Úäå ÇáÅßÓíá ßá 10 ËæÇäò
Sub TriggerTimer()
    On Error Resume Next
    ' ÇÓÊÏÚÇÁ ãÍÑß ÇáÊÍÏíË ÏÇÎá ÇáİæÑã (ÊÃßÏ Ãä ÇáÇÓã UserForm70)
    Call UserForm70.Auto_Pulse_Engine
End Sub
