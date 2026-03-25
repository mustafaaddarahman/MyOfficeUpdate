Attribute VB_Name = "change_tow_languge"
#If Win64 Then
Private Declare PtrSafe Function Keyboard Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal ss As String, ByVal sss As Long) As LongPtr
#Else
Private Declare Function Keyboard Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal ss As String, ByVal sss As Long) As Long
#End If

Sub arb()
Call Keyboard("00000401", 1)
End Sub

Sub englsh()
Call Keyboard("00000409", 1)
End Sub
