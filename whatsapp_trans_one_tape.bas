Attribute VB_Name = "whatsapp_trans_one_tape"
Sub whatsapp_trans_one_tape_1()
Dim contact As String
Dim upsum As Double
Dim lastR As Integer
Dim frow As Integer
ThisWorkbook.FollowHyperlink "https://web.whatsapp.com/"
Application.Wait (Now + TimeValue("00:00:10"))
lastR = Sheets(1).Cells(Rows.count, "e").End(xlUp).row

'lastr = 238
For frow = UserForm42.TextBox3.value To UserForm42.TextBox3.value

'If UserForm42.TextBox1.Text = contact And UserForm42.TextBox2.Text Like "*" & Sheets(1).Cells(frow, "e") & "*" Then

Sheets("whatsapp").Cells(1, "b") = Sheets(1).Cells(frow, "e")
Sheets("whatsapp").Cells(1, "e") = Sheets(1).Cells(frow, "g")
Sheets("whatsapp").Cells(1, "g") = Sheets("whatsapp").Cells(1, "p")
Sheets("whatsapp").Cells(1, "j") = Sheets(1).Cells(frow, "be")
Sheets("whatsapp").Cells(1, "k") = Sheets(1).Cells(frow, "az")
Sheets("whatsapp").Cells(2, "e") = Sheets(1).Cells(frow, "b")
Sheets("whatsapp").Cells(2, "h") = Sheets(1).Cells(frow, "f")
Sheets("whatsapp").Cells(2, "j") = Sheets(1).Cells(frow, "h")
Sheets("whatsapp").Cells(2, "l") = Sheets(1).Cells(frow, "i")
Sheets("whatsapp").Cells(3, "l") = Sheets(1).Cells(frow, "z")
Sheets("whatsapp").Cells(3, "j") = Sheets(1).Cells(frow, "ac")
Sheets("whatsapp").Cells(3, "h") = Sheets(1).Cells(frow, "n")
Sheets("whatsapp").Cells(3, "f") = Sheets(1).Cells(frow, "p")
Sheets("whatsapp").Cells(3, "d") = Sheets(1).Cells(frow, "r")
Sheets("whatsapp").Cells(3, "b") = Sheets(1).Cells(frow, "j")
Sheets("whatsapp").Cells(4, "b") = Sheets(1).Cells(frow, "aa")
Sheets("whatsapp").Cells(4, "d") = Sheets(1).Cells(frow, "t")
Sheets("whatsapp").Cells(4, "f") = Sheets(1).Cells(frow, "u")
Sheets("whatsapp").Cells(4, "h") = Sheets(1).Cells(frow, "s")
Sheets("whatsapp").Cells(4, "j") = Sheets(1).Cells(frow, "ad")
Sheets("whatsapp").Cells(4, "l") = Sheets(1).Cells(frow, "x")
Sheets("whatsapp").Cells(5, "l") = Sheets(1).Cells(frow, "ao")
Sheets("whatsapp").Cells(5, "j") = Sheets(1).Cells(frow, "an")
Sheets("whatsapp").Cells(5, "h") = Sheets(1).Cells(frow, "ai")
Sheets("whatsapp").Cells(5, "f") = Sheets(1).Cells(frow, "au")
Sheets("whatsapp").Cells(5, "d") = Sheets(1).Cells(frow, "k")
Sheets("whatsapp").Cells(5, "b") = Sheets(1).Cells(frow, "aj")
Sheets("whatsapp").Cells(6, "b") = Sheets(1).Cells(frow, "ar")
Sheets("whatsapp").Cells(6, "d") = Sheets(1).Cells(frow, "at")
Sheets("whatsapp").Cells(6, "f") = Sheets(1).Cells(frow, "af")
Sheets("whatsapp").Cells(6, "h") = Sheets(1).Cells(frow, "ap")
Sheets("whatsapp").Cells(6, "j") = Sheets(1).Cells(frow, "ah")
Sheets("whatsapp").Cells(6, "l") = Sheets(1).Cells(frow, "bh")
Sheets("whatsapp").Cells(7, "d") = Sheets(1).Cells(frow, "bj")
Sheets("whatsapp").Cells(7, "f") = Sheets(1).Cells(frow, "av")
Sheets("whatsapp").Cells(7, "h") = Sheets(1).Cells(frow, "aw")
Sheets("whatsapp").Cells(7, "j") = Sheets(1).Cells(frow, "cc")
contact = Sheets(1).Cells(frow, "cc")


contact = Sheets(1).Cells(frow, "cc")
'If contact = "000" Then
'GoTo 1
'ElseIf frow = 20 Then
'MsgBox " „ «—”«· Ã„Ì⁄ «‘—ÿ… «·—« » »‰Ã«Õ", vbOKOnly + vbInformation + vbCritical, "—”«·…  ‰»ÌÂ"
'Exit Sub

'Else
Call SendKeys("{TAB}", True)
Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("{TAB}", True)
Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("{TAB}", True)
Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("{TAB}", True)
Application.Wait (Now + TimeValue("00:00:01"))




Call SendKeys(contact, True)
Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("~", True)
Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("ﬁ«„ ‰Ÿ«„ «·—Ê« » »«—”«· «‘—ÿ… «·—Ê« » ·Â–« «·‘Â— „⁄  ÕÌ«  ‘⁄»… «·„⁄·Ê„«  „’‰⁄ «·ﬂ—«„…", True)
Application.Wait (Now + TimeValue("00:00:01"))

Sheets("whatsapp").Range("a1:l7").Copy
Application.Wait (Now + TimeValue("00:00:01"))
Application.SendKeys "^v"

Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("~", True)
Application.Wait (Now + TimeValue("00:00:01"))
Call SendKeys("{TAB}", True)
Application.Wait (Now + TimeValue("00:00:01"))
'End If
Next

MsgBox " „ «—”«· Ã„Ì⁄ «‘—ÿ… «·—« » »‰Ã«Õ", vbOKOnly + vbInformation + vbCritical, "—”«·…  ‰»ÌÂ"



End Sub




