Attribute VB_Name = "marge_down_form17"
Sub marge_down_form17_1()
On Error Resume Next
Dim al_munatham As String
Dim al_muhaseb As String
Dim al_mudakek As String
Dim al_mudeerhesabat As String
Dim al_mudeermalia As String
Dim al_mudeermassnaa As String


al_munatham = ""
al_muhaseb = ""
al_mudakek = ""
al_mudeerhesabat = ""
al_mudeermalia = ""
al_mudeermassnaa = ""
UserForm17.TextBox4.value = Sheets(28).Cells(Rows.count, 1).End(xlUp).row + 4

'====================================================
Sheets(28).Range("a9  : " & "bz" & UserForm17.TextBox4.value + 3).UnMerge
'Range("a9  : " & "br" & UserForm17.TextBox4.Value + 3).

'Sheets(28).Range("a9  : " & "br" & UserForm17.TextBox4.Value + 3).
'=================================
Sheets(28).Range("a" & UserForm17.TextBox4.value + 3 & " : " & "bz" & UserForm17.TextBox4.value).Merge

'Sheets(28).Range("f" & UserForm17.TextBox4.Value + 3 & " : " & "g" & UserForm17.TextBox4.Value).Merge
'Sheets(28).Range("i" & UserForm17.TextBox4.Value + 3 & " : " & "k" & UserForm17.TextBox4.Value).Merge
'Sheets(28).Range("m" & UserForm17.TextBox4.Value + 3 & " : " & "o" & UserForm17.TextBox4.Value).Merge
'Sheets(28).Range("q" & UserForm17.TextBox4.Value + 3 & " : " & "s" & UserForm17.TextBox4.Value).Merge





al_munatham = " «·„‰Ÿ„"
al_muhaseb = "«·„Õ«”»"
al_mudakek = "«·„œﬁﬁ "
al_mudeerhesabat = "„œÌ— «· œﬁÌﬁ"
al_mudeermalia = "„œÌ— «·„«·Ì  "
al_mudeermassnaa = "„œÌ— «·„’‰⁄"
Sheets(28).Range("a" & UserForm17.TextBox4.value + 3 & " : " & "bz" & UserForm17.TextBox4.value).value = al_munatham & "                            " & al_muhaseb _
& "                         " & al_mudakek & "                         " & al_mudeerhesabat _
& "                            " & al_mudeermassnaa
Sheets(28).Range("a" & UserForm17.TextBox4.value + 3 & " : " & "bz" & UserForm17.TextBox4.value).HorizontalAlignment = xlCenter
End Sub


