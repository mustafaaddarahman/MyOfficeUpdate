Attribute VB_Name = "marge_down_form32"
Sub marge_down_form32_1()
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
UserForm32.TextBox4.value = Sheets(39).Cells(Rows.count, 1).End(xlUp).row + 4

'====================================================
Sheets(39).Range("a9  : " & "br" & UserForm32.TextBox4.value + 3).UnMerge
'Range("a9  : " & "br" & UserForm32.TextBox4.Value + 3).

'Sheets(39).Range("a9  : " & "br" & UserForm32.TextBox4.Value + 3).
'=================================
Sheets(39).Range("a" & UserForm32.TextBox4.value + 3 & " : " & "br" & UserForm32.TextBox4.value).Merge

'Sheets(39).Range("f" & UserForm32.TextBox4.Value + 3 & " : " & "g" & UserForm32.TextBox4.Value).Merge
'Sheets(39).Range("i" & UserForm32.TextBox4.Value + 3 & " : " & "k" & UserForm32.TextBox4.Value).Merge
'Sheets(39).Range("m" & UserForm32.TextBox4.Value + 3 & " : " & "o" & UserForm32.TextBox4.Value).Merge
'Sheets(39).Range("q" & UserForm32.TextBox4.Value + 3 & " : " & "s" & UserForm32.TextBox4.Value).Merge





al_munatham = " «·„‰Ÿ„"
al_muhaseb = "«·„Õ«”»"
al_mudakek = "«·„œﬁﬁ "
al_mudeerhesabat = "„œÌ— «· œﬁÌﬁ"
al_mudeermalia = "„œÌ— «·„«·Ì  "
al_mudeermassnaa = "„œÌ— «·„’‰⁄"
Sheets(39).Range("a" & UserForm32.TextBox4.value + 3 & " : " & "br" & UserForm32.TextBox4.value).value = al_munatham & "                            " & al_muhaseb _
& "                         " & al_mudakek & "                         " & al_mudeerhesabat _
& "                            " & al_mudeermassnaa
Sheets(39).Range("a" & UserForm32.TextBox4.value + 3 & " : " & "br" & UserForm32.TextBox4.value).HorizontalAlignment = xlCenter

End Sub

